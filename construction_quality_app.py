from __future__ import annotations

import base64
from io import BytesIO
import json
import mimetypes
import os
import re
from dataclasses import dataclass
from typing import Any

from dotenv import load_dotenv
from openai import OpenAI
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import streamlit as st

load_dotenv()

APP_TITLE = "Construction Quality Inspector"
DEFAULT_MODEL = "gpt-5.1"
MODEL_OPTIONS = ["gpt-5.1", "gpt-5-mini", "gpt-4.1"]
QUALITY_LEVELS = ["Excellent", "Good", "Fair", "Poor", "Critical"]
CONFIDENCE_LEVELS = ["High", "Medium", "Low"]
WORK_TYPE_OPTIONS = [
    "Structural concrete",
    "Masonry",
    "Plastering",
    "Flooring",
    "Waterproofing",
    "Painting",
    "MEP installation",
    "Facade work",
    "Finishing work",
    "Housekeeping and safety",
    "Other",
]
SYSTEM_PROMPT = """
You are a construction quality inspector assistant.
Assess visible workmanship, quality issues, safety concerns, finish consistency,
material condition, and signs of poor execution from a single construction-site image.

Rules:
- Base your judgment only on what is visible in the image.
- Do not claim compliance with codes, structural integrity, or hidden defects unless you clearly state that they cannot be verified visually.
- Keep the report practical for a site manager.
- Return strict JSON only.
- If image quality is limited, lower confidence and mention the limitation.
- Use concise field values.
""".strip()


@dataclass
class InspectionResult:
    image_name: str
    project_name: str
    work_type: str
    overall_score: int
    quality_level: str
    summary: str
    observations: list[str]
    risks: list[str]
    recommended_actions: list[str]
    confidence: str
    limitations: list[str]

    def normalized(self) -> dict[str, Any]:
        return {
            "image_name": self.image_name,
            "project_name": self.project_name,
            "work_type": self.work_type,
            "overall_score": self.overall_score,
            "quality_level": self.quality_level,
            "summary": self.summary,
            "observations": self.observations,
            "risks": self.risks,
            "recommended_actions": self.recommended_actions,
            "confidence": self.confidence,
            "limitations": self.limitations,
        }


def default_result(image_name: str, project_name: str, work_type: str, message: str) -> InspectionResult:
    return InspectionResult(
        image_name=image_name,
        project_name=project_name,
        work_type=work_type,
        overall_score=0,
        quality_level="Critical",
        summary=message,
        observations=[message],
        risks=["Analysis did not complete successfully."],
        recommended_actions=["Retry the upload or review the API configuration."],
        confidence="Low",
        limitations=["No AI assessment was produced."],
    )


def clean_list(value: Any, fallback: str, limit: int = 4) -> list[str]:
    if isinstance(value, list):
        items = [str(item).strip() for item in value if str(item).strip()]
        if items:
            return items[:limit]
    return [fallback]


def clamp_score(value: Any) -> int:
    try:
        return max(0, min(100, int(round(float(value)))))
    except (TypeError, ValueError):
        return 0


def normalize_quality_level(value: Any, score: int) -> str:
    text = str(value or "").strip().title()
    if text in QUALITY_LEVELS:
        return text
    if score >= 85:
        return "Excellent"
    if score >= 70:
        return "Good"
    if score >= 55:
        return "Fair"
    if score >= 35:
        return "Poor"
    return "Critical"


def normalize_confidence(value: Any) -> str:
    text = str(value or "").strip().title()
    if text in CONFIDENCE_LEVELS:
        return text
    return "Medium"


def sanitize_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    return cleaned.strip("._") or "report"


def parse_json_response(raw_text: str) -> dict[str, Any]:
    raw_text = raw_text.strip()
    if not raw_text:
        raise ValueError("The model returned an empty response.")

    try:
        return json.loads(raw_text)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", raw_text, re.DOTALL)
        if not match:
            raise ValueError("The model response was not valid JSON.")
        return json.loads(match.group(0))


def normalize_result(
    image_name: str,
    project_name: str,
    work_type: str,
    payload: dict[str, Any],
) -> InspectionResult:
    score = clamp_score(payload.get("overall_score"))
    return InspectionResult(
        image_name=image_name,
        project_name=project_name.strip() or "Untitled project",
        work_type=work_type.strip() or "General construction",
        overall_score=score,
        quality_level=normalize_quality_level(payload.get("quality_level"), score),
        summary=str(payload.get("summary") or "No summary returned.").strip(),
        observations=clean_list(payload.get("observations"), "No observations returned."),
        risks=clean_list(payload.get("risks"), "No specific risks identified.", limit=3),
        recommended_actions=clean_list(
            payload.get("recommended_actions"),
            "Capture clearer images and review the area on site.",
            limit=3,
        ),
        confidence=normalize_confidence(payload.get("confidence")),
        limitations=clean_list(
            payload.get("limitations"),
            "This review is based only on visible conditions in the photo.",
            limit=2,
        ),
    )


def encode_image(uploaded_file: Any) -> str:
    image_bytes = uploaded_file.getvalue()
    mime_type = uploaded_file.type or mimetypes.guess_type(uploaded_file.name)[0] or "image/jpeg"
    encoded = base64.b64encode(image_bytes).decode("ascii")
    return f"data:{mime_type};base64,{encoded}"


def build_image_prompt(image_name: str, project_name: str, work_type: str, inspection_brief: str) -> str:
    extra_brief = inspection_brief.strip() or "Focus on overall workmanship, finish quality, and visible defects."
    schema = {
        "image_name": image_name,
        "project_name": project_name,
        "work_type": work_type,
        "overall_score": "integer from 0 to 100",
        "quality_level": "Excellent | Good | Fair | Poor | Critical",
        "summary": "one short paragraph under 50 words",
        "observations": ["up to 4 short bullet-style strings"],
        "risks": ["up to 3 short bullet-style strings"],
        "recommended_actions": ["up to 3 short bullet-style strings"],
        "confidence": "High | Medium | Low",
        "limitations": ["1 or 2 short strings"],
    }
    return (
        f"Inspect the construction-site image named '{image_name}'. "
        f"Project name: {project_name}. "
        f"Type of work: {work_type}. "
        f"Additional inspection brief: {extra_brief}\n\n"
        "Return a strict JSON object that follows this schema exactly:\n"
        f"{json.dumps(schema, indent=2)}"
    )


def analyze_image(
    client: OpenAI,
    model: str,
    uploaded_file: Any,
    project_name: str,
    work_type: str,
    inspection_brief: str,
) -> InspectionResult:
    try:
        response = client.responses.create(
            model=model,
            input=[
                {
                    "role": "system",
                    "content": [{"type": "input_text", "text": SYSTEM_PROMPT}],
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "input_text",
                            "text": build_image_prompt(
                                uploaded_file.name,
                                project_name,
                                work_type,
                                inspection_brief,
                            ),
                        },
                        {"type": "input_image", "image_url": encode_image(uploaded_file)},
                    ],
                },
            ],
        )
        payload = parse_json_response(response.output_text)
        return normalize_result(uploaded_file.name, project_name, work_type, payload)
    except Exception as exc:
        return default_result(uploaded_file.name, project_name, work_type, f"Analysis failed: {exc}")


def summarize_site(
    client: OpenAI,
    model: str,
    reports: list[InspectionResult],
    inspection_brief: str,
) -> dict[str, Any]:
    average_score = round(sum(report.overall_score for report in reports) / len(reports))
    fallback = {
        "average_score": average_score,
        "site_quality": normalize_quality_level(None, average_score),
        "executive_summary": "Site summary generated from the per-image scores only.",
        "priority_actions": [
            "Review the lowest-scoring images first.",
            "Validate issues on site before taking corrective action.",
            "Capture follow-up photos after rework.",
        ],
        "common_themes": ["Mixed visible quality across uploaded images."],
    }

    try:
        summary_prompt = {
            "inspection_brief": inspection_brief.strip() or "General quality review",
            "reports": [report.normalized() for report in reports],
            "output_schema": {
                "average_score": "integer 0 to 100",
                "site_quality": "Excellent | Good | Fair | Poor | Critical",
                "executive_summary": "2 sentence site overview",
                "priority_actions": ["up to 4 short strings"],
                "common_themes": ["up to 4 short strings"],
            },
        }
        response = client.responses.create(
            model=model,
            input=[
                {
                    "role": "system",
                    "content": [
                        {
                            "type": "input_text",
                            "text": (
                                "You summarize construction-photo inspection results. "
                                "Only use the report data provided. Return strict JSON only."
                            ),
                        }
                    ],
                },
                {
                    "role": "user",
                    "content": [{"type": "input_text", "text": json.dumps(summary_prompt, indent=2)}],
                },
            ],
        )
        payload = parse_json_response(response.output_text)
        score = clamp_score(payload.get("average_score", average_score))
        return {
            "average_score": score,
            "site_quality": normalize_quality_level(payload.get("site_quality"), score),
            "executive_summary": str(payload.get("executive_summary") or fallback["executive_summary"]).strip(),
            "priority_actions": clean_list(
                payload.get("priority_actions"),
                fallback["priority_actions"][0],
                limit=4,
            ),
            "common_themes": clean_list(payload.get("common_themes"), fallback["common_themes"][0]),
        }
    except Exception:
        return fallback


def build_pdf_report(report: InspectionResult) -> bytes:
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=18 * mm,
        rightMargin=18 * mm,
        topMargin=18 * mm,
        bottomMargin=18 * mm,
    )
    styles = getSampleStyleSheet()
    title_style = styles["Title"]
    title_style.textColor = colors.HexColor("#143d59")
    title_style.fontName = "Helvetica-Bold"

    heading_style = ParagraphStyle(
        "SectionHeading",
        parent=styles["Heading2"],
        textColor=colors.HexColor("#143d59"),
        spaceAfter=6,
        spaceBefore=8,
    )
    body_style = ParagraphStyle(
        "Body",
        parent=styles["BodyText"],
        fontName="Helvetica",
        leading=14,
        spaceAfter=5,
    )
    bullet_style = ParagraphStyle(
        "BulletBody",
        parent=body_style,
        leftIndent=12,
        bulletIndent=0,
    )

    score_table = Table(
        [
            ["Project", report.project_name],
            ["Work type", report.work_type],
            ["Image", report.image_name],
            ["Quality level", report.quality_level],
            ["Score", f"{report.overall_score}/100"],
            ["Confidence", report.confidence],
        ],
        colWidths=[38 * mm, 126 * mm],
    )
    score_table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#eaf2f8")),
                ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#1f2933")),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (1, 0), (1, -1), "Helvetica"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#c7d3dd")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("PADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )

    story = [
        Paragraph("Construction Quality Report", title_style),
        Spacer(1, 5),
        score_table,
        Spacer(1, 12),
        Paragraph("Summary", heading_style),
        Paragraph(report.summary, body_style),
        Paragraph("Observations", heading_style),
    ]

    for item in report.observations:
        story.append(Paragraph(item, bullet_style, bulletText="-"))

    story.append(Paragraph("Risks", heading_style))
    for item in report.risks:
        story.append(Paragraph(item, bullet_style, bulletText="-"))

    story.append(Paragraph("Recommended Actions", heading_style))
    for item in report.recommended_actions:
        story.append(Paragraph(item, bullet_style, bulletText="-"))

    story.append(Paragraph("Limitations", heading_style))
    for item in report.limitations:
        story.append(Paragraph(item, bullet_style, bulletText="-"))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def quality_badge(level: str) -> str:
    colors = {
        "Excellent": "#1d8348",
        "Good": "#2e86c1",
        "Fair": "#b7950b",
        "Poor": "#ca6f1e",
        "Critical": "#c0392b",
    }
    color = colors.get(level, "#566573")
    return (
        f"<span style='background:{color}; color:white; padding:0.35rem 0.6rem; "
        "border-radius:999px; font-size:0.85rem; font-weight:600;'>"
        f"{level}</span>"
    )


st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)
st.caption(
    "Upload construction-site photos and generate AI-assisted quality reports for each image, "
    "plus an overall site summary."
)

with st.sidebar:
    st.header("Analysis Settings")
    api_key = st.text_input(
        "OpenAI API key",
        value=os.getenv("OPENAI_API_KEY", ""),
        type="password",
        help="Loaded from your local environment when available.",
    )
    model = st.selectbox("Model", MODEL_OPTIONS, index=MODEL_OPTIONS.index(DEFAULT_MODEL))
    inspection_brief = st.text_area(
        "Inspection focus",
        value="Check workmanship, finishing quality, visible defects, curing, alignment, and housekeeping.",
        help="Optional guidance for the AI on what to focus on.",
    )
    st.info(
        "This tool assesses only visible conditions in the uploaded photos. "
        "It is a site-review assistant, not a substitute for a licensed engineer or inspector."
    )

uploaded_files = st.file_uploader(
    "Upload construction photos",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True,
    help="You can upload several images and generate a separate report for each one.",
)

uploaded_names = tuple(file.name for file in uploaded_files) if uploaded_files else ()
if st.session_state.get("uploaded_names") != uploaded_names:
    st.session_state["uploaded_names"] = uploaded_names
    st.session_state.pop("inspection_reports", None)
    st.session_state.pop("site_summary", None)

analyze_clicked = st.button("Analyze Photos", type="primary", disabled=not uploaded_files or not api_key)

if not uploaded_files:
    st.write("Add a few site photos to begin the inspection.")

if uploaded_files:
    st.subheader("Uploaded Images")
    preview_columns = st.columns(min(3, len(uploaded_files)))
    for index, uploaded_file in enumerate(uploaded_files):
        with preview_columns[index % len(preview_columns)]:
            st.image(uploaded_file, caption=uploaded_file.name, use_container_width=True)

    st.subheader("Image Details")
    st.caption("Add project context for each image before running the analysis.")
    for uploaded_file in uploaded_files:
        with st.container(border=True):
            st.write(uploaded_file.name)
            st.text_input(
                "Project name",
                key=f"project_name::{uploaded_file.name}",
                placeholder="e.g. Tower A Podium Deck",
            )
            selected_work_type = st.selectbox(
                "Type of work",
                WORK_TYPE_OPTIONS,
                key=f"work_type::{uploaded_file.name}",
            )
            if selected_work_type == "Other":
                st.text_input(
                    "Custom work type",
                    key=f"work_type_custom::{uploaded_file.name}",
                    placeholder="Describe the work type",
                )

if analyze_clicked:
    missing_context: list[str] = []
    image_context: dict[str, dict[str, str]] = {}
    for uploaded_file in uploaded_files:
        project_name = st.session_state.get(f"project_name::{uploaded_file.name}", "").strip()
        work_type = st.session_state.get(f"work_type::{uploaded_file.name}", "").strip()
        if work_type == "Other":
            work_type = st.session_state.get(f"work_type_custom::{uploaded_file.name}", "").strip()

        if not project_name or not work_type:
            missing_context.append(uploaded_file.name)
            continue

        image_context[uploaded_file.name] = {
            "project_name": project_name,
            "work_type": work_type,
        }

    if missing_context:
        st.error(
            "Please add a project name and work type for each uploaded image before analyzing: "
            + ", ".join(missing_context)
        )
        st.stop()

    client = OpenAI(api_key=api_key)
    progress_bar = st.progress(0)
    reports: list[InspectionResult] = []

    for index, uploaded_file in enumerate(uploaded_files, start=1):
        context = image_context[uploaded_file.name]
        reports.append(
            analyze_image(
                client,
                model,
                uploaded_file,
                context["project_name"],
                context["work_type"],
                inspection_brief,
            )
        )
        progress_bar.progress(index / len(uploaded_files))

    site_summary = summarize_site(client, model, reports, inspection_brief)
    st.session_state["inspection_reports"] = [report.normalized() for report in reports]
    st.session_state["site_summary"] = site_summary
    progress_bar.empty()

reports_data = st.session_state.get("inspection_reports")
site_summary = st.session_state.get("site_summary")

if reports_data and site_summary:
    reports = [
        normalize_result(
            item["image_name"],
            item.get("project_name", ""),
            item.get("work_type", ""),
            item,
        )
        for item in reports_data
    ]

    score_col, quality_col, images_col = st.columns(3)
    score_col.metric("Average Site Score", f"{site_summary['average_score']}/100")
    quality_col.write("Site Quality")
    quality_col.markdown(quality_badge(site_summary["site_quality"]), unsafe_allow_html=True)
    images_col.metric("Images Reviewed", len(reports))

    st.subheader("Site Summary")
    st.write(site_summary["executive_summary"])

    summary_col, actions_col = st.columns(2)
    with summary_col:
        st.write("Common themes")
        for theme in site_summary["common_themes"]:
            st.write(f"- {theme}")
    with actions_col:
        st.write("Priority actions")
        for action in site_summary["priority_actions"]:
            st.write(f"- {action}")

    st.subheader("Per-Image Reports")
    for report in reports:
        with st.container(border=True):
            header_col, score_col = st.columns([4, 1])
            with header_col:
                st.markdown(f"### {report.image_name}")
                st.caption(f"Project: {report.project_name} | Work type: {report.work_type}")
                st.markdown(quality_badge(report.quality_level), unsafe_allow_html=True)
            with score_col:
                st.metric("Score", f"{report.overall_score}/100")

            image_match = next(
                (uploaded_file for uploaded_file in uploaded_files if uploaded_file.name == report.image_name),
                None,
            )
            if image_match is not None:
                st.image(image_match, use_container_width=True)

            st.write(report.summary)
            detail_col_1, detail_col_2 = st.columns(2)
            with detail_col_1:
                st.write("Observations")
                for item in report.observations:
                    st.write(f"- {item}")
                st.write("Risks")
                for item in report.risks:
                    st.write(f"- {item}")
            with detail_col_2:
                st.write("Recommended actions")
                for item in report.recommended_actions:
                    st.write(f"- {item}")
                st.write("Limitations")
                for item in report.limitations:
                    st.write(f"- {item}")
                st.caption(f"Confidence: {report.confidence}")
                pdf_bytes = build_pdf_report(report)
                st.download_button(
                    "Download PDF Report",
                    data=pdf_bytes,
                    file_name=(
                        f"{sanitize_filename(report.project_name)}_"
                        f"{sanitize_filename(report.image_name)}_quality_report.pdf"
                    ),
                    mime="application/pdf",
                    key=f"pdf::{report.image_name}",
                    use_container_width=True,
                )
