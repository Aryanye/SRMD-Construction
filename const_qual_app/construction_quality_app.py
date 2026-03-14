from __future__ import annotations

import base64
from io import BytesIO
import json
import mimetypes
import os
from pathlib import Path
import re
from dataclasses import dataclass
from typing import Any

from dotenv import load_dotenv
from openai import OpenAI
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Image as RLImage
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import streamlit as st

load_dotenv(dotenv_path=Path(__file__).resolve().parents[1] / ".env")

APP_TITLE = "Construction Quality Inspector"
DEFAULT_MODEL = "gpt-5.1"
MODEL_OPTIONS = ["gpt-5.1", "gpt-5-mini", "gpt-4.1"]
QUALITY_LEVELS = ["Excellent", "Good", "Fair", "Poor", "Critical"]
CONFIDENCE_LEVELS = ["High", "Medium", "Low"]
SYSTEM_PROMPT = """
You are a construction quality inspector assistant.
Assess uploaded construction-site photos as one project batch.

Rules:
- Analyze the full batch together. If only one image is provided, analyze that single image.
- Base judgments only on what is visible in the photos.
- Do not claim code compliance, structural adequacy, or hidden defects unless you clearly state that they cannot be verified visually.
- Clearly identify the most important improvements needed.
- Return strict JSON only.
""".strip()


@dataclass
class BatchInspectionResult:
    project_name: str
    image_names: list[str]
    overall_score: int
    quality_level: str
    executive_summary: str
    strengths: list[str]
    concerns: list[str]
    key_improvements: list[str]
    image_findings: list[dict[str, str]]
    confidence: str
    limitations: list[str]

    def normalized(self) -> dict[str, Any]:
        return {
            "project_name": self.project_name,
            "image_names": self.image_names,
            "overall_score": self.overall_score,
            "quality_level": self.quality_level,
            "executive_summary": self.executive_summary,
            "strengths": self.strengths,
            "concerns": self.concerns,
            "key_improvements": self.key_improvements,
            "image_findings": self.image_findings,
            "confidence": self.confidence,
            "limitations": self.limitations,
        }


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


def normalize_image_findings(value: Any, image_names: list[str]) -> list[dict[str, str]]:
    findings: list[dict[str, str]] = []
    if isinstance(value, list):
        for item in value:
            if isinstance(item, dict):
                image_name = str(item.get("image_name") or "").strip()
                note = str(item.get("finding") or "").strip()
                if image_name and note:
                    findings.append({"image_name": image_name, "finding": note})
    if findings:
        return findings[: len(image_names)]
    return [{"image_name": name, "finding": "No separate image note returned."} for name in image_names]


def normalize_result(project_name: str, image_names: list[str], payload: dict[str, Any]) -> BatchInspectionResult:
    score = clamp_score(payload.get("overall_score"))
    return BatchInspectionResult(
        project_name=project_name.strip() or "Untitled project",
        image_names=image_names,
        overall_score=score,
        quality_level=normalize_quality_level(payload.get("quality_level"), score),
        executive_summary=str(payload.get("executive_summary") or "No summary returned.").strip(),
        strengths=clean_list(payload.get("strengths"), "No strengths returned."),
        concerns=clean_list(payload.get("concerns"), "No concerns returned."),
        key_improvements=clean_list(
            payload.get("key_improvements"),
            "No key improvements returned.",
            limit=5,
        ),
        image_findings=normalize_image_findings(payload.get("image_findings"), image_names),
        confidence=normalize_confidence(payload.get("confidence")),
        limitations=clean_list(
            payload.get("limitations"),
            "This review is based only on visible conditions in the uploaded photos.",
            limit=3,
        ),
    )


def default_result(project_name: str, image_names: list[str], message: str) -> BatchInspectionResult:
    return BatchInspectionResult(
        project_name=project_name,
        image_names=image_names,
        overall_score=0,
        quality_level="Critical",
        executive_summary=message,
        strengths=["No analysis was completed."],
        concerns=["Analysis did not complete successfully."],
        key_improvements=["Retry the upload or review the API configuration."],
        image_findings=[{"image_name": name, "finding": message} for name in image_names],
        confidence="Low",
        limitations=["No AI assessment was produced."],
    )


def encode_image(uploaded_file: Any) -> str:
    image_bytes = uploaded_file.getvalue()
    mime_type = uploaded_file.type or mimetypes.guess_type(uploaded_file.name)[0] or "image/jpeg"
    encoded = base64.b64encode(image_bytes).decode("ascii")
    return f"data:{mime_type};base64,{encoded}"


def build_batch_prompt(project_name: str, image_names: list[str], inspection_brief: str) -> str:
    extra_brief = inspection_brief.strip() or "Focus on workmanship, visible defects, finishes, alignment, curing, and housekeeping."
    schema = {
        "project_name": project_name,
        "overall_score": "integer from 0 to 100",
        "quality_level": "Excellent | Good | Fair | Poor | Critical",
        "executive_summary": "short paragraph under 70 words",
        "strengths": ["up to 4 short strings"],
        "concerns": ["up to 5 short strings"],
        "key_improvements": ["up to 5 short strings ordered by priority"],
        "image_findings": [
            {
                "image_name": "one of the uploaded image names",
                "finding": "one short image-specific note",
            }
        ],
        "confidence": "High | Medium | Low",
        "limitations": ["1 to 3 short strings"],
    }
    return (
        f"Inspect this batch of construction-site photos for the project '{project_name}'. "
        f"The uploaded image names are: {', '.join(image_names)}. "
        f"Additional inspection brief: {extra_brief}\n\n"
        "Analyze the full batch as one project report. If only one image is uploaded, base the full report on that single image. "
        "Be explicit about the most important improvements that should be made.\n\n"
        "Return a strict JSON object that follows this schema exactly:\n"
        f"{json.dumps(schema, indent=2)}"
    )


def analyze_batch(
    client: OpenAI,
    model: str,
    uploaded_files: list[Any],
    project_name: str,
    inspection_brief: str,
) -> BatchInspectionResult:
    image_names = [file.name for file in uploaded_files]
    try:
        content: list[dict[str, str]] = [
            {
                "type": "input_text",
                "text": build_batch_prompt(project_name, image_names, inspection_brief),
            }
        ]
        for uploaded_file in uploaded_files:
            content.append({"type": "input_image", "image_url": encode_image(uploaded_file)})

        response = client.responses.create(
            model=model,
            input=[
                {
                    "role": "system",
                    "content": [{"type": "input_text", "text": SYSTEM_PROMPT}],
                },
                {
                    "role": "user",
                    "content": content,
                },
            ],
        )
        payload = parse_json_response(response.output_text)
        return normalize_result(project_name, image_names, payload)
    except Exception as exc:
        return default_result(project_name, image_names, f"Analysis failed: {exc}")


def make_pdf_image(image_bytes: bytes, image_name: str) -> list[Any]:
    image_buffer = BytesIO(image_bytes)
    reader = ImageReader(image_buffer)
    original_width, original_height = reader.getSize()
    max_width = 82 * mm
    max_height = 58 * mm
    scale = min(max_width / original_width, max_height / original_height)
    pdf_image = RLImage(image_buffer, width=original_width * scale, height=original_height * scale)

    styles = getSampleStyleSheet()
    caption_style = ParagraphStyle(
        "ImageCaption",
        parent=styles["BodyText"],
        alignment=1,
        fontSize=8.5,
        textColor=colors.HexColor("#425466"),
        spaceBefore=4,
    )
    return [pdf_image, Spacer(1, 3), Paragraph(image_name, caption_style)]


def build_pdf_report(result: BatchInspectionResult, image_assets: list[dict[str, Any]]) -> bytes:
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        leftMargin=16 * mm,
        rightMargin=16 * mm,
        topMargin=16 * mm,
        bottomMargin=16 * mm,
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

    summary_table = Table(
        [
            ["Project", result.project_name],
            ["Images reviewed", str(len(result.image_names))],
            ["Quality level", result.quality_level],
            ["Score", f"{result.overall_score}/100"],
            ["Confidence", result.confidence],
        ],
        colWidths=[40 * mm, 130 * mm],
    )
    summary_table.setStyle(
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

    story: list[Any] = [
        Paragraph("Construction Quality Batch Report", title_style),
        Spacer(1, 6),
        summary_table,
        Spacer(1, 12),
        Paragraph("Executive Summary", heading_style),
        Paragraph(result.executive_summary, body_style),
        Paragraph("What Is Working Well", heading_style),
    ]

    for item in result.strengths:
        story.append(Paragraph(item, bullet_style, bulletText="-"))

    story.append(Paragraph("Key Concerns", heading_style))
    for item in result.concerns:
        story.append(Paragraph(item, bullet_style, bulletText="-"))

    story.append(Paragraph("Priority Improvements", heading_style))
    for item in result.key_improvements:
        story.append(Paragraph(item, bullet_style, bulletText="-"))

    story.append(Paragraph("Image Notes", heading_style))
    for item in result.image_findings:
        story.append(
            Paragraph(
                f"<b>{item['image_name']}</b>: {item['finding']}",
                body_style,
            )
        )

    story.append(Paragraph("Limitations", heading_style))
    for item in result.limitations:
        story.append(Paragraph(item, bullet_style, bulletText="-"))

    if image_assets:
        story.append(Paragraph("Photo Record", heading_style))
        image_cells: list[list[Any]] = []
        row: list[Any] = []
        for asset in image_assets:
            row.append(make_pdf_image(asset["bytes"], asset["name"]))
            if len(row) == 2:
                image_cells.append(row)
                row = []
        if row:
            row.append("")
            image_cells.append(row)

        image_table = Table(image_cells, colWidths=[86 * mm, 86 * mm], hAlign="LEFT")
        image_table.setStyle(
            TableStyle(
                [
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("BOX", (0, 0), (-1, -1), 0.4, colors.HexColor("#d6dee6")),
                    ("INNERGRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#d6dee6")),
                    ("PADDING", (0, 0), (-1, -1), 8),
                ]
            )
        )
        story.append(image_table)

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def quality_badge(level: str) -> str:
    level_colors = {
        "Excellent": "#1d8348",
        "Good": "#2e86c1",
        "Fair": "#b7950b",
        "Poor": "#ca6f1e",
        "Critical": "#c0392b",
    }
    color = level_colors.get(level, "#566573")
    return (
        f"<span style='background:{color}; color:white; padding:0.35rem 0.6rem; "
        "border-radius:999px; font-size:0.85rem; font-weight:600;'>"
        f"{level}</span>"
    )


st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)
st.caption(
    "Upload one or more photos from the same project. The app will analyze the full batch together "
    "and generate one project report."
)

with st.sidebar:
    st.header("Analysis Settings")
    model = st.selectbox("Model", MODEL_OPTIONS, index=MODEL_OPTIONS.index(DEFAULT_MODEL))
    inspection_brief = st.text_area(
        "Inspection focus",
        value="Check workmanship, visible defects, finish quality, alignment, curing, and housekeeping.",
        help="Optional guidance for the AI on what to focus on.",
    )
    st.info(
        "This tool assesses only visible conditions in the uploaded photos. "
        "It is a site-review assistant, not a substitute for a licensed engineer or inspector."
    )

api_key = os.getenv("OPENAI_API_KEY", "").strip()
if not api_key:
    st.warning("OpenAI API key not found in the environment. Add `OPENAI_API_KEY` to the repo root `.env` file.")

project_name = st.text_input("Project name", placeholder="e.g. SRMD Tower A Podium")
uploaded_files = st.file_uploader(
    "Upload project photos",
    type=["jpg", "jpeg", "png", "webp"],
    accept_multiple_files=True,
    help="Upload one photo or a group of photos from the same project batch.",
)

uploaded_names = tuple(file.name for file in uploaded_files) if uploaded_files else ()
if st.session_state.get("uploaded_names") != uploaded_names:
    st.session_state["uploaded_names"] = uploaded_names
    st.session_state.pop("batch_report", None)

analyze_clicked = st.button("Analyze Project Batch", type="primary", disabled=not uploaded_files or not api_key)

if not uploaded_files:
    st.write("Add one or more project photos to begin the batch inspection.")

if uploaded_files:
    st.subheader("Uploaded Project Photos")
    preview_columns = st.columns(min(3, len(uploaded_files)))
    for index, uploaded_file in enumerate(uploaded_files):
        with preview_columns[index % len(preview_columns)]:
            st.image(uploaded_file, caption=uploaded_file.name, use_container_width=True)

if analyze_clicked:
    if not project_name.strip():
        st.error("Please enter a project name for this batch before analyzing.")
        st.stop()

    client = OpenAI(api_key=api_key)
    with st.spinner("Analyzing the uploaded project batch..."):
        result = analyze_batch(client, model, list(uploaded_files), project_name.strip(), inspection_brief)
    st.session_state["batch_report"] = result.normalized()

batch_report_data = st.session_state.get("batch_report")
if batch_report_data:
    result = normalize_result(
        batch_report_data.get("project_name", project_name.strip()),
        batch_report_data.get("image_names", list(uploaded_names)),
        batch_report_data,
    )

    metric_col, badge_col, image_col = st.columns(3)
    metric_col.metric("Project Score", f"{result.overall_score}/100")
    badge_col.write("Quality Level")
    badge_col.markdown(quality_badge(result.quality_level), unsafe_allow_html=True)
    image_col.metric("Images Reviewed", len(result.image_names))

    st.subheader("Executive Summary")
    st.write(result.executive_summary)

    left_col, right_col = st.columns(2)
    with left_col:
        st.write("What is working well")
        for item in result.strengths:
            st.write(f"- {item}")
        st.write("Key concerns")
        for item in result.concerns:
            st.write(f"- {item}")
    with right_col:
        st.write("Priority improvements")
        for item in result.key_improvements:
            st.write(f"- {item}")
        st.write("Limitations")
        for item in result.limitations:
            st.write(f"- {item}")
        st.caption(f"Confidence: {result.confidence}")

    st.subheader("Image Notes")
    for item in result.image_findings:
        st.write(f"- **{item['image_name']}**: {item['finding']}")

    pdf_assets = [
        {"name": uploaded_file.name, "bytes": uploaded_file.getvalue()}
        for uploaded_file in uploaded_files
        if uploaded_file.name in result.image_names
    ]
    pdf_bytes = build_pdf_report(result, pdf_assets)
    st.download_button(
        "Download Project PDF Report",
        data=pdf_bytes,
        file_name=f"{sanitize_filename(result.project_name)}_construction_quality_report.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
