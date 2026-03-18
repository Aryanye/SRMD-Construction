from __future__ import annotations

import base64
from io import BytesIO
import json
import mimetypes
import os
from pathlib import Path
import re
from dataclasses import dataclass, field
from typing import Any

from dotenv import load_dotenv
from openai import OpenAI
import requests
from PIL import Image as PILImage, ImageDraw
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.platypus import Image as RLImage
from reportlab.platypus import KeepTogether, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
import streamlit as st

load_dotenv(dotenv_path=Path(__file__).resolve().parents[1] / ".env")

APP_TITLE = "Construction Quality Inspector"
DEFAULT_MODEL = "gpt-5.1"
ZOHO_PORTAL_ID = os.getenv("ZOHO_PORTAL_ID", "60062895348")
ZOHO_CLIENT_ID = os.getenv("ZOHO_CLIENT_ID", "").strip()
ZOHO_CLIENT_SECRET = os.getenv("ZOHO_CLIENT_SECRET", "").strip()
ZOHO_REFRESH_TOKEN = os.getenv("ZOHO_REFRESH_TOKEN", "").strip()
MODEL_OPTIONS = ["gpt-5.1", "gpt-5-mini", "gpt-4.1"]
QUALITY_LEVELS = ["Excellent", "Good", "Fair", "Poor", "Critical"]
CONFIDENCE_LEVELS = ["High", "Medium", "Low"]
SYSTEM_PROMPT = """
You are a senior construction quality inspector with deep expertise across all construction disciplines.
Assess the uploaded construction-site photos as one project batch.

Step 1 — Discipline detection: Identify every construction discipline that has visible work in the photos.
Common disciplines include (but are not limited to): Structural, Civil/Earthworks, Concrete Works,
Formwork, Reinforcement/Rebar, Masonry/Blockwork, MEP – Mechanical, MEP – Electrical,
MEP – Plumbing/Drainage, Architectural Finishes, Waterproofing, Roofing, Facade/Cladding,
Safety/Housekeeping, Landscaping. List all that apply.

Step 2 — Comprehensive analysis: For each identified discipline, assess quality standards,
workmanship, visible defects, safety compliance, and required corrective actions. Consider:
- Material quality and condition
- Workmanship and installation accuracy
- Alignment, levels, and tolerances
- Surface finish and curing (where applicable)
- Safety hazards, PPE usage, and site housekeeping
- Sequencing and coordination between trades

Step 3 — For each image, identify specific areas of concern and mark their approximate positions
using normalised bounding-box coordinates [x1, y1, x2, y2] where (0,0) is top-left and (1,1) is bottom-right.

Rules:
- Analyze the full batch together. If only one image is provided, base the full report on that image.
- Base all judgments only on what is visibly present in the photos.
- Do not assert code compliance, structural adequacy, or hidden defects without explicitly stating they cannot be verified visually.
- Clearly identify the most important improvements, ordered by priority.
- Return strict JSON only — no markdown, no prose outside the JSON object.
""".strip()

ANNOTATION_COLORS = [
    (231, 76, 60),    # red
    (230, 126, 34),   # orange
    (41, 128, 185),   # blue
    (142, 68, 173),   # purple
    (22, 160, 133),   # teal
    (39, 174, 96),    # green
]


@dataclass
class BatchInspectionResult:
    project_name: str
    image_names: list[str]
    overall_score: int
    quality_level: str
    executive_summary: str
    disciplines: list[str]
    discipline_notes: list[dict[str, str]]
    strengths: list[str]
    concerns: list[str]
    key_improvements: list[str]
    image_findings: list[dict[str, Any]]
    confidence: str
    limitations: list[str]

    def normalized(self) -> dict[str, Any]:
        return {
            "project_name": self.project_name,
            "image_names": self.image_names,
            "overall_score": self.overall_score,
            "quality_level": self.quality_level,
            "executive_summary": self.executive_summary,
            "disciplines": self.disciplines,
            "discipline_notes": self.discipline_notes,
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


def normalize_image_findings(value: Any, image_names: list[str]) -> list[dict[str, Any]]:
    findings: list[dict[str, Any]] = []
    if isinstance(value, list):
        for item in value:
            if isinstance(item, dict):
                image_name = str(item.get("image_name") or "").strip()
                note = str(item.get("finding") or "").strip()
                if image_name and note:
                    highlights = _normalize_highlights(item.get("highlights"))
                    findings.append({"image_name": image_name, "finding": note, "highlights": highlights})
    if findings:
        return findings[: len(image_names)]
    return [{"image_name": name, "finding": "No separate image note returned.", "highlights": []} for name in image_names]


def _normalize_highlights(value: Any) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    if not isinstance(value, list):
        return result
    for item in value:
        if not isinstance(item, dict):
            continue
        bbox = item.get("bbox")
        if not isinstance(bbox, list) or len(bbox) != 4:
            continue
        try:
            x1, y1, x2, y2 = [float(c) for c in bbox]
        except (TypeError, ValueError):
            continue
        # Clamp to [0,1]
        x1, y1, x2, y2 = (max(0.0, min(1.0, v)) for v in (x1, y1, x2, y2))
        label = str(item.get("label") or "").strip()
        detail = str(item.get("detail") or "").strip()
        if label or detail:
            result.append({"label": label, "detail": detail, "bbox": [x1, y1, x2, y2]})
    return result[:6]


def _normalize_discipline_notes(value: Any) -> list[dict[str, str]]:
    notes: list[dict[str, str]] = []
    if not isinstance(value, list):
        return notes
    for item in value:
        if isinstance(item, dict):
            discipline = str(item.get("discipline") or "").strip()
            note = str(item.get("note") or "").strip()
            if discipline and note:
                notes.append({"discipline": discipline, "note": note})
    return notes[:12]


def normalize_result(project_name: str, image_names: list[str], payload: dict[str, Any]) -> BatchInspectionResult:
    score = clamp_score(payload.get("overall_score"))
    return BatchInspectionResult(
        project_name=project_name.strip() or "Untitled project",
        image_names=image_names,
        overall_score=score,
        quality_level=normalize_quality_level(payload.get("quality_level"), score),
        executive_summary=str(payload.get("executive_summary") or "No summary returned.").strip(),
        disciplines=clean_list(payload.get("disciplines"), "Unknown discipline", limit=12),
        discipline_notes=_normalize_discipline_notes(payload.get("discipline_notes")),
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
        disciplines=[],
        discipline_notes=[],
        strengths=["No analysis was completed."],
        concerns=["Analysis did not complete successfully."],
        key_improvements=["Retry the upload or review the API configuration."],
        image_findings=[{"image_name": name, "finding": message, "highlights": []} for name in image_names],
        confidence="Low",
        limitations=["No AI assessment was produced."],
    )


def encode_image(uploaded_file: Any) -> str:
    image_bytes = uploaded_file.getvalue()
    mime_type = uploaded_file.type or mimetypes.guess_type(uploaded_file.name)[0] or "image/jpeg"
    encoded = base64.b64encode(image_bytes).decode("ascii")
    return f"data:{mime_type};base64,{encoded}"


def build_batch_prompt(
    project_name: str,
    image_names: list[str],
    task_context: str,
    inspection_brief: str,
) -> str:
    task_ctx = task_context.strip() or "Not provided."
    focus = inspection_brief.strip() or "General quality and workmanship across all visible trades."
    schema = {
        "project_name": project_name,
        "disciplines": [
            "list every discipline with visible work, e.g. 'Concrete Works', 'Reinforcement/Rebar', 'MEP – Electrical'"
        ],
        "overall_score": "integer from 0 to 100",
        "quality_level": "Excellent | Good | Fair | Poor | Critical",
        "executive_summary": "paragraph under 80 words summarising all identified disciplines and overall site condition",
        "discipline_notes": [
            {
                "discipline": "discipline name exactly as listed in disciplines array",
                "note": "one concise paragraph assessing quality, defects, and required actions for this discipline",
            }
        ],
        "strengths": ["up to 4 short strings"],
        "concerns": ["up to 5 short strings ordered by severity"],
        "key_improvements": ["up to 5 short strings ordered by priority"],
        "image_findings": [
            {
                "image_name": "one of the uploaded image names",
                "finding": "one short image-level summary",
                "highlights": [
                    {
                        "label": "2–4 word label, e.g. 'Exposed rebar'",
                        "detail": "one sentence describing the issue at this location",
                        "bbox": [0.1, 0.2, 0.6, 0.8],
                    }
                ],
            }
        ],
        "confidence": "High | Medium | Low",
        "limitations": ["1 to 3 short strings"],
    }
    return (
        f"Project: '{project_name}'. Uploaded images: {', '.join(image_names)}.\n"
        f"Task context: {task_ctx}\n"
        f"Inspection focus: {focus}\n\n"
        "Follow the two-step process in the system prompt: first identify all disciplines, "
        "then provide a thorough per-discipline analysis. "
        "For each image, mark every area of concern with a normalised bounding box [x1, y1, x2, y2] "
        "where (0,0) is top-left and (1,1) is bottom-right.\n\n"
        "Return a strict JSON object following this schema exactly:\n"
        f"{json.dumps(schema, indent=2)}"
    )


def analyze_batch(
    client: OpenAI,
    model: str,
    uploaded_files: list[Any],
    project_name: str,
    task_context: str,
    inspection_brief: str,
) -> BatchInspectionResult:
    image_names = [file.name for file in uploaded_files]
    try:
        content: list[dict[str, str]] = [
            {
                "type": "input_text",
                "text": build_batch_prompt(project_name, image_names, task_context, inspection_brief),
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


def annotate_image(image_bytes: bytes, highlights: list[dict[str, Any]]) -> bytes:
    """Draw numbered bounding boxes on image regions identified by the AI."""
    img = PILImage.open(BytesIO(image_bytes)).convert("RGB")
    w, h = img.size
    overlay = PILImage.new("RGBA", (w, h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(overlay)

    for i, hl in enumerate(highlights):
        bbox = hl.get("bbox", [])
        if len(bbox) != 4:
            continue
        x1, y1, x2, y2 = bbox
        px1 = int(x1 * w)
        py1 = int(y1 * h)
        px2 = int(x2 * w)
        py2 = int(y2 * h)
        if px2 <= px1 or py2 <= py1:
            continue

        r, g, b = ANNOTATION_COLORS[i % len(ANNOTATION_COLORS)]
        # Semi-transparent fill
        draw.rectangle([px1, py1, px2, py2], fill=(r, g, b, 50))
        # Solid border
        border = max(3, int(min(w, h) * 0.004))
        draw.rectangle([px1, py1, px2, py2], outline=(r, g, b, 230), width=border)

        # Number badge
        badge_r = max(14, int(min(w, h) * 0.025))
        bx = px1 + badge_r + border
        by = py1 + badge_r + border
        draw.ellipse([bx - badge_r, by - badge_r, bx + badge_r, by + badge_r], fill=(r, g, b, 255))
        # Draw number text
        font_size = max(10, badge_r)
        try:
            from PIL import ImageFont
            font = ImageFont.truetype("/System/Library/Fonts/Helvetica.ttc", font_size)
        except Exception:
            font = ImageFont.load_default()
        draw.text((bx, by), str(i + 1), fill=(255, 255, 255, 255), font=font, anchor="mm")

    img_rgba = img.convert("RGBA")
    composited = PILImage.alpha_composite(img_rgba, overlay).convert("RGB")
    out = BytesIO()
    composited.save(out, format="JPEG", quality=88)
    return out.getvalue()


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


def build_compact_table(
    title: str,
    rows: list[list[Any]],
    col_widths: list[float],
    header_color: str = "#143d59",
    body_fill: str = "#f8fbfd",
    row_heights: list[float] | None = None,
) -> Table:
    header_style = ParagraphStyle(
        "TableHeader",
        parent=getSampleStyleSheet()["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=9.5,
        leading=11,
        textColor=colors.white,
        spaceAfter=0,
    )
    title_row = [[Paragraph(title, header_style)] + [""] * (len(col_widths) - 1)]
    table = Table(title_row + rows, colWidths=col_widths, rowHeights=row_heights, hAlign="LEFT")
    table.setStyle(
        TableStyle(
            [
                ("SPAN", (0, 0), (-1, 0)),
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(header_color)),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor(body_fill)),
                ("TEXTCOLOR", (0, 1), (-1, -1), colors.HexColor("#1f2933")),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#d6dee6")),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("PADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    return table


def build_matrix_card(
    title: str,
    items: list[str],
    item_style: ParagraphStyle,
    body_fill: str,
    max_rows: int = 4,
) -> Table:
    number_style = ParagraphStyle(
        "CardNumber",
        parent=item_style,
        fontName="Helvetica-Bold",
        alignment=1,
        leading=9,
    )
    rows: list[list[Any]] = []
    trimmed_items = items[:max_rows]
    for index in range(max_rows):
        if index < len(trimmed_items):
            rows.append(
                [
                    Paragraph(str(index + 1), number_style),
                    Paragraph(trimmed_items[index], item_style),
                ]
            )
        else:
            rows.append(["", ""])

    table = build_compact_table(
        title,
        rows,
        [10 * mm, 76 * mm],
        body_fill=body_fill,
        row_heights=[10 * mm] + [15 * mm] * max_rows,
    )
    return table


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
        fontSize=9,
        leading=11,
        spaceAfter=4,
    )
    compact_style = ParagraphStyle(
        "CompactBody",
        parent=body_style,
        fontSize=8,
        leading=9.5,
        spaceAfter=0,
    )

    disciplines_text = ", ".join(result.disciplines) if result.disciplines else "Not identified"
    summary_table = Table(
        [
            ["Project", result.project_name],
            ["Disciplines", disciplines_text],
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

    strengths_rows = result.strengths
    concerns_rows = result.concerns
    improvements_rows = result.key_improvements
    limitation_rows = result.limitations
    image_note_rows = [
        [Paragraph(f"<b>{item['image_name']}</b>", compact_style), Paragraph(item["finding"], compact_style)]
        for item in result.image_findings
    ]

    # Discipline assessment table rows
    discipline_note_rows = [
        [Paragraph(f"<b>{dn['discipline']}</b>", compact_style), Paragraph(dn["note"], compact_style)]
        for dn in result.discipline_notes
    ]

    summary_story: list[Any] = [
        Paragraph("Construction Quality Batch Report", title_style),
        Spacer(1, 4),
        summary_table,
        Spacer(1, 8),
        build_compact_table(
            "Executive Summary",
            [[Paragraph(result.executive_summary, compact_style)]],
            [172 * mm],
        ),
    ]

    if discipline_note_rows:
        summary_story += [
            Spacer(1, 6),
            build_compact_table(
                "Discipline Assessment",
                discipline_note_rows,
                [48 * mm, 124 * mm],
                body_fill="#f5f8ff",
            ),
        ]

    summary_story += [
        Spacer(1, 6),
        Table(
            [
                [
                    build_matrix_card("What Is Working Well", strengths_rows, compact_style, body_fill="#f6fbf6"),
                    build_matrix_card("Key Concerns", concerns_rows, compact_style, body_fill="#fff7f5"),
                ],
                [
                    build_matrix_card("Priority Improvements", improvements_rows, compact_style, body_fill="#fffdf2"),
                    build_matrix_card("Limitations", limitation_rows, compact_style, body_fill="#f6f8fa"),
                ],
            ],
            colWidths=[86 * mm, 86 * mm],
            hAlign="LEFT",
        ),
        Spacer(1, 6),
        build_compact_table("Image Notes", image_note_rows, [44 * mm, 128 * mm], body_fill="#f8fbfd"),
    ]

    story: list[Any] = [KeepTogether(summary_story)]

    # Build lookup: image name → highlights
    highlights_by_name: dict[str, list[dict[str, Any]]] = {
        finding["image_name"]: finding.get("highlights", [])
        for finding in result.image_findings
    }

    if image_assets:
        story.append(Spacer(1, 10))
        story.append(Paragraph("Photo Record", heading_style))

        legend_num_style = ParagraphStyle(
            "LegendNum",
            parent=compact_style,
            fontName="Helvetica-Bold",
            alignment=1,
        )
        legend_label_style = ParagraphStyle(
            "LegendLabel",
            parent=compact_style,
            fontName="Helvetica-Bold",
        )

        for asset in image_assets:
            highlights = highlights_by_name.get(asset["name"], [])
            try:
                annotated_bytes = annotate_image(asset["bytes"], highlights) if highlights else asset["bytes"]
            except Exception:
                annotated_bytes = asset["bytes"]

            image_block = make_pdf_image(annotated_bytes, asset["name"])
            story += image_block
            story.append(Spacer(1, 4))

            if highlights:
                legend_rows = [
                    [
                        Paragraph(str(i + 1), legend_num_style),
                        Paragraph(f"<b>{hl.get('label', '')}</b>", legend_label_style),
                        Paragraph(hl.get("detail", ""), compact_style),
                    ]
                    for i, hl in enumerate(highlights)
                ]
                legend_table = build_compact_table(
                    f"Annotated findings — {asset['name']}",
                    legend_rows,
                    [8 * mm, 44 * mm, 120 * mm],
                    header_color="#2c3e50",
                    body_fill="#fafbfc",
                )
                story.append(legend_table)
                story.append(Spacer(1, 8))

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


@st.cache_data(ttl=300, show_spinner=False)
def fetch_zoho_project_names() -> list[str]:
    """Fetch active project names from Zoho Projects containing the word 'Project'."""
    token_resp = requests.post(
        "https://accounts.zoho.in/oauth/v2/token",
        params={
            "client_id": ZOHO_CLIENT_ID,
            "client_secret": ZOHO_CLIENT_SECRET,
            "refresh_token": ZOHO_REFRESH_TOKEN,
            "grant_type": "refresh_token",
        },
        timeout=10,
    )
    token_resp.raise_for_status()
    access_token = token_resp.json()["access_token"]

    projects_resp = requests.get(
        f"https://projectsapi.zoho.in/restapi/portal/{ZOHO_PORTAL_ID}/projects/",
        headers={"Authorization": f"Zoho-oauthtoken {access_token}"},
        params={"per_page": 100},
        timeout=10,
    )
    projects_resp.raise_for_status()
    projects = projects_resp.json().get("projects", [])
    return sorted([p["name"] for p in projects if "project" in p["name"].lower()])


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

_zoho_ready = all([ZOHO_CLIENT_ID, ZOHO_CLIENT_SECRET, ZOHO_REFRESH_TOKEN, ZOHO_PORTAL_ID])
if _zoho_ready:
    try:
        with st.spinner("Loading projects from Zoho…"):
            _zoho_names = fetch_zoho_project_names()
        if _zoho_names:
            project_name = st.selectbox(
                "Project name",
                options=_zoho_names,
                help="Projects imported from Zoho Projects (names containing 'Project'). Refreshes every 5 minutes.",
            )
        else:
            st.warning("No projects containing 'Project' found in Zoho Projects.")
            project_name = st.text_input("Project name", placeholder="e.g. Projects - NGH A")
    except Exception as _zoho_err:
        st.warning(f"Could not load Zoho projects ({_zoho_err}). Enter the name manually.")
        project_name = st.text_input("Project name", placeholder="e.g. Projects - NGH A")
else:
    project_name = st.text_input("Project name", placeholder="e.g. SRMD Tower A Podium")
task_context = st.text_area(
    "Task context (optional)",
    placeholder=(
        "Describe the specific work or stage being inspected, "
        "e.g. 'Day 3 of Level 5 slab formwork installation' or "
        "'Post-pour inspection of retaining wall at Grid B'"
    ),
    help="Providing context helps the AI give a more accurate, targeted analysis.",
    height=80,
)
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
        result = analyze_batch(
            client, model, list(uploaded_files), project_name.strip(), task_context, inspection_brief
        )
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

    if result.disciplines:
        st.markdown(
            "**Disciplines identified:** " + " &nbsp;·&nbsp; ".join(result.disciplines),
            unsafe_allow_html=True,
        )

    st.subheader("Executive Summary")
    st.write(result.executive_summary)

    if result.discipline_notes:
        st.subheader("Discipline Assessment")
        for dn in result.discipline_notes:
            st.markdown(f"**{dn['discipline']}** — {dn['note']}")

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
        for i, hl in enumerate(item.get("highlights", [])):
            st.caption(f"  [{i + 1}] {hl.get('label', '')} — {hl.get('detail', '')}")

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
