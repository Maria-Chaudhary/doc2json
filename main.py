"""Doc2JSON — Universal Document → JSON Extractor

Upload a document (PDF, Word, PowerPoint, Excel, RTF, HTML, text/code files)
and turn it into structured, schema-validated JSON using Gemini.
"""

import html
import json
import os
import tempfile
import time
from pathlib import Path

import gradio as gr
import pymupdf
from bs4 import BeautifulSoup
from docx import Document
from dotenv import load_dotenv
from google import genai
from google.genai import errors as genai_errors
from google.genai import types
from openpyxl import load_workbook
from pptx import Presentation
from pydantic import ValidationError
from striprtf.striprtf import rtf_to_text

from prompts import SYSTEM_PROMPT, build_prompt
from schemas import DocumentData


# ──────────────────────────────────────────────────────────────────────────
# Configuration
# ──────────────────────────────────────────────────────────────────────────

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

MAX_TEXT_CHARS = 120_000
DEFAULT_MODEL = "gemini-2.5-flash"

# Used automatically if the main model stays overloaded (HTTP 503).
FALLBACK_MODELS = ["gemini-2.5-flash-lite"]
# Seconds to wait before each retry of a temporary server error.
RETRY_DELAYS = (3, 8, 15)
TEMPORARY_CODES = {500, 502, 503, 504}

# Gemini-compatible structured JSON schema.
DOCUMENT_SCHEMA = {
    "type": "OBJECT",
    "properties": {
        "document_type": {"type": "STRING"},
        "title": {"type": "STRING", "nullable": True},
        "summary": {"type": "STRING", "nullable": True},
        "key_information": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "properties": {
                    "key": {"type": "STRING"},
                    "value": {"type": "STRING"},
                },
                "required": ["key", "value"],
            },
        },
        "sections": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "properties": {
                    "heading": {"type": "STRING"},
                    "content": {"type": "STRING"},
                },
                "required": ["heading", "content"],
            },
        },
        "entities": {
            "type": "ARRAY",
            "items": {
                "type": "OBJECT",
                "properties": {
                    "entity_type": {"type": "STRING"},
                    "value": {"type": "STRING"},
                },
                "required": ["entity_type", "value"],
            },
        },
        "language": {"type": "STRING", "nullable": True},
        "warnings": {
            "type": "ARRAY",
            "items": {"type": "STRING"},
        },
    },
    "required": [
        "document_type",
        "title",
        "summary",
        "language",
        "key_information",
        "sections",
        "entities",
        "warnings",
    ],
}

TEXT_EXTENSIONS = {
    ".txt", ".md", ".markdown", ".csv", ".tsv",
    ".json", ".jsonl", ".xml", ".yaml", ".yml",
    ".log", ".ini", ".cfg", ".conf", ".py", ".js",
    ".ts", ".jsx", ".tsx", ".java", ".c", ".cpp",
    ".h", ".hpp", ".cs", ".go", ".rs", ".php",
    ".rb", ".swift", ".kt", ".sql", ".sh", ".bat",
    ".ps1", ".css", ".scss", ".toml", ".rst", ".tex",
}

SUPPORTED_EXTENSIONS = TEXT_EXTENSIONS | {
    ".pdf", ".docx", ".pptx", ".xlsx", ".xlsm",
    ".rtf", ".html", ".htm",
}

UPLOAD_TYPES = sorted(SUPPORTED_EXTENSIONS)

EMPTY_STATUS = (
    '<div class="status status-idle">Upload a document and press '
    "<b>Extract to JSON</b> to begin.</div>"
)


# ──────────────────────────────────────────────────────────────────────────
# Text extraction
# ──────────────────────────────────────────────────────────────────────────

def read_text_file(path: Path) -> str:
    """Read a text file trying common encodings."""
    for encoding in ("utf-8-sig", "cp1252", "latin-1"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError(
        "This file could not be decoded. "
        "Save it as UTF-8 text and upload it again."
    )


def _row_text(cells) -> str:
    return " | ".join(cell.text.strip() for cell in cells)


def extract_document_text(file_path: str) -> tuple[str, str]:
    """Extract readable text from a supported document format.

    Returns (text, human_readable_format_name).
    """
    path = Path(file_path)

    if not path.exists() or not path.is_file():
        raise ValueError(
            "The uploaded file could not be found. Please upload it again."
        )

    extension = path.suffix.lower()

    if extension not in SUPPORTED_EXTENSIONS:
        raise ValueError(
            f"Unsupported file format: {extension or 'unknown'}. "
            "Please upload a supported document format."
        )

    # PDF ------------------------------------------------------------------
    if extension == ".pdf":
        with pymupdf.open(path) as pdf:
            if pdf.needs_pass:
                raise ValueError(
                    "This PDF is password-protected. "
                    "Remove the password and upload it again."
                )
            text = "\n\n".join(page.get_text("text") for page in pdf)

        if len(text.strip()) < 30:
            raise ValueError(
                "Very little text was found in this PDF. It may be scanned "
                "or image-only. Run OCR first to make it searchable, "
                "then upload it again."
            )
        return text, "PDF"

    # Word -----------------------------------------------------------------
    if extension == ".docx":
        document = Document(str(path))
        parts = [p.text for p in document.paragraphs if p.text.strip()]

        for table in document.tables:
            for row in table.rows:
                line = _row_text(row.cells)
                if line.strip(" |"):
                    parts.append(line)
        return "\n".join(parts), "Word document"

    # PowerPoint -----------------------------------------------------------
    if extension == ".pptx":
        presentation = Presentation(str(path))
        parts = []

        for number, slide in enumerate(presentation.slides, start=1):
            slide_parts = []

            for shape in slide.shapes:
                if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
                    slide_parts.append(shape.text_frame.text.strip())

                if getattr(shape, "has_table", False) and shape.has_table:
                    for row in shape.table.rows:
                        slide_parts.append(_row_text(row.cells))

            if slide.has_notes_slide:
                notes = slide.notes_slide.notes_text_frame.text.strip()
                if notes:
                    slide_parts.append(f"Notes: {notes}")

            parts.append(f"Slide {number}\n" + "\n".join(slide_parts))
        return "\n\n".join(parts), "PowerPoint presentation"

    # Excel ----------------------------------------------------------------
    if extension in {".xlsx", ".xlsm"}:
        workbook = load_workbook(path, read_only=True, data_only=True)
        parts = []
        try:
            for sheet in workbook.worksheets:
                parts.append(f"Sheet: {sheet.title}")
                for row in sheet.iter_rows(values_only=True):
                    values = ["" if v is None else str(v) for v in row]
                    if any(v.strip() for v in values):
                        parts.append(" | ".join(values))
        finally:
            workbook.close()
        return "\n".join(parts), "Spreadsheet"

    # RTF ------------------------------------------------------------------
    if extension == ".rtf":
        return rtf_to_text(read_text_file(path)), "RTF document"

    # HTML -----------------------------------------------------------------
    if extension in {".html", ".htm"}:
        soup = BeautifulSoup(read_text_file(path), "html.parser")
        for element in soup(["script", "style", "noscript"]):
            element.decompose()
        return soup.get_text(separator="\n", strip=True), "HTML document"

    # Plain text / source code ---------------------------------------------
    return read_text_file(path), "Text-based document"


# ──────────────────────────────────────────────────────────────────────────
# Rendering helpers
# ──────────────────────────────────────────────────────────────────────────

def render_status(kind: str, message: str) -> str:
    return f'<div class="status status-{kind}">{message}</div>'


def render_stats(data: dict, detected_format: str, chars: int, model: str) -> str:
    cards = [
        ("Format", detected_format),
        ("Characters", f"{chars:,}"),
        ("Key facts", str(len(data.get("key_information") or []))),
        ("Sections", str(len(data.get("sections") or []))),
        ("Entities", str(len(data.get("entities") or []))),
        ("Model", model),
    ]
    items = "".join(
        f'<div class="stat"><div class="stat-value">{html.escape(value)}</div>'
        f'<div class="stat-label">{label}</div></div>'
        for label, value in cards
    )
    return f'<div class="stats">{items}</div>'


def render_overview(data: dict) -> str:
    title = data.get("title") or "Untitled document"
    doc_type = data.get("document_type") or "Unknown"
    summary = data.get("summary") or "_No summary was generated._"
    language = data.get("language")

    meta = f"**Document type:** {doc_type}"
    if language:
        meta += f"  ·  **Language:** {language}"

    text = f"## {title}\n\n{meta}\n\n{summary}"

    warnings = data.get("warnings") or []
    if warnings:
        text += "\n\n---\n\n**⚠️ Notes & uncertainties**\n\n" + "\n".join(
            f"- {w}" for w in warnings
        )
    return text


def render_sections(data: dict) -> str:
    sections = data.get("sections") or []
    if not sections:
        return "_No sections were identified._"
    return "\n\n---\n\n".join(
        f"### {s['heading']}\n\n{s['content']}" for s in sections
    )


def safe_filename(stem: str) -> str:
    cleaned = "".join(c if c.isalnum() or c in "-_" else "_" for c in stem)
    return cleaned[:60] or "document"


# ──────────────────────────────────────────────────────────────────────────
# Main handler
# ──────────────────────────────────────────────────────────────────────────

def extract_document(file_path: str | None, extra_instructions: str):
    """Extract structured JSON from an uploaded document."""

    if not file_path:
        raise gr.Error("Please upload a document first.")

    api_key = os.getenv("GEMINI_API_KEY", "").strip()
    model = os.getenv("GEMINI_MODEL", "").strip() or DEFAULT_MODEL

    if not api_key:
        raise gr.Error("GEMINI_API_KEY is missing. Add your key to the .env file.")

    try:
        document_text, detected_format = extract_document_text(file_path)

        if not document_text.strip():
            raise gr.Error("No readable text was found in this document.")

        if len(document_text) > MAX_TEXT_CHARS:
            raise gr.Error(
                f"The document has {len(document_text):,} characters, which exceeds "
                f"the {MAX_TEXT_CHARS:,} character limit. Split it into smaller "
                "files and try again."
            )

        prompt = build_prompt(
            document_text=document_text,
            detected_format=detected_format,
            extra_instructions=(extra_instructions or "").strip(),
        )

        client = genai.Client(api_key=api_key)
        response, used_model = generate_with_retry(
            client,
            model,
            prompt,
            types.GenerateContentConfig(
                system_instruction=SYSTEM_PROMPT,
                response_mime_type="application/json",
                response_schema=DOCUMENT_SCHEMA,
                temperature=0,
            ),
        )

        if not response.text:
            raise gr.Error("The AI returned an empty response. Please try again.")

        # Validate the AI output before showing it.
        validated = DocumentData.model_validate_json(response.text)
        data = validated.model_dump(mode="json")
        output_json = json.dumps(data, ensure_ascii=False, indent=2)

        # Downloadable JSON file.
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            suffix=".json",
            prefix=f"{safe_filename(Path(file_path).stem)}_",
            delete=False,
        ) as tmp:
            tmp.write(output_json)
            json_path = tmp.name

        status = render_status(
            "ok",
            "✓ Extraction complete — output passed schema validation.",
        )
        stats = render_stats(data, detected_format, len(document_text), used_model)

        key_rows = [
            [item["key"], item["value"]] for item in data.get("key_information") or []
        ]
        entity_rows = [
            [item["entity_type"], item["value"]] for item in data.get("entities") or []
        ]

        return (
            status,
            stats,
            render_overview(data),
            key_rows,
            entity_rows,
            render_sections(data),
            output_json,
            json_path,
        )

    except gr.Error:
        raise

    # NOTE: pydantic's ValidationError subclasses ValueError, so it must be
    # caught first.
    except ValidationError as exc:
        print("Validation error:", exc)
        raise gr.Error(
            "The AI response did not match the JSON schema. "
            "Try again or shorten the document."
        ) from exc

    except ValueError as exc:
        raise gr.Error(str(exc)) from exc

    except Exception as exc:
        # Technical details stay in the terminal.
        print(f"Extraction error: {type(exc).__name__}: {exc}")
        message = str(exc).lower()

        if "503" in message or "unavailable" in message or "high demand" in message:
            raise gr.Error(
                "Google's Gemini servers are overloaded right now (not a problem "
                "with your app). The app already retried automatically. "
                "Please wait a minute and try again."
            ) from exc

        if "404" in message or "not found" in message:
            raise gr.Error(
                f"The model '{model}' was not found or is unavailable. "
                "Check GEMINI_MODEL in your .env file."
            ) from exc

        if any(k in message for k in ("api key", "api_key", "unauthorized", "permission", "401", "403")):
            raise gr.Error(
                "Gemini rejected the request. Check your API key and model access."
            ) from exc

        if any(k in message for k in ("429", "quota", "rate limit", "resource_exhausted")):
            raise gr.Error(
                "Gemini rate limit or quota reached. Wait a moment and try again."
            ) from exc

        raise gr.Error(
            "Extraction failed. Check the terminal for the technical error and "
            "verify your API key, model, internet connection, and API quota."
        ) from exc


def generate_with_retry(client, model: str, prompt: str, config):
    """Call Gemini, retrying temporary server errors (503 etc.) with backoff
    and falling back to another model if the main one stays overloaded.

    Returns (response, model_actually_used).
    """
    candidates = [model] + [m for m in FALLBACK_MODELS if m != model]
    last_exc = None

    for candidate in candidates:
        for delay in (0, *RETRY_DELAYS):
            if delay:
                time.sleep(delay)
            try:
                response = client.models.generate_content(
                    model=candidate, contents=prompt, config=config
                )
                return response, candidate
            except genai_errors.APIError as exc:
                if getattr(exc, "code", None) in TEMPORARY_CODES:
                    print(f"{candidate}: temporary error {exc.code}, retrying...")
                    last_exc = exc
                    continue
                raise

    raise last_exc


def clear_all():
    return (
        None,                # file
        "",                  # instructions
        EMPTY_STATUS,        # status
        "",                  # stats
        "",                  # overview
        [],                  # key info
        [],                  # entities
        "",                  # sections
        "",                  # json
        None,                # download
    )


# ──────────────────────────────────────────────────────────────────────────
# Styling  —  "Midnight Aurora" theme
# Deep midnight canvas, drifting aurora glows, frosted-glass panels and
# neon gradient accents. Looks identical in light and dark browser modes.
# ──────────────────────────────────────────────────────────────────────────

AURORA_VIOLET = gr.themes.Color(
    c50="#f1edff", c100="#e3dbff", c200="#c9b9ff", c300="#ad94ff",
    c400="#9270ff", c500="#7c5cff", c600="#6a43f0", c700="#5732cc",
    c800="#442a9e", c900="#31206f", c950="#1d1342",
    name="aurora_violet",
)

AURORA_AQUA = gr.themes.Color(
    c50="#ecfeff", c100="#cffafe", c200="#a5f3fc", c300="#67e8f9",
    c400="#22d3ee", c500="#06b6d4", c600="#0891b2", c700="#0e7490",
    c800="#155e75", c900="#164e63", c950="#083344",
    name="aurora_aqua",
)

MIDNIGHT = gr.themes.Color(
    c50="#f4f6ff", c100="#e6eaf8", c200="#cbd2ec", c300="#a3add4",
    c400="#7783ae", c500="#566189", c600="#3f496d", c700="#2d3553",
    c800="#1a2038", c900="#0f1428", c950="#070a16",
    name="midnight",
)

# One dark palette, applied to BOTH light and dark variants so the app
# always renders in the aurora look.
_AURORA_SETTINGS = dict(
    body_background_fill="#070a16",
    background_fill_primary="rgba(255,255,255,0.035)",
    background_fill_secondary="rgba(255,255,255,0.05)",
    body_text_color="#e8ecff",
    body_text_color_subdued="#9aa6cf",

    block_background_fill="rgba(255,255,255,0.045)",
    block_border_color="rgba(255,255,255,0.10)",
    block_border_width="1px",
    block_shadow="0 18px 40px -22px rgba(0,0,0,0.8)",
    block_label_background_fill="rgba(124,92,255,0.18)",
    block_label_text_color="#cbbcff",
    block_title_text_color="#e8ecff",

    input_background_fill="rgba(7,10,22,0.65)",
    input_border_color="rgba(255,255,255,0.14)",
    input_border_color_focus="#22d3ee",
    input_shadow_focus="0 0 0 3px rgba(34,211,238,0.22)",
    input_placeholder_color="#6c789f",

    button_primary_background_fill="linear-gradient(110deg, #7c5cff 0%, #c04dff 50%, #ff5fa2 100%)",
    button_primary_background_fill_hover="linear-gradient(110deg, #8f74ff 0%, #d066ff 50%, #ff7ab3 100%)",
    button_primary_text_color="#ffffff",
    button_primary_border_color="transparent",
    button_primary_shadow="0 10px 30px -8px rgba(192,77,255,0.65)",

    button_secondary_background_fill="rgba(255,255,255,0.06)",
    button_secondary_background_fill_hover="rgba(255,255,255,0.12)",
    button_secondary_border_color="rgba(255,255,255,0.18)",
    button_secondary_text_color="#e8ecff",

    link_text_color="#67e8f9",
    color_accent_soft="rgba(124,92,255,0.18)",
    border_color_primary="rgba(255,255,255,0.10)",
    table_even_background_fill="rgba(255,255,255,0.02)",
    table_odd_background_fill="rgba(255,255,255,0.05)",
    table_border_color="rgba(255,255,255,0.10)",
)

THEME = gr.themes.Soft(
    primary_hue=AURORA_VIOLET,
    secondary_hue=AURORA_AQUA,
    neutral_hue=MIDNIGHT,
    radius_size=gr.themes.sizes.radius_xxl,
    spacing_size=gr.themes.sizes.spacing_md,
    font=[gr.themes.GoogleFont("Manrope"), "ui-sans-serif", "system-ui", "sans-serif"],
    font_mono=[gr.themes.GoogleFont("JetBrains Mono"), "ui-monospace", "monospace"],
).set(
    **_AURORA_SETTINGS,
    **{f"{k}_dark": v for k, v in _AURORA_SETTINGS.items()},
)

CSS = """
@import url('https://fonts.googleapis.com/css2?family=Sora:wght@500;600;700;800&display=swap');

:root {
    --au-bg: #070a16;
    --au-violet: #7c5cff;
    --au-magenta: #ff5fa2;
    --au-aqua: #22d3ee;
    --au-mint: #5eead4;
    --au-text: #e8ecff;
    --au-muted: #9aa6cf;
    --au-glass: rgba(255, 255, 255, 0.045);
    --au-edge: rgba(255, 255, 255, 0.10);
    --au-head: 'Sora', 'Manrope', ui-sans-serif, system-ui, sans-serif;
    --au-grad: linear-gradient(110deg, #7c5cff 0%, #c04dff 45%, #ff5fa2 100%);
    --au-grad-cool: linear-gradient(110deg, #22d3ee 0%, #7c5cff 100%);
}

/* ── Canvas: midnight with aurora glows ─────────────────────────────── */
body, gradio-app, .gradio-container {
    background-color: var(--au-bg) !important;
    color: var(--au-text);
}
body {
    background-image:
        radial-gradient(900px 600px at 12% -8%, rgba(124, 92, 255, 0.34), transparent 60%),
        radial-gradient(800px 560px at 96% 4%, rgba(34, 211, 238, 0.20), transparent 62%),
        radial-gradient(900px 640px at 70% 108%, rgba(255, 95, 162, 0.22), transparent 60%) !important;
    background-attachment: fixed !important;
}
gradio-app, .gradio-container { background-image: none !important; }
.gradio-container { max-width: 1280px !important; margin: 0 auto !important; padding-top: 22px !important; }

/* ── Frosted glass panels ───────────────────────────────────────────── */
.gradio-container .block,
.gradio-container .form,
.gradio-container .gr-group {
    background: var(--au-glass) !important;
    border: 1px solid var(--au-edge) !important;
    border-radius: 20px !important;
    backdrop-filter: blur(14px) saturate(140%);
    -webkit-backdrop-filter: blur(14px) saturate(140%);
}
.gradio-container .block.padded,
.gradio-container .html-container { background: transparent !important; border: none !important; backdrop-filter: none; }

/* ── Hero ───────────────────────────────────────────────────────────── */
.hero {
    position: relative;
    overflow: hidden;
    border-radius: 28px;
    padding: 54px 52px 48px;
    margin-bottom: 22px;
    background:
        linear-gradient(135deg, rgba(124, 92, 255, 0.22) 0%, rgba(255, 95, 162, 0.10) 55%, rgba(34, 211, 238, 0.14) 100%),
        rgba(255, 255, 255, 0.04);
    border: 1px solid rgba(255, 255, 255, 0.14);
    box-shadow:
        0 30px 80px -30px rgba(124, 92, 255, 0.55),
        inset 0 1px 0 rgba(255, 255, 255, 0.18);
    backdrop-filter: blur(18px);
    -webkit-backdrop-filter: blur(18px);
}
/* drifting glow orbs: the one orchestrated moment of motion */
.hero::before, .hero::after {
    content: "";
    position: absolute;
    border-radius: 50%;
    filter: blur(60px);
    pointer-events: none;
    animation: drift 14s ease-in-out infinite alternate;
}
.hero::before {
    width: 340px; height: 340px; right: -80px; top: -120px;
    background: radial-gradient(circle, rgba(34, 211, 238, 0.65), transparent 70%);
}
.hero::after {
    width: 300px; height: 300px; right: 180px; bottom: -170px;
    background: radial-gradient(circle, rgba(255, 95, 162, 0.55), transparent 70%);
    animation-duration: 18s;
}
@keyframes drift {
    from { transform: translate3d(0, 0, 0) scale(1); }
    to   { transform: translate3d(-40px, 30px, 0) scale(1.18); }
}
.hero-inner { position: relative; z-index: 1; }
.hero h1 {
    margin: 0 0 14px 0;
    font-family: var(--au-head);
    font-size: clamp(2rem, 4.2vw, 3.3rem);
    font-weight: 800;
    line-height: 1.06;
    letter-spacing: -0.035em;
    background: linear-gradient(100deg, #ffffff 0%, #e3d9ff 45%, #a9f1ff 100%);
    -webkit-background-clip: text;
            background-clip: text;
    -webkit-text-fill-color: transparent;
            color: transparent;
}
.hero p {
    margin: 0;
    font-size: 1.06rem;
    line-height: 1.7;
    max-width: 660px;
    color: #c3cdf0;
}
.chips { margin-top: 26px; display: flex; flex-wrap: wrap; gap: 9px; }
.chip {
    padding: 6px 15px;
    border-radius: 999px;
    font-size: 0.8rem;
    font-weight: 600;
    color: #e8ecff;
    background: rgba(255, 255, 255, 0.08);
    border: 1px solid rgba(255, 255, 255, 0.18);
    backdrop-filter: blur(6px);
    transition: background .2s ease, border-color .2s ease, box-shadow .2s ease;
}
.chip:hover {
    background: rgba(124, 92, 255, 0.28);
    border-color: rgba(173, 148, 255, 0.8);
    box-shadow: 0 0 22px -4px rgba(124, 92, 255, 0.8);
}

/* ── Section titles ─────────────────────────────────────────────────── */
.section-title {
    font-family: var(--au-head);
    font-weight: 700;
    font-size: 1.02rem;
    letter-spacing: -0.01em;
    margin: 8px 0 8px;
    color: #f2f4ff;
    display: flex; align-items: center; gap: 10px;
}
.section-title::before {
    content: "";
    width: 10px; height: 10px; border-radius: 50%;
    background: var(--au-grad);
    box-shadow: 0 0 14px 2px rgba(192, 77, 255, 0.7);
}

/* ── Status ─────────────────────────────────────────────────────────── */
.status {
    padding: 14px 18px;
    border-radius: 16px;
    font-size: 0.93rem;
    border: 1px solid transparent;
    backdrop-filter: blur(10px);
}
.status-idle {
    background: rgba(255, 255, 255, 0.05);
    border-color: rgba(255, 255, 255, 0.12);
    color: #b9c4ea;
}
.status-ok {
    background: rgba(94, 234, 212, 0.10);
    border-color: rgba(94, 234, 212, 0.45);
    color: #7ff3e0;
    font-weight: 700;
    box-shadow: 0 0 28px -10px rgba(94, 234, 212, 0.7);
}

/* ── Stats ──────────────────────────────────────────────────────────── */
.stats {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(138px, 1fr));
    gap: 12px;
    margin: 4px 0 12px;
}
.stat {
    position: relative;
    overflow: hidden;
    padding: 18px 16px 14px;
    border-radius: 18px;
    background: linear-gradient(160deg, rgba(255, 255, 255, 0.09), rgba(255, 255, 255, 0.03));
    border: 1px solid rgba(255, 255, 255, 0.13);
    box-shadow: 0 16px 34px -20px rgba(0, 0, 0, 0.9);
}
.stat::after {
    content: "";
    position: absolute; left: 0; right: 0; bottom: 0; height: 3px;
    background: var(--au-grad-cool);
    opacity: 0.85;
}
.stat-value {
    font-family: var(--au-head);
    font-size: 1.3rem;
    font-weight: 800;
    letter-spacing: -0.02em;
    word-break: break-word;
    background: var(--au-grad-cool);
    -webkit-background-clip: text;
            background-clip: text;
    -webkit-text-fill-color: transparent;
            color: transparent;
}
.stat-label { font-size: 0.78rem; margin-top: 4px; color: var(--au-muted); font-weight: 600; }

/* ── Tabs as pills ──────────────────────────────────────────────────── */
.tab-wrapper, .tab-container { border: none !important; }
.tab-container::after { display: none !important; }
button[role="tab"] {
    font-family: var(--au-head) !important;
    font-weight: 600 !important;
    font-size: 0.9rem !important;
    color: var(--au-muted) !important;
    background: transparent !important;
    border: 1px solid transparent !important;
    border-radius: 999px !important;
    padding: 8px 18px !important;
    transition: color .18s ease, background .18s ease;
}
button[role="tab"]:hover { color: #fff !important; background: rgba(255, 255, 255, 0.07) !important; }
button[role="tab"][aria-selected="true"] {
    color: #fff !important;
    background: linear-gradient(110deg, rgba(124, 92, 255, 0.55), rgba(255, 95, 162, 0.40)) !important;
    border-color: rgba(255, 255, 255, 0.22) !important;
    box-shadow: 0 8px 22px -10px rgba(192, 77, 255, 0.9);
}
.tabitem { border: none !important; background: transparent !important; }

/* ── Inputs, upload, tables, code ───────────────────────────────────── */
.gradio-container textarea, .gradio-container input[type="text"] { color: var(--au-text) !important; }
.gradio-container label span, .gradio-container .block-label { color: #cbbcff !important; font-weight: 600; }

.gradio-container [data-testid="file-upload-dropzone"],
.gradio-container .upload-container button,
.gradio-container .file-preview-holder {
    border: 2px dashed rgba(173, 148, 255, 0.45) !important;
    border-radius: 20px !important;
    background:
        radial-gradient(circle at 50% 0%, rgba(124, 92, 255, 0.20), transparent 70%),
        rgba(255, 255, 255, 0.02) !important;
    color: #c3cdf0 !important;
    transition: border-color .2s ease, box-shadow .2s ease;
}
.gradio-container [data-testid="file-upload-dropzone"]:hover,
.gradio-container .upload-container button:hover {
    border-color: var(--au-aqua) !important;
    box-shadow: 0 0 36px -10px rgba(34, 211, 238, 0.7);
}

.gradio-container table { font-size: 0.92rem; color: var(--au-text); }
.gradio-container table thead th {
    font-family: var(--au-head);
    background: rgba(124, 92, 255, 0.22) !important;
    color: #f2f0ff !important;
}
.gradio-container table td, .gradio-container table th { border-color: rgba(255, 255, 255, 0.08) !important; }

.prose, .md, .prose p, .prose li, .md p, .md li { color: #d5dcf8 !important; }
.prose h1, .prose h2, .prose h3, .md h1, .md h2, .md h3 {
    font-family: var(--au-head);
    letter-spacing: -0.02em;
    color: #ffffff !important;
}
.prose h2, .md h2 {
    padding-bottom: 8px;
    border-bottom: 1px solid rgba(255, 255, 255, 0.14);
}
.prose strong, .md strong { color: #fff; }
.prose hr, .md hr { border: 0; border-top: 1px solid rgba(255, 255, 255, 0.12); }
.prose a, .md a { color: var(--au-aqua) !important; }
.prose code, .md code { background: rgba(124, 92, 255, 0.2); border-radius: 6px; padding: 1px 6px; }

.gradio-container .cm-editor, .gradio-container .cm-gutters {
    background: rgba(5, 7, 16, 0.85) !important;
    border-radius: 14px;
}

/* ── Buttons ────────────────────────────────────────────────────────── */
button.primary {
    font-family: var(--au-head) !important;
    font-weight: 700 !important;
    font-size: 1rem !important;
    letter-spacing: 0.005em;
    border: none !important;
    border-radius: 16px !important;
    transition: transform .18s ease, box-shadow .18s ease, filter .18s ease;
}
button.primary:hover {
    transform: translateY(-2px);
    box-shadow: 0 16px 38px -8px rgba(255, 95, 162, 0.75) !important;
    filter: brightness(1.06);
}
button.primary:active { transform: translateY(0); }
button.secondary {
    font-weight: 600 !important;
    border-radius: 16px !important;
    backdrop-filter: blur(8px);
}

/* ── Footer, scrollbars ─────────────────────────────────────────────── */
.footer {
    text-align: center; font-size: 0.82rem; margin-top: 22px;
    color: var(--au-muted);
}
footer { display: none !important; }
* { scrollbar-width: thin; scrollbar-color: rgba(173, 148, 255, 0.5) transparent; }

/* ── Responsive & accessibility ─────────────────────────────────────── */
@media (max-width: 760px) {
    .hero { padding: 34px 24px; border-radius: 22px; }
}
:focus-visible { outline: 2px solid var(--au-aqua); outline-offset: 2px; }
@media (prefers-reduced-motion: reduce) {
    .hero::before, .hero::after { animation: none; }
    button.primary, .chip { transition: none; }
}
"""

HERO_HTML = """
<div class="hero">
  <div class="hero-inner">
    <h1>Doc2JSON</h1>
    <p>Universal document → JSON extractor. Upload any document and convert it into clean, structured, schema-validated
    JSON — with a readable overview, key facts, entities, and sections.</p>
    <div class="chips">
      <span class="chip">PDF</span><span class="chip">Word</span>
      <span class="chip">PowerPoint</span><span class="chip">Excel</span>
      <span class="chip">RTF</span><span class="chip">HTML</span>
      <span class="chip">Markdown &amp; Code</span><span class="chip">Pydantic validated</span>
    </div>
  </div>
</div>
"""


# ──────────────────────────────────────────────────────────────────────────
# UI
# ──────────────────────────────────────────────────────────────────────────

with gr.Blocks(title="Doc2JSON") as app:

    gr.HTML(HERO_HTML)

    with gr.Row(equal_height=False):
        # ── Left column: input ──────────────────────────────────────────
        with gr.Column(scale=4, min_width=340):
            gr.HTML('<div class="section-title">1 · Upload</div>')
            file_input = gr.File(
                label="Document",
                file_types=UPLOAD_TYPES,
                type="filepath",
            )

            gr.HTML('<div class="section-title">2 · Customize (optional)</div>')
            extra_instructions = gr.Textbox(
                label="Extraction instructions",
                placeholder="e.g. Focus on the research methods and findings.",
                lines=3,
            )
            gr.Examples(
                examples=[
                    ["Focus on the research methods and findings."],
                    ["Extract all dates, amounts, and parties involved."],
                    ["List every skill, job title, and employer."],
                ],
                inputs=[extra_instructions],
                label="Example instructions",
            )

            with gr.Row():
                extract_button = gr.Button("Extract to JSON", variant="primary", scale=3)
                clear_button = gr.Button("Clear", variant="secondary", scale=1)

            status_output = gr.HTML(EMPTY_STATUS)

        # ── Right column: results ───────────────────────────────────────
        with gr.Column(scale=7, min_width=420):
            stats_output = gr.HTML()

            with gr.Tabs():
                with gr.Tab("Overview"):
                    overview_output = gr.Markdown()

                with gr.Tab("Key information"):
                    key_output = gr.Dataframe(
                        headers=["Key", "Value"],
                        datatype=["str", "str"],
                        interactive=False,
                        wrap=True,
                    )

                with gr.Tab("Entities"):
                    entities_output = gr.Dataframe(
                        headers=["Type", "Value"],
                        datatype=["str", "str"],
                        interactive=False,
                        wrap=True,
                    )

                with gr.Tab("Sections"):
                    sections_output = gr.Markdown()

                with gr.Tab("JSON"):
                    json_output = gr.Code(
                        label="Validated JSON",
                        language="json",
                        interactive=False,
                    )
                    download_output = gr.File(label="Download JSON", interactive=False)

    gr.HTML(
        '<div class="footer">Your document text is sent to the Gemini API for '
        "processing. Review the output before relying on it.</div>"
    )

    outputs = [
        status_output,
        stats_output,
        overview_output,
        key_output,
        entities_output,
        sections_output,
        json_output,
        download_output,
    ]

    extract_button.click(
        fn=extract_document,
        inputs=[file_input, extra_instructions],
        outputs=outputs,
        show_progress="full",
    )

    clear_button.click(
        fn=clear_all,
        inputs=None,
        outputs=[file_input, extra_instructions, *outputs],
    )


if __name__ == "__main__":
    app.queue(max_size=16).launch(show_error=True, theme=THEME, css=CSS)