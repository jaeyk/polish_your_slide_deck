from pathlib import Path
import re
from typing import Literal

import pymupdf as fitz
from mcp.server.fastmcp import FastMCP
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"
PROFILES = {
    "academic": {
        "font_family": "Georgia",
        "title_color": RGBColor(35, 54, 74),
        "body_color": RGBColor(47, 55, 63),
        "accent_color": RGBColor(105, 78, 82),
    },
    "professional": {
        "font_family": "Aptos",
        "title_color": RGBColor(31, 45, 61),
        "body_color": RGBColor(55, 65, 75),
        "accent_color": RGBColor(0, 112, 112),
    },
}

mcp = FastMCP("slide-deck-polisher")


def _text_shapes(shapes):
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _text_shapes(shape.shapes)
        elif shape.has_text_frame:
            yield shape


def _placeholder_type(shape):
    if not shape.is_placeholder:
        return None
    return shape.placeholder_format.type


def _parse_color(value, fallback):
    if value is None:
        return fallback
    if not re.fullmatch(r"#?[0-9A-Fa-f]{6}", value.strip()):
        raise ValueError(f"Invalid hex color '{value}'. Use a value such as '#1F2D3D'.")
    hex_value = value.strip().lstrip("#")
    return RGBColor(*(int(hex_value[index : index + 2], 16) for index in (0, 2, 4)))


def _style_shape(shape, font_family, title_color, body_color, accent_color):
    placeholder_type = _placeholder_type(shape)
    is_title = placeholder_type in {
        PP_PLACEHOLDER.TITLE,
        PP_PLACEHOLDER.CENTER_TITLE,
    }
    is_subtitle = placeholder_type == PP_PLACEHOLDER.SUBTITLE

    for paragraph in shape.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_family
            if is_title:
                run.font.size = Pt(30)
                run.font.bold = True
                run.font.color.rgb = title_color
            elif is_subtitle:
                run.font.size = Pt(20)
                run.font.color.rgb = accent_color
            else:
                if run.font.size is None:
                    run.font.size = Pt(18)
                run.font.color.rgb = body_color


def _source_document(filename):
    if Path(filename).name != filename or Path(filename).suffix.lower() not in {
        ".pptx",
        ".pdf",
    }:
        raise ValueError("Provide the filename of a .pptx or .pdf file in the data folder.")

    source = (DATA_DIR / filename).resolve()
    if source.parent != DATA_DIR.resolve() or not source.is_file():
        raise FileNotFoundError(
            f"No presentation named '{filename}' exists in the data folder."
        )
    return source


def _create_pdf_draft(source, destination):
    presentation = Presentation()
    layout = presentation.slide_layouts[1]
    with fitz.open(source) as pdf:
        for page_number, page in enumerate(pdf, start=1):
            page_text = page.get_text("text").strip()
            if not page_text:
                raise ValueError(
                    f"PDF page {page_number} has no selectable text. Scanned PDFs need OCR before import."
                )
            lines = [line.strip() for line in page_text.splitlines() if line.strip()]
            title = lines[0] if lines else f"Page {page_number}"
            body = "\n".join(lines[1:])
            slide = presentation.slides.add_slide(layout)
            slide.shapes.title.text = title
            body_shape = next(
                (
                    shape
                    for shape in slide.placeholders
                    if shape != slide.shapes.title and shape.has_text_frame
                ),
                None,
            )
            if body_shape is None:
                body_shape = slide.shapes.add_textbox(
                    Inches(0.7), Inches(1.6), Inches(12), Inches(5.2)
                )
            body_shape.text = body
    destination.parent.mkdir(parents=True, exist_ok=True)
    presentation.save(str(destination))
    return destination


def _editable_pptx_path(source, refresh_pdf_draft=False):
    if source.suffix.lower() == ".pptx":
        return source
    draft = OUTPUT_DIR / f"{source.stem}-draft.pptx"
    if refresh_pdf_draft or not draft.is_file():
        _create_pdf_draft(source, draft)
    return draft


def _role_for_shape(shape):
    placeholder_type = _placeholder_type(shape)
    if placeholder_type in {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}:
        return "title"
    if placeholder_type == PP_PLACEHOLDER.SUBTITLE:
        return "subtitle"
    return "text"


def _escape_markdown_text(text):
    escaped_lines = []
    for line in text.splitlines():
        if re.fullmatch(r"## Slide \d+", line) or re.fullmatch(
            r"<!-- pptx-shape: \d+(?:; role: [^>]+)? -->", line
        ):
            line = "\\" + line
        escaped_lines.append(line)
    return "\n".join(escaped_lines)


def _parse_markdown_edits(markdown):
    slides = {}
    current_slide = None
    current_shape = None
    content_lines = []

    def save_shape():
        if current_slide is not None and current_shape is not None:
            slides.setdefault(current_slide, {})[current_shape] = "\n".join(
                content_lines
            ).strip("\n")

    for line in markdown.splitlines():
        slide_match = re.fullmatch(r"## Slide (\d+)", line)
        shape_match = re.fullmatch(
            r"<!-- pptx-shape: (\d+)(?:; role: [^>]+)? -->", line
        )
        if slide_match:
            save_shape()
            current_slide = int(slide_match.group(1))
            current_shape = None
            content_lines = []
        elif shape_match and current_slide is not None:
            save_shape()
            current_shape = int(shape_match.group(1))
            content_lines = []
        elif current_shape is not None:
            if line.startswith("\\") and (
                re.fullmatch(r"## Slide \d+", line[1:])
                or re.fullmatch(
                    r"<!-- pptx-shape: \d+(?:; role: [^>]+)? -->", line[1:]
                )
            ):
                line = line[1:]
            content_lines.append(line)
    save_shape()
    return slides


def _apply_markdown_edits(presentation, markdown_path):
    edits = _parse_markdown_edits(markdown_path.read_text(encoding="utf-8"))
    for slide_number, shape_edits in edits.items():
        if not 1 <= slide_number <= len(presentation.slides):
            raise ValueError(f"Markdown references missing slide {slide_number}.")
        slide = presentation.slides[slide_number - 1]
        shapes_by_id = {shape.shape_id: shape for shape in _text_shapes(slide.shapes)}
        for shape_id, text in shape_edits.items():
            shape = shapes_by_id.get(shape_id)
            if shape is None:
                raise ValueError(
                    f"Markdown references missing text shape {shape_id} on slide {slide_number}."
                )
            shape.text_frame.text = text


@mcp.tool()
def list_decks() -> list[str]:
    """List PowerPoint decks and PDFs available in the data folder."""
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    return sorted(
        path.name
        for path in DATA_DIR.iterdir()
        if path.is_file() and path.suffix.lower() in {".pptx", ".pdf"}
    )


@mcp.tool()
def export_deck_to_markdown(filename: str) -> str:
    """Export slide text or PDF page text to editable Markdown in output/."""
    source = _source_document(filename)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    destination = OUTPUT_DIR / f"{source.stem}.md"
    editable_source = _editable_pptx_path(
        source, refresh_pdf_draft=source.suffix.lower() == ".pdf"
    )
    presentation = Presentation(str(editable_source))
    lines = [
        f"# Editable text for {source.name}",
        "",
        "Edit text below each shape marker; leave marker comments unchanged.",
        "",
    ]
    for slide_number, slide in enumerate(presentation.slides, start=1):
        lines.append(f"## Slide {slide_number}")
        lines.append("")
        for shape in _text_shapes(slide.shapes):
            lines.append(
                f"<!-- pptx-shape: {shape.shape_id}; role: {_role_for_shape(shape)} -->"
            )
            lines.append(_escape_markdown_text(shape.text_frame.text))
            lines.append("")
    destination.write_text("\n".join(lines), encoding="utf-8")
    return f"Exported {len(presentation.slides)} slides to {destination.relative_to(ROOT)}"


@mcp.tool()
def apply_markdown_edits(filename: str) -> str:
    """Apply the edited output/<deck-name>.md text to a new editable PPTX copy."""
    source = _source_document(filename)
    markdown_path = OUTPUT_DIR / f"{source.stem}.md"
    if not markdown_path.is_file():
        raise FileNotFoundError(
            f"No editable Markdown file found at {markdown_path.relative_to(ROOT)}."
        )
    editable_source = _editable_pptx_path(source, refresh_pdf_draft=True)
    presentation = Presentation(str(editable_source))
    _apply_markdown_edits(presentation, markdown_path)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    destination = OUTPUT_DIR / f"{source.stem}-edited.pptx"
    presentation.save(str(destination))
    return f"Applied Markdown edits to {destination.relative_to(ROOT)}"


@mcp.tool()
def polish_deck(
    filename: str,
    audience: Literal["academic", "professional"] = "academic",
    primary_color: str | None = None,
    accent_color: str | None = None,
    font_family: str | None = None,
) -> str:
    """Polish a deck from data/ and save its copy in output/."""
    source = _source_document(filename)
    profile = PROFILES[audience]
    selected_font = (font_family or profile["font_family"]).strip()
    if not selected_font or len(selected_font) > 64:
        raise ValueError("font_family must contain between 1 and 64 characters.")

    title_color = _parse_color(primary_color, profile["title_color"])
    body_color = profile["body_color"]
    selected_accent = _parse_color(accent_color, profile["accent_color"])

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    destination = OUTPUT_DIR / f"{source.stem}-polished.pptx"
    editable_source = _editable_pptx_path(source, refresh_pdf_draft=True)
    presentation = Presentation(str(editable_source))
    markdown_path = OUTPUT_DIR / f"{source.stem}.md"
    edited_path = OUTPUT_DIR / f"{source.stem}-edited.pptx"
    if markdown_path.is_file():
        _apply_markdown_edits(presentation, markdown_path)
    elif edited_path.is_file():
        presentation = Presentation(str(edited_path))
    for slide in presentation.slides:
        for shape in _text_shapes(slide.shapes):
            _style_shape(
                shape,
                selected_font,
                title_color,
                body_color,
                selected_accent,
            )
    presentation.save(str(destination))

    return (
        f"Polished {len(presentation.slides)} slides for a {audience} audience: "
        f"{destination.relative_to(ROOT)}"
    )


if __name__ == "__main__":
    mcp.run()
