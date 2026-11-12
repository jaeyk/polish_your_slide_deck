from pathlib import Path
import io
import math
import re
from typing import Literal

import pymupdf as fitz
from mcp.server.fastmcp import FastMCP
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.dml import MSO_FILL
from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE, PP_PLACEHOLDER
from pptx.enum.text import MSO_ANCHOR
from pptx.oxml.ns import qn
from pptx.text.text import _Run
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
OUTPUT_DIR = ROOT / "output"
PREFERENCES_PATH = ROOT / "design-preferences.md"

# The user's chosen default style (see design-preferences.md).
DEFAULT_STYLE = "newsroom_light"
PROFILES = {
    # Cool white, near-black text, Carolina Navy emphasis blocks, in Overpass
    # (based on Highway Gothic, the US road-sign face). White screenshots and
    # charts blend into the background.
    "newsroom_light": {
        "font_family": "Overpass",
        "title_color": RGBColor(0x14, 0x1A, 0x22),
        "body_color": RGBColor(0x14, 0x1A, 0x22),
        "accent_color": RGBColor(0x13, 0x29, 0x4B),
        "background": RGBColor(0xF5, 0xF7, 0xFA),
    },
    # Deep navy, white Public Sans, signal-yellow emphasis. Public Sans is the
    # US government's USWDS typeface; install it locally so PowerPoint uses it.
    "civic_signal": {
        "font_family": "Public Sans",
        "title_color": RGBColor(0xFF, 0xFF, 0xFF),
        "body_color": RGBColor(0xFF, 0xFF, 0xFF),
        "accent_color": RGBColor(0xFF, 0xD3, 0x4E),
        "background": RGBColor(0x0B, 0x1A, 0x2E),
    },
    "academic": {
        "font_family": "Georgia",
        "title_color": RGBColor(35, 54, 74),
        "body_color": RGBColor(47, 55, 63),
        "accent_color": RGBColor(105, 78, 82),
        "background": None,
    },
    "professional": {
        "font_family": "Aptos",
        "title_color": RGBColor(31, 45, 61),
        "body_color": RGBColor(55, 65, 75),
        "accent_color": RGBColor(0, 112, 112),
        "background": None,
    },
}

# Large-type floors from design-preferences.md (audiences tune out on small text).
TITLE_MIN_PT = 40
BODY_MIN_PT = 24
LABEL_MIN_PT = 18
# Title placeholders this long are really captions, so they get the body floor.
CAPTION_TITLE_MIN_WORDS = 12
MIN_CONTRAST = 7.0
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
# Pictures with more transparent area than this get a white plate behind them,
# so dark lines (e.g. chart axes) stay visible on a dark background.
TRANSPARENT_SHARE_FOR_PLATE = 0.10
# Balanced slides sit this fraction of the slide height above true center.
OPTICAL_LIFT = 0.022

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


def _luminance(rgb):
    channels = []
    for value in rgb:
        value = value / 255
        channels.append(value / 12.92 if value <= 0.03928 else ((value + 0.055) / 1.055) ** 2.4)
    return 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2]


def _contrast(first, second):
    lighter, darker = sorted((_luminance(first), _luminance(second)), reverse=True)
    return (lighter + 0.05) / (darker + 0.05)


def _all_runs(paragraph):
    """Text runs plus fields (e.g. slide numbers), which python-pptx's runs skip."""
    for element in paragraph._p:
        if element.tag in (qn("a:r"), qn("a:fld")):
            yield _Run(element, paragraph)


def _text_role(shape):
    placeholder_type = _placeholder_type(shape)
    text = shape.text_frame.text.strip()
    if placeholder_type in {
        PP_PLACEHOLDER.SLIDE_NUMBER,
        PP_PLACEHOLDER.FOOTER,
        PP_PLACEHOLDER.DATE,
    } or re.match(r"(source|note)s?\s*:", text, re.IGNORECASE):
        return "label"
    if placeholder_type in {PP_PLACEHOLDER.TITLE, PP_PLACEHOLDER.CENTER_TITLE}:
        return "caption" if len(text.split()) >= CAPTION_TITLE_MIN_WORDS else "title"
    if placeholder_type == PP_PLACEHOLDER.SUBTITLE:
        return "subtitle"
    return "body"


def _size_floor(role):
    return {"title": TITLE_MIN_PT, "label": LABEL_MIN_PT}.get(role, BODY_MIN_PT)


def _is_white_highlight(highlight):
    """White highlights were invisible on the original light slides."""
    color = highlight[0] if len(highlight) else None
    if color is None:
        return True
    if color.tag == qn("a:schemeClr"):
        return color.get("val") in {"lt1", "bg1"}
    if color.tag == qn("a:srgbClr"):
        value = color.get("val", "FFFFFF")
        return _luminance(tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))) > 0.85
    return False


def _set_highlight(run, fill):
    """Recolor a visible text highlight to fill; return True if one remains.

    White highlights are removed, since they only show up on a dark background.
    """
    rPr = run._r.find(qn("a:rPr"))
    highlight = rPr.find(qn("a:highlight")) if rPr is not None else None
    if highlight is None:
        return False
    if _is_white_highlight(highlight):
        rPr.remove(highlight)
        return False
    for child in list(highlight):
        highlight.remove(child)
    color = highlight.makeelement(qn("a:srgbClr"), {"val": str(fill)})
    highlight.append(color)
    return True


def _shape_background(shape, slide_background):
    try:
        if shape.fill.type == MSO_FILL.SOLID:
            return shape.fill.fore_color.rgb
    except (AttributeError, TypeError, ValueError):
        pass
    return slide_background


def _readable(color, background):
    """Keep color if it meets MIN_CONTRAST on background, else pick white or the background's opposite."""
    if background is None or _contrast(color, background) >= MIN_CONTRAST:
        return color
    if _contrast(WHITE, background) >= MIN_CONTRAST:
        return WHITE
    return RGBColor(0x0B, 0x1A, 0x2E)


def _style_shape(shape, font_family, title_color, body_color, accent_color, background):
    """Apply font, palette, highlights, contrast, and size floors to one text shape.

    Returns the role and whether any run was enlarged, for overflow checks.
    """
    role = _text_role(shape)
    shape_background = _shape_background(shape, background)
    role_color = {"title": title_color, "subtitle": accent_color}.get(role, body_color)
    floor = _size_floor(role)
    enlarged = False

    for paragraph in shape.text_frame.paragraphs:
        for run in _all_runs(paragraph):
            run.font.name = font_family
            if background is not None and _set_highlight(run, accent_color):
                # Inverted block: accent highlight with background-colored text.
                run.font.color.rgb = background
            elif _is_highlight_color(run.font.color):
                # Keep deliberate emphasis colors unless they are unreadable.
                run.font.color.rgb = _readable(run.font.color.rgb, shape_background)
            else:
                run.font.color.rgb = _readable(role_color, shape_background)
            size = run.font.size
            if size is None or size.pt < floor:
                run.font.size = Pt(floor)
                enlarged = enlarged or size is not None
    return role, enlarged


def _is_highlight_color(color):
    """True for an explicit, saturated run color (e.g. red emphasis)."""
    try:
        rgb = color.rgb
    except AttributeError:
        return False
    if rgb is None:
        return False
    return max(rgb) - min(rgb) > 60


def _estimated_text_height_pt(shape):
    """Rough rendered height: average glyph width ~0.52 em, line height 1.2."""
    usable_width = max(shape.width / 12700 - 14.4, 1)
    height = 0.0
    paragraphs = list(shape.text_frame.paragraphs)
    # Trailing empty paragraphs don't render as visible lines.
    while paragraphs and not "".join(run.text for run in _all_runs(paragraphs[-1])).strip():
        paragraphs.pop()
    for paragraph in paragraphs:
        sizes = [run.font.size.pt for run in _all_runs(paragraph) if run.font.size]
        size = max(sizes, default=BODY_MIN_PT)
        text = "".join(run.text for run in _all_runs(paragraph))
        lines = max(1, math.ceil(len(text) * 0.52 * size / usable_width))
        height += lines * size * 1.2
    return height + 7.2


def _overlaps(a, b):
    return a[0] < b[2] and b[0] < a[2] and a[1] < b[3] and b[1] < a[3]


def _overflow_warning(slide_number, shape, slide, slide_height_pt):
    """Warn when enlarged text runs off the slide or into a picture it didn't touch before."""
    left, top = shape.left / 12700, shape.top / 12700
    right = left + shape.width / 12700
    original = (left, top, right, top + shape.height / 12700)
    grown = (left, top, right, top + _estimated_text_height_pt(shape))
    snippet = shape.text_frame.text.strip().replace("\n", " ")[:50]
    if grown[3] > slide_height_pt:
        return f"Slide {slide_number}: '{snippet}' may run off the bottom of the slide"
    for picture in slide.shapes:
        if picture.shape_type != MSO_SHAPE_TYPE.PICTURE:
            continue
        box = (
            picture.left / 12700,
            picture.top / 12700,
            (picture.left + picture.width) / 12700,
            (picture.top + picture.height) / 12700,
        )
        if _overlaps(grown, box) and not _overlaps(original, box):
            return f"Slide {slide_number}: '{snippet}' may now overlap a picture"
    return None


def _style_lines(shapes, accent_color):
    """Recolor outlines (e.g. callout rings) on autoshapes to the accent color."""
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            _style_lines(shape.shapes, accent_color)
        elif shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
            if shape.line.fill.type == MSO_FILL.SOLID:
                shape.line.color.rgb = accent_color


def _transparent_share(picture):
    try:
        image = Image.open(io.BytesIO(picture.image.blob))
    except Exception:
        return 0.0
    if image.mode not in ("RGBA", "LA") and "transparency" not in image.info:
        return 0.0
    histogram = image.convert("RGBA").getchannel("A").histogram()
    return sum(histogram[:128]) / max(sum(histogram), 1)


def _balance_vertically(slide, slide_height_pt):
    """Center a text-only slide's content slightly above the middle.

    Slides with pictures, tables, charts, or drawn shapes keep their layout.
    Returns True when the slide was moved.
    """
    content = []
    for shape in slide.shapes:
        if not shape.has_text_frame:
            return False
        if _text_role(shape) == "label" or not shape.text_frame.text.strip():
            continue
        content.append(shape)
    if not content:
        return False
    for shape in content:
        shape.text_frame.vertical_anchor = MSO_ANCHOR.TOP
    tops = [shape.top / 12700 for shape in content]
    bottoms = [top + _estimated_text_height_pt(shape) for top, shape in zip(tops, content)]
    height = max(bottoms) - min(tops)
    target_top = (slide_height_pt - height) / 2 - slide_height_pt * OPTICAL_LIFT
    shift = target_top - min(tops)
    if abs(shift) < 6:
        return False
    for shape in content:
        shape.top = int(shape.top + shift * 12700)
    return True


def _plate_transparent_pictures(slide):
    """Put a white plate behind mostly transparent pictures; return how many."""
    pictures = [
        shape
        for shape in slide.shapes
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE
        and _transparent_share(shape) > TRANSPARENT_SHARE_FOR_PLATE
    ]
    for picture in pictures:
        plate = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, picture.left, picture.top, picture.width, picture.height
        )
        plate.name = "Polish plate"
        plate.fill.solid()
        plate.fill.fore_color.rgb = WHITE
        plate.line.fill.background()
        plate.shadow.inherit = False
        picture._element.addprevious(plate._element)
    return len(pictures)


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
            # Rewriting text drops run formatting and fields (e.g. slide numbers),
            # so only touch shapes whose text actually changed.
            if text != shape.text_frame.text.strip("\n"):
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
def get_design_preferences() -> str:
    """Return the user's slide design preferences. Read before proposing or applying any style."""
    if not PREFERENCES_PATH.is_file():
        raise FileNotFoundError("No design-preferences.md found at the project root.")
    return PREFERENCES_PATH.read_text(encoding="utf-8")


@mcp.tool()
def polish_deck(
    filename: str,
    audience: Literal["newsroom_light", "civic_signal", "academic", "professional"] = DEFAULT_STYLE,
    primary_color: str | None = None,
    accent_color: str | None = None,
    font_family: str | None = None,
) -> str:
    """Polish a deck from data/ and save its copy in output/.

    `audience` picks the style profile. The default, newsroom_light, is the
    user's chosen style: cool white background, near-black text, and Carolina
    Navy emphasis blocks. civic_signal is its dark counterpart. Every profile
    enforces the large-type floors (titles 40 pt, body 24 pt, labels 18 pt)
    and vertically balances text-only slides, per design-preferences.md.
    """
    source = _source_document(filename)
    profile = PROFILES[audience]
    selected_font = (font_family or profile["font_family"]).strip()
    if not selected_font or len(selected_font) > 64:
        raise ValueError("font_family must contain between 1 and 64 characters.")

    title_color = _parse_color(primary_color, profile["title_color"])
    body_color = profile["body_color"]
    selected_accent = _parse_color(accent_color, profile["accent_color"])
    background = profile["background"]

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

    slide_height_pt = presentation.slide_height / 12700
    warnings = []
    plates = 0
    balanced = 0
    for slide_number, slide in enumerate(presentation.slides, start=1):
        if background is not None:
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = background
            _style_lines(slide.shapes, selected_accent)
            if _luminance(background) < 0.5:
                plates += _plate_transparent_pictures(slide)
        for shape in _text_shapes(slide.shapes):
            _, enlarged = _style_shape(
                shape,
                selected_font,
                title_color,
                body_color,
                selected_accent,
                background,
            )
            if enlarged:
                warning = _overflow_warning(slide_number, shape, slide, slide_height_pt)
                if warning:
                    warnings.append(warning)
        if _balance_vertically(slide, slide_height_pt):
            balanced += 1
    presentation.save(str(destination))

    summary = (
        f"Polished {len(presentation.slides)} slides with the {audience} style: "
        f"{destination.relative_to(ROOT)}"
    )
    if balanced:
        summary += f"\nVertically balanced {balanced} text-only slide(s)."
    if plates:
        summary += f"\nAdded white plates behind {plates} transparent picture(s)."
    if warnings:
        summary += "\nCheck these slides (text was enlarged to the size floors):\n- " + "\n- ".join(warnings)
    return summary


if __name__ == "__main__":
    mcp.run()
