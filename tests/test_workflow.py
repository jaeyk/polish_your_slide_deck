import io
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pymupdf as fitz
from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Inches, Pt

import server


class WorkflowTests(unittest.TestCase):
    def setUp(self):
        self.temporary_directory = tempfile.TemporaryDirectory()
        self.root = Path(self.temporary_directory.name)
        self.data_directory = self.root / "data"
        self.output_directory = self.root / "output"
        self.data_directory.mkdir()
        self.output_directory.mkdir()
        self.path_patch = patch.multiple(
            server,
            ROOT=self.root,
            DATA_DIR=self.data_directory,
            OUTPUT_DIR=self.output_directory,
        )
        self.path_patch.start()

    def tearDown(self):
        self.path_patch.stop()
        self.temporary_directory.cleanup()

    def _create_pptx(self):
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[1])
        slide.shapes.title.text = "Source title"
        body_shape = next(
            shape
            for shape in slide.placeholders
            if shape != slide.shapes.title and shape.has_text_frame
        )
        body_shape.text = "Source body"
        presentation.save(self.data_directory / "deck.pptx")

    def _create_pdf(self):
        document = fitz.open()
        page = document.new_page()
        page.insert_text((72, 72), "PDF heading\nFirst paragraph\nSecond paragraph")
        document.save(self.data_directory / "deck.pdf")
        document.close()

    def test_pptx_markdown_edits_are_applied_and_polished(self):
        self._create_pptx()
        server.export_deck_to_markdown("deck.pptx")
        markdown_path = self.output_directory / "deck.md"
        markdown = markdown_path.read_text(encoding="utf-8")
        markdown_path.write_text(
            markdown.replace("Source title", "Edited title").replace(
                "Source body", "Edited body"
            ),
            encoding="utf-8",
        )

        server.apply_markdown_edits("deck.pptx")
        server.polish_deck("deck.pptx", audience="academic")

        polished = Presentation(self.output_directory / "deck-polished.pptx")
        slide = polished.slides[0]
        self.assertEqual(slide.shapes.title.text, "Edited title")
        self.assertEqual(slide.placeholders[1].text, "Edited body")
        self.assertEqual(slide.shapes.title.text_frame.paragraphs[0].runs[0].font.name, "Georgia")

    def test_pdf_becomes_editable_powerpoint_and_uses_professional_profile(self):
        self._create_pdf()
        server.export_deck_to_markdown("deck.pdf")
        markdown_path = self.output_directory / "deck.md"
        markdown = markdown_path.read_text(encoding="utf-8")
        self.assertIn("PDF heading", markdown)
        markdown_path.write_text(
            markdown.replace("PDF heading", "Edited PDF heading"),
            encoding="utf-8",
        )

        server.polish_deck("deck.pdf", audience="professional")

        polished = Presentation(self.output_directory / "deck-polished.pptx")
        self.assertEqual(len(polished.slides), 1)
        self.assertEqual(polished.slides[0].shapes.title.text, "Edited PDF heading")
        self.assertEqual(
            polished.slides[0].shapes.title.text_frame.paragraphs[0].runs[0].font.name,
            "Aptos",
        )

    def _create_styled_pptx(self):
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[5])
        slide.shapes.title.text = "Highlighted title"
        title_run = slide.shapes.title.text_frame.paragraphs[0].runs[0]
        title_run.font.size = Pt(30)
        rPr = title_run._r.get_or_add_rPr()
        highlight = rPr.makeelement(qn("a:highlight"), {})
        highlight.append(highlight.makeelement(qn("a:srgbClr"), {"val": "C9DAF8"}))
        rPr.append(highlight)

        body = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(6), Inches(1))
        body.text_frame.text = "Small body text"
        body.text_frame.paragraphs[0].runs[0].font.size = Pt(20)
        source = slide.shapes.add_textbox(Inches(1), Inches(3.5), Inches(6), Inches(0.5))
        source.text_frame.text = "Source: USDA"
        source.text_frame.paragraphs[0].runs[0].font.size = Pt(12)
        large = slide.shapes.add_textbox(Inches(1), Inches(4.5), Inches(6), Inches(1))
        large.text_frame.text = "Already large"
        large.text_frame.paragraphs[0].runs[0].font.size = Pt(50)

        transparent = Image.new("RGBA", (40, 40), (0, 0, 0, 0))
        transparent.paste((0, 0, 0, 255), (0, 0, 10, 40))
        buffer = io.BytesIO()
        transparent.save(buffer, format="PNG")
        buffer.seek(0)
        slide.shapes.add_picture(buffer, Inches(7), Inches(1), Inches(2), Inches(2))
        presentation.save(self.data_directory / "styled.pptx")

    def test_civic_signal_dark_style(self):
        self._create_styled_pptx()
        result = server.polish_deck("styled.pptx", audience="civic_signal")

        polished = Presentation(self.output_directory / "styled-polished.pptx")
        slide = polished.slides[0]
        self.assertIn("civic_signal", result)
        self.assertEqual(slide.background.fill.fore_color.rgb, RGBColor(0x0B, 0x1A, 0x2E))

        title_run = slide.shapes.title.text_frame.paragraphs[0].runs[0]
        self.assertEqual(title_run.font.name, "Public Sans")
        # Highlight becomes an inverted signal-yellow block with navy text.
        highlight = title_run._r.find(qn("a:rPr")).find(qn("a:highlight"))
        self.assertEqual(highlight[0].get("val"), "FFD34E")
        self.assertEqual(title_run.font.color.rgb, RGBColor(0x0B, 0x1A, 0x2E))
        self.assertEqual(title_run.font.size, Pt(40))

        texts = {shape.text_frame.text: shape for shape in slide.shapes if shape.has_text_frame}
        body_run = texts["Small body text"].text_frame.paragraphs[0].runs[0]
        self.assertEqual(body_run.font.size, Pt(24))
        self.assertEqual(body_run.font.color.rgb, RGBColor(0xFF, 0xFF, 0xFF))
        self.assertEqual(texts["Source: USDA"].text_frame.paragraphs[0].runs[0].font.size, Pt(18))
        self.assertEqual(texts["Already large"].text_frame.paragraphs[0].runs[0].font.size, Pt(50))

        # A white plate sits directly behind the transparent picture.
        shapes = list(slide.shapes)
        picture_index = next(i for i, s in enumerate(shapes) if s.shape_type == MSO_SHAPE_TYPE.PICTURE)
        plate = shapes[picture_index - 1]
        self.assertEqual(plate.name, "Polish plate")
        self.assertEqual(plate.fill.fore_color.rgb, RGBColor(0xFF, 0xFF, 0xFF))
        self.assertIn("white plates behind 1", result)

    def test_unreadable_emphasis_color_is_fixed_on_dark_background(self):
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
        box.text_frame.text = "Dark red emphasis"
        run = box.text_frame.paragraphs[0].runs[0]
        run.font.size = Pt(28)
        run.font.color.rgb = RGBColor(0x8B, 0x00, 0x00)
        presentation.save(self.data_directory / "emphasis.pptx")

        server.polish_deck("emphasis.pptx", audience="civic_signal")

        polished = Presentation(self.output_directory / "emphasis-polished.pptx")
        color = polished.slides[0].shapes[0].text_frame.paragraphs[0].runs[0].font.color.rgb
        self.assertGreaterEqual(server._contrast(color, RGBColor(0x0B, 0x1A, 0x2E)), server.MIN_CONTRAST)

    def test_white_highlight_is_removed_not_turned_yellow(self):
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(1))
        box.text_frame.text = "UNC-Chapel Hill"
        rPr = box.text_frame.paragraphs[0].runs[0]._r.get_or_add_rPr()
        highlight = rPr.makeelement(qn("a:highlight"), {})
        highlight.append(highlight.makeelement(qn("a:schemeClr"), {"val": "lt1"}))
        rPr.append(highlight)
        presentation.save(self.data_directory / "white.pptx")

        server.polish_deck("white.pptx", audience="civic_signal")

        polished = Presentation(self.output_directory / "white-polished.pptx")
        run = polished.slides[0].shapes[0].text_frame.paragraphs[0].runs[0]
        self.assertIsNone(run._r.find(qn("a:rPr")).find(qn("a:highlight")))
        self.assertEqual(run.font.color.rgb, RGBColor(0xFF, 0xFF, 0xFF))

    def test_newsroom_light_is_the_default_style(self):
        self._create_styled_pptx()
        result = server.polish_deck("styled.pptx")

        polished = Presentation(self.output_directory / "styled-polished.pptx")
        slide = polished.slides[0]
        self.assertIn("newsroom_light", result)
        self.assertEqual(slide.background.fill.fore_color.rgb, RGBColor(0xF5, 0xF7, 0xFA))
        title_run = slide.shapes.title.text_frame.paragraphs[0].runs[0]
        # Highlight becomes a Carolina Navy block with light text.
        highlight = title_run._r.find(qn("a:rPr")).find(qn("a:highlight"))
        self.assertEqual(highlight[0].get("val"), "13294B")
        self.assertEqual(title_run.font.color.rgb, RGBColor(0xF5, 0xF7, 0xFA))
        texts = {shape.text_frame.text: shape for shape in slide.shapes if shape.has_text_frame}
        body_run = texts["Small body text"].text_frame.paragraphs[0].runs[0]
        self.assertEqual(body_run.font.color.rgb, RGBColor(0x14, 0x1A, 0x22))
        # Plates are only needed on dark backgrounds.
        self.assertFalse(any(shape.name == "Polish plate" for shape in slide.shapes))

    def test_text_only_slide_is_balanced_vertically(self):
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        box = slide.shapes.add_textbox(Inches(1), Inches(5.5), Inches(8), Inches(1))
        box.text_frame.text = "Study 1: Administrative Checkpoints"
        box.text_frame.paragraphs[0].runs[0].font.size = Pt(40)
        presentation.save(self.data_directory / "low.pptx")

        result = server.polish_deck("low.pptx")

        polished = Presentation(self.output_directory / "low-polished.pptx")
        shape = polished.slides[0].shapes[0]
        slide_height = polished.slide_height / 12700
        top = shape.top / 12700
        bottom = top + server._estimated_text_height_pt(shape)
        self.assertIn("Vertically balanced 1", result)
        self.assertLess(top, slide_height - bottom)  # optical lift: a bit above center
        self.assertLess(abs(top - (slide_height - bottom)), slide_height * 0.1)

    def test_design_preferences_are_served(self):
        (self.root / "design-preferences.md").write_text("# Design Preferences\nLarge type.", encoding="utf-8")
        with patch.object(server, "PREFERENCES_PATH", self.root / "design-preferences.md"):
            self.assertIn("Large type", server.get_design_preferences())


if __name__ == "__main__":
    unittest.main()
