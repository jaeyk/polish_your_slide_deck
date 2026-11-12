import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pymupdf as fitz
from pptx import Presentation

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
        server.polish_deck("deck.pptx")

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


if __name__ == "__main__":
    unittest.main()
