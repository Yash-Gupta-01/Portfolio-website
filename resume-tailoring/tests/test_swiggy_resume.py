import importlib.util
import tempfile
import unittest
from pathlib import Path

from docx import Document


SCRIPT = Path(__file__).parents[1] / "scripts" / "build_resume.py"


class SwiggyResumeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        spec = importlib.util.spec_from_file_location("build_resume", SCRIPT)
        cls.builder = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(cls.builder)

    def test_swiggy_resume_is_metric_first_and_honest(self):
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "swiggy.docx"
            self.builder.create_resume_document(output, company="Swiggy", role="SDE1")
            document = Document(output)
            paragraphs = [p.text.strip() for p in document.paragraphs if p.text.strip()]
            text = "\n".join(paragraphs)
            bullets = [p.lstrip("• ") for p in paragraphs if p.startswith("•")]
            first_bullet = next(p for p in document.paragraphs if p.text.strip().startswith("•"))

        self.assertIn("Open to relocate to Bangalore", text)
        self.assertTrue(bullets[0].startswith("14,000-row"), bullets[0])
        self.assertIn("14,000-row spike readiness", text)
        self.assertTrue(any(b.startswith("3×") for b in bullets))
        self.assertTrue(any(b.startswith("1,200 TPS") for b in bullets))
        self.assertIn("Go (Golang)", text)
        self.assertIn("Agentic AI", text)
        self.assertIn("100% payment reconciliation", text)
        self.assertIn("2 additional hackathon podiums", text)
        self.assertNotIn("zero data loss", text.lower())
        self.assertNotIn("1 year of experience", text.lower())
        self.assertGreaterEqual(first_bullet.runs[0].font.size.pt, 9.5)
        self.assertGreaterEqual(first_bullet.paragraph_format.line_spacing, 1.08)


if __name__ == "__main__":
    unittest.main()
