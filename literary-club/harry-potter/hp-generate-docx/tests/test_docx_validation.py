import base64
import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


ROOT = Path(__file__).parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


VALIDATOR = load_module("validate_hp_docx_tests", SCRIPTS / "validate_hp_docx.py")
BUILDER = load_module("build_hp_docx_validation_tests", SCRIPTS / "build_hp_docx.py")

PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
)
MARKDOWN = """# Страница 1

## Иврит

זה **טוב**.

## Подстрочный перевод

| Иврит | Перевод |
|---|---|
| **טוב** | **хорошо** |
"""


class DocxValidationTests(unittest.TestCase):
    def build_fixture(self, root: Path) -> tuple[Path, Path]:
        markdown = root / "HP_ch1_1_1_translate.md"
        image = root / "HP_ch1_page_1.png"
        output = root / "result.docx"
        markdown.write_text(MARKDOWN, encoding="utf-8")
        image.write_bytes(PNG)
        BUILDER.build_docx(markdown, [image], output)
        return markdown, output

    def test_generated_docx_passes_structural_validation(self):
        with tempfile.TemporaryDirectory() as tmp:
            markdown, output = self.build_fixture(Path(tmp))
            pages, table_count = VALIDATOR.expected_from_markdown(markdown)
            report = VALIDATOR.validate_docx(
                output,
                expected_pages=pages,
                expected_table_count=table_count,
            )
            self.assertEqual(report.pages, (1,))
            self.assertEqual(report.drawings, 1)
            self.assertEqual(report.tables, 1)

    def test_validator_rejects_non_explicit_table_width(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            markdown, output = self.build_fixture(root)
            broken = root / "broken.docx"

            with ZipFile(output) as source, ZipFile(broken, "w", ZIP_DEFLATED) as target:
                for info in source.infolist():
                    payload = source.read(info.filename)
                    if info.filename == "word/document.xml":
                        xml = etree.fromstring(payload)
                        ns = {"w": VALIDATOR.W_NS}
                        width = xml.xpath(".//w:tbl[1]/w:tblPr/w:tblW", namespaces=ns)[0]
                        width.set(f"{{{VALIDATOR.W_NS}}}type", "auto")
                        width.set(f"{{{VALIDATOR.W_NS}}}w", "0")
                        payload = etree.tostring(
                            xml, xml_declaration=True, encoding="UTF-8", standalone=True
                        )
                    target.writestr(info, payload)

            pages, table_count = VALIDATOR.expected_from_markdown(markdown)
            with self.assertRaisesRegex(ValueError, "width is not explicitly 100%"):
                VALIDATOR.validate_docx(
                    broken,
                    expected_pages=pages,
                    expected_table_count=table_count,
                )


if __name__ == "__main__":
    unittest.main()
