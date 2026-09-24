#!/usr/bin/env python3
from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional, Sequence
from zipfile import BadZipFile, ZipFile

from docx import Document
from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
W = f"{{{W_NS}}}"

HEBREW_RE = re.compile(r"[\u0590-\u05FF]")
LTR_RE = re.compile(r"[A-Za-z\u0400-\u052F]")
PAGE_RE = re.compile(r"^Страница\s+(\d+)\s*$")
MD_PAGE_RE = re.compile(r"^#\s+Страница\s+(\d+)\s*$", re.MULTILINE)
MD_TABLE_DELIM_RE = re.compile(
    r"^\|\s*:?[-]+:?(?:\s*\|\s*:?[-]+:?)+\s*\|\s*$", re.MULTILINE
)
SECTION_HEADINGS = {
    "Иврит",
    "Подстрочный перевод",
    "Литературный перевод",
    "Список сложных слов",
    "Различия ивритского и русского переводов",
}


@dataclass(frozen=True)
class ValidationReport:
    output_docx: Path
    pages: tuple[int, ...]
    drawings: int
    tables: int
    hebrew_runs: int


def _text(element) -> str:
    return "".join(element.xpath(".//w:t/text()", namespaces=NS))


def _attr(element, name: str) -> Optional[str]:
    return None if element is None else element.get(f"{W}{name}")


def _enabled(element) -> bool:
    return element is not None and _attr(element, "val") not in {"0", "false", "off"}


def expected_from_markdown(markdown_path: Path) -> tuple[list[int], int]:
    source = markdown_path.read_text(encoding="utf-8")
    return (
        [int(value) for value in MD_PAGE_RE.findall(source)],
        len(MD_TABLE_DELIM_RE.findall(source)),
    )


def validate_docx(
    docx_path: Path,
    *,
    expected_pages: Optional[Sequence[int]] = None,
    expected_table_count: Optional[int] = None,
) -> ValidationReport:
    docx_path = Path(docx_path)
    errors: list[str] = []

    try:
        document = Document(str(docx_path))
        with ZipFile(docx_path) as archive:
            root = etree.fromstring(archive.read("word/document.xml"))
    except (BadZipFile, KeyError, OSError, ValueError, etree.XMLSyntaxError) as exc:
        raise ValueError(f"ERROR: DOCX validation failed: cannot open document: {exc}") from exc

    body = root.find(f"{W}body")
    if body is None:
        raise ValueError("ERROR: DOCX validation failed: document body is missing")

    body_children = list(body)
    headings: list[tuple[int, object]] = []
    for element in body_children:
        if element.tag != f"{W}p":
            continue
        match = PAGE_RE.fullmatch(_text(element).strip())
        if match:
            headings.append((int(match.group(1)), element))

    actual_pages = [number for number, _ in headings]
    if expected_pages is not None and actual_pages != list(expected_pages):
        errors.append(
            f"page headings/order mismatch: expected {list(expected_pages)}, got {actual_pages}"
        )
    elif not actual_pages:
        errors.append("no page headings found")

    for index, (page_number, heading) in enumerate(headings):
        if index > 0 and not heading.xpath("./w:pPr/w:pageBreakBefore", namespaces=NS):
            errors.append(f"page {page_number}: missing pageBreakBefore")
        position = body_children.index(heading)
        following = body_children[position + 1] if position + 1 < len(body_children) else None
        if following is None or following.tag != f"{W}p" or not following.xpath(
            ".//w:drawing", namespaces=NS
        ):
            errors.append(f"page {page_number}: illustration is not immediately after heading")

    drawings = root.xpath(".//w:drawing", namespaces=NS)
    expected_drawings = len(expected_pages) if expected_pages is not None else len(actual_pages)
    if len(drawings) != expected_drawings:
        errors.append(
            f"illustration count mismatch: expected {expected_drawings}, got {len(drawings)}"
        )

    paragraph_texts = [_text(p).strip() for p in root.xpath(".//w:p", namespaces=NS)]
    leaked = sorted(SECTION_HEADINGS.intersection(paragraph_texts))
    if leaked:
        errors.append(f"section headings were not removed: {', '.join(leaked)}")
    combined_text = "\n".join(paragraph_texts)
    if "**" in combined_text or "`" in combined_text:
        errors.append("visible inline Markdown markers remain")

    section = document.sections[0]
    text_width_twips = int(
        round((section.page_width - section.left_margin - section.right_margin) / 635)
    )
    tables = root.xpath(".//w:tbl", namespaces=NS)
    if expected_table_count is not None and len(tables) != expected_table_count:
        errors.append(
            f"table count mismatch: expected {expected_table_count}, got {len(tables)}"
        )

    for table_index, table in enumerate(tables, start=1):
        if not table.xpath("./w:tblPr/w:tblLayout[@w:type='fixed']", namespaces=NS):
            errors.append(f"table {table_index}: layout is not fixed")
        if not table.xpath("./w:tblPr/w:tblW[@w:type='pct' and @w:w='5000']", namespaces=NS):
            errors.append(f"table {table_index}: width is not explicitly 100%")

        grid_widths: list[int] = []
        for grid_col in table.xpath("./w:tblGrid/w:gridCol", namespaces=NS):
            try:
                grid_widths.append(int(_attr(grid_col, "w") or "0"))
            except ValueError:
                grid_widths.append(0)
        if (
            not grid_widths
            or any(width <= 0 for width in grid_widths)
            or sum(grid_widths) != text_width_twips
            or max(grid_widths) - min(grid_widths) > 1
        ):
            errors.append(f"table {table_index}: column grid is not an even full-width grid")

        for row_index, row in enumerate(table.xpath("./w:tr", namespaces=NS), start=1):
            if not row.xpath("./w:trPr/w:cantSplit", namespaces=NS):
                errors.append(f"table {table_index}, row {row_index}: cantSplit is missing")

    hebrew_runs = 0
    for run_index, run in enumerate(root.xpath(".//w:r", namespaces=NS), start=1):
        if not HEBREW_RE.search(_text(run)):
            continue
        hebrew_runs += 1
        r_pr = run.find(f"{W}rPr")
        fonts = None if r_pr is None else r_pr.find(f"{W}rFonts")
        if r_pr is None or any(_attr(fonts, name) != "David" for name in ("ascii", "hAnsi", "cs")):
            errors.append(f"Hebrew run {run_index}: David font is missing")
            continue
        if not _enabled(r_pr.find(f"{W}rtl")) or not _enabled(r_pr.find(f"{W}cs")):
            errors.append(f"Hebrew run {run_index}: rtl/cs is missing")
        if _attr(r_pr.find(f"{W}sz"), "val") != "36" or _attr(
            r_pr.find(f"{W}szCs"), "val"
        ) != "36":
            errors.append(f"Hebrew run {run_index}: size is not 18 pt")
        if _enabled(r_pr.find(f"{W}b")) and not _enabled(r_pr.find(f"{W}bCs")):
            errors.append(f"Hebrew run {run_index}: bold complex-script flag is missing")

    for paragraph_index, paragraph in enumerate(root.xpath(".//w:p", namespaces=NS), start=1):
        value = _text(paragraph)
        if len(HEBREW_RE.findall(value)) <= len(LTR_RE.findall(value)):
            continue
        p_pr = paragraph.find(f"{W}pPr")
        if p_pr is None or not _enabled(p_pr.find(f"{W}bidi")):
            errors.append(f"Hebrew paragraph {paragraph_index}: bidi is missing")
        if p_pr is None or _attr(p_pr.find(f"{W}jc"), "val") not in {"right", "start"}:
            errors.append(f"Hebrew paragraph {paragraph_index}: alignment is not right")

    if errors:
        raise ValueError("ERROR: DOCX validation failed:\n- " + "\n- ".join(errors))

    return ValidationReport(
        output_docx=docx_path,
        pages=tuple(actual_pages),
        drawings=len(drawings),
        tables=len(tables),
        hebrew_runs=hebrew_runs,
    )


def parse_args(argv: Optional[Iterable[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Validate generated Hebrew Harry Potter DOCX structure without rendering."
    )
    parser.add_argument("docx", type=Path)
    parser.add_argument("--markdown", type=Path, help="Source HP_ch*_translate.md")
    return parser.parse_args(argv)


def main(argv: Optional[Iterable[str]] = None) -> int:
    args = parse_args(argv)
    pages = None
    table_count = None
    if args.markdown is not None:
        pages, table_count = expected_from_markdown(args.markdown)
    try:
        report = validate_docx(
            args.docx,
            expected_pages=pages,
            expected_table_count=table_count,
        )
    except (OSError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 1
    print(
        f"VALID: {report.output_docx} pages={len(report.pages)} "
        f"drawings={report.drawings} tables={report.tables} "
        f"hebrew_runs={report.hebrew_runs}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
