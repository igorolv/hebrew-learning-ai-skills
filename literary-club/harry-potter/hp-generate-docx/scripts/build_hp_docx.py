#!/usr/bin/env python3
from __future__ import annotations

import argparse
import math
import os
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, List, Optional

from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.table import WD_ALIGN_VERTICAL, WD_ROW_HEIGHT_RULE, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


HEBREW_RE = re.compile(r"[\u0590-\u05FF]")
PAGE_RE = re.compile(r"^#\s+Страница\s+(\d+)\s*$")
SUBHEADING_RE = re.compile(r"^##\s+")
MD_TABLE_ROW_RE = re.compile(r"^\|.*\|\s*$")
DELIM_RE = re.compile(r"^\|\s*:?[-]+:?(?:\s*\|\s*:?[-]+:?)+\s*\|\s*$")
MD_NAME_RE = re.compile(r"HP_ch(\d+)_(\d+)_(\d+)_translate\.md$", re.IGNORECASE)
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".webp"}


@dataclass
class Block:
    kind: str  # paragraph | table
    content: object


@dataclass
class PageContent:
    number: int
    blocks: List[Block] = field(default_factory=list)


@dataclass
class BuildResult:
    output_docx: Path
    render_dir: Optional[Path] = None


def contains_hebrew(text: str) -> bool:
    return bool(HEBREW_RE.search(text or ""))


def extract_last_number(text: str) -> Optional[int]:
    nums = re.findall(r"(\d+)", text)
    return int(nums[-1]) if nums else None


def normalize_table_row(line: str) -> List[str]:
    stripped = line.strip().strip("|")
    return [cell.strip() for cell in stripped.split("|")]


def parse_markdown(markdown_text: str) -> List[PageContent]:
    lines = markdown_text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    pages: List[PageContent] = []
    current: Optional[PageContent] = None
    paragraph_buffer: List[str] = []
    i = 0

    def flush_paragraph() -> None:
        nonlocal paragraph_buffer, current
        if current is None:
            paragraph_buffer = []
            return
        text = "\n".join(paragraph_buffer).strip()
        paragraph_buffer = []
        if text:
            current.blocks.append(Block("paragraph", text))

    while i < len(lines):
        line = lines[i]
        page_match = PAGE_RE.match(line.strip())
        if page_match:
            flush_paragraph()
            current = PageContent(number=int(page_match.group(1)))
            pages.append(current)
            i += 1
            continue

        if current is None:
            i += 1
            continue

        stripped = line.strip()
        if SUBHEADING_RE.match(stripped):
            flush_paragraph()
            i += 1
            continue

        if MD_TABLE_ROW_RE.match(stripped):
            flush_paragraph()
            table_lines = [stripped]
            i += 1
            while i < len(lines) and MD_TABLE_ROW_RE.match(lines[i].strip()):
                table_lines.append(lines[i].strip())
                i += 1
            rows = [normalize_table_row(x) for x in table_lines]
            if len(rows) >= 2 and DELIM_RE.match(table_lines[1]):
                rows = rows[2:]  # drop markdown header + delimiter
            if rows:
                current.blocks.append(Block("table", rows))
            continue

        if stripped == "":
            flush_paragraph()
            i += 1
            continue

        paragraph_buffer.append(line)
        i += 1

    flush_paragraph()
    return pages


def load_images_from_zip(zip_path: Path) -> dict[int, Path]:
    temp_dir = Path(tempfile.mkdtemp(prefix="hp_docx_images_"))
    mapping: dict[int, Path] = {}
    with zipfile.ZipFile(zip_path, "r") as zf:
        for info in zf.infolist():
            if info.is_dir():
                continue
            name = Path(info.filename).name
            suffix = Path(name).suffix.lower()
            if suffix not in IMAGE_SUFFIXES:
                continue
            page_number = extract_last_number(name)
            if page_number is None:
                continue
            target = temp_dir / name
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info) as src, open(target, "wb") as dst:
                shutil.copyfileobj(src, dst)
            mapping[page_number] = target
    return mapping


def set_repeat_table_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_row_no_break(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    cant_split = OxmlElement("w:cantSplit")
    tr_pr.append(cant_split)


def set_cell_width(cell, width_twips: int) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_w = tc_pr.find(qn("w:tcW"))
    if tc_w is None:
        tc_w = OxmlElement("w:tcW")
        tc_pr.append(tc_w)
    tc_w.set(qn("w:w"), str(width_twips))
    tc_w.set(qn("w:type"), "dxa")


def set_paragraph_bidi(paragraph, enabled: bool) -> None:
    p_pr = paragraph._p.get_or_add_pPr()
    bidi = p_pr.find(qn("w:bidi"))
    if bidi is None:
        bidi = OxmlElement("w:bidi")
        p_pr.append(bidi)
    bidi.set(qn("w:val"), "1" if enabled else "0")


def set_paragraph_jc(paragraph, value: str) -> None:
    """Set w:jc directly using OOXML values ('start', 'end', 'center', 'both')."""
    p_pr = paragraph._p.get_or_add_pPr()
    jc = p_pr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        p_pr.append(jc)
    jc.set(qn("w:val"), value)


def ensure_rtl_run(run, rtl: bool) -> None:
    r_pr = run._r.get_or_add_rPr()
    rtl_el = r_pr.find(qn("w:rtl"))
    if rtl_el is None:
        rtl_el = OxmlElement("w:rtl")
        r_pr.append(rtl_el)
    rtl_el.set(qn("w:val"), "1" if rtl else "0")


def set_run_font(run, *, font_name: str, font_size_pt: int, rtl: bool) -> None:
    run.font.name = font_name
    run.font.size = Pt(font_size_pt)
    r_pr = run._r.get_or_add_rPr()

    r_fonts = r_pr.find(qn("w:rFonts"))
    if r_fonts is None:
        r_fonts = OxmlElement("w:rFonts")
        r_pr.append(r_fonts)
    for attr in ("w:ascii", "w:hAnsi", "w:cs"):
        r_fonts.set(qn(attr), font_name)

    sz = r_pr.find(qn("w:sz"))
    if sz is None:
        sz = OxmlElement("w:sz")
        r_pr.append(sz)
    sz.set(qn("w:val"), str(font_size_pt * 2))

    sz_cs = r_pr.find(qn("w:szCs"))
    if sz_cs is None:
        sz_cs = OxmlElement("w:szCs")
        r_pr.append(sz_cs)
    sz_cs.set(qn("w:val"), str(font_size_pt * 2))

    ensure_rtl_run(run, rtl)


def style_paragraph_text(paragraph, text: str, force_hebrew: Optional[bool] = None) -> None:
    has_hebrew = contains_hebrew(text) if force_hebrew is None else force_hebrew
    paragraph.text = ""
    run = paragraph.add_run(text)
    if has_hebrew:
        set_paragraph_bidi(paragraph, True)
        set_paragraph_jc(paragraph, "start")
        set_run_font(run, font_name="David", font_size_pt=18, rtl=True)
    else:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
        set_paragraph_bidi(paragraph, False)
        set_run_font(run, font_name="Times New Roman", font_size_pt=12, rtl=False)

    paragraph.paragraph_format.space_after = Pt(6)
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.15


def style_page_heading(paragraph, text: str) -> None:
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(text)
    run.bold = True
    set_run_font(run, font_name="Times New Roman", font_size_pt=14, rtl=False)


def set_document_defaults(document: Document) -> None:
    style = document.styles["Normal"]
    style.font.name = "Times New Roman"
    style.font.size = Pt(12)
    section = document.sections[0]
    section.page_width = Cm(21)
    section.page_height = Cm(29.7)
    section.top_margin = Cm(2.54)
    section.bottom_margin = Cm(2.54)
    section.left_margin = Cm(2.54)
    section.right_margin = Cm(2.54)


def get_text_width_emu(document: Document) -> int:
    section = document.sections[-1]
    return section.page_width - section.left_margin - section.right_margin


def get_text_width_twips(document: Document) -> int:
    return int(round(get_text_width_emu(document) / 635))


def add_picture_for_page(document: Document, image_path: Path) -> None:
    p = document.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(12)
    width = get_text_width_emu(document)
    run = p.add_run()
    run.add_picture(str(image_path), width=width)


def build_table(document: Document, rows: List[List[str]]) -> None:
    if not rows:
        return
    col_count = max(len(r) for r in rows)
    table = document.add_table(rows=len(rows), cols=col_count)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    tbl_pr = table._tbl.tblPr
    tbl_layout = tbl_pr.find(qn("w:tblLayout"))
    if tbl_layout is None:
        tbl_layout = OxmlElement("w:tblLayout")
        tbl_pr.append(tbl_layout)
    tbl_layout.set(qn("w:type"), "fixed")

    table_width_twips = get_text_width_twips(document)
    col_width = max(1, math.floor(table_width_twips / col_count))

    for row_idx, row_data in enumerate(rows):
        row = table.rows[row_idx]
        row.height_rule = WD_ROW_HEIGHT_RULE.AUTO
        set_row_no_break(row)
        for col_idx in range(col_count):
            cell = row.cells[col_idx]
            set_cell_width(cell, col_width)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            tc_pr = cell._tc.get_or_add_tcPr()
            no_wrap = tc_pr.find(qn("w:noWrap"))
            if no_wrap is not None:
                tc_pr.remove(no_wrap)

            for p in cell.paragraphs:
                p.clear()
            text = row_data[col_idx] if col_idx < len(row_data) else ""
            paragraph = cell.paragraphs[0]
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.paragraph_format.left_indent = Pt(0)
            paragraph.paragraph_format.right_indent = Pt(0)
            style_paragraph_text(paragraph, text)

            tc_mar = tc_pr.find(qn("w:tcMar"))
            if tc_mar is None:
                tc_mar = OxmlElement("w:tcMar")
                tc_pr.append(tc_mar)
            for side in ("top", "left", "bottom", "right"):
                el = tc_mar.find(qn(f"w:{side}"))
                if el is None:
                    el = OxmlElement(f"w:{side}")
                    tc_mar.append(el)
                el.set(qn("w:w"), "100")
                el.set(qn("w:type"), "dxa")

    document.add_paragraph().paragraph_format.space_after = Pt(6)


def derive_output_name(md_path: Path) -> str:
    match = MD_NAME_RE.search(md_path.name)
    if match:
        chapter, page_from, page_to = match.groups()
        return f"Гарри Поттер глава {chapter} страницы {page_from}-{page_to}.docx"
    return f"{md_path.stem}.docx"


def render_docx(docx_path: Path, out_dir: Path) -> None:
    out_dir.mkdir(parents=True, exist_ok=True)
    env = os.environ.copy()
    profile_dir = out_dir / "lo_profile"
    profile_dir.mkdir(parents=True, exist_ok=True)
    env["HOME"] = str(profile_dir)
    subprocess.run(
        [
            "libreoffice",
            "--headless",
            f"-env:UserInstallation=file://{profile_dir}",
            "--convert-to",
            "pdf",
            "--outdir",
            str(out_dir),
            str(docx_path),
        ],
        check=True,
        env=env,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )


def build_docx(md_path: Path, images_zip_path: Path, output_path: Optional[Path], render: bool) -> BuildResult:
    pages = parse_markdown(md_path.read_text(encoding="utf-8"))
    image_map = load_images_from_zip(images_zip_path)

    document = Document()
    set_document_defaults(document)

    for idx, page in enumerate(pages):
        if idx > 0:
            document.add_section(WD_SECTION_START.NEW_PAGE)
            set_document_defaults(document)

        heading = document.add_paragraph()
        style_page_heading(heading, f"Страница {page.number}")

        image_path = image_map.get(page.number)
        if image_path and image_path.exists():
            add_picture_for_page(document, image_path)

        for block in page.blocks:
            if block.kind == "paragraph":
                for chunk in str(block.content).split("\n"):
                    para = document.add_paragraph()
                    style_paragraph_text(para, chunk)
            elif block.kind == "table":
                build_table(document, block.content)  # type: ignore[arg-type]

    out_path = output_path or (md_path.parent / derive_output_name(md_path))
    document.save(str(out_path))

    render_dir = None
    if render:
        render_dir = out_path.with_suffix("")
        render_docx(out_path, render_dir)

    return BuildResult(output_docx=out_path, render_dir=render_dir)


def parse_args(argv: Optional[Iterable[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build Harry Potter Hebrew markdown into DOCX.")
    parser.add_argument("markdown", type=Path, help="Path to HP_ch*_translate.md")
    parser.add_argument("images_zip", type=Path, help="Path to ZIP with illustrations")
    parser.add_argument("-o", "--output", type=Path, help="Output DOCX path")
    parser.add_argument("--no-render", action="store_true", help="Skip LibreOffice PDF render")
    return parser.parse_args(argv)


def main(argv: Optional[Iterable[str]] = None) -> int:
    args = parse_args(argv)
    result = build_docx(
        md_path=args.markdown,
        images_zip_path=args.images_zip,
        output_path=args.output,
        render=not args.no_render,
    )
    print(f"DOCX: {result.output_docx}")
    if result.render_dir:
        print(f"RENDER_DIR: {result.render_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
