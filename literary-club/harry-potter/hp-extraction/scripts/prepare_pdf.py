#!/usr/bin/env python3
"""
Подготовка страниц из исходника Гарри Поттера для визуального чтения.

Поддерживает два формата входного файла:
  1. ZIP-архив с расширением .pdf (внутри N.jpeg / N.txt)
  2. Настоящий PDF-документ

В обоих случаях на выходе — единообразная структура:
  output_dir/
    N.png       — изображение страницы (300 dpi)
    N.txt       — текстовый слой PDF для посимвольной сверки

Использование:
  python prepare_pdf.py <input.pdf> <first_page> <last_page> [--output-dir <dir>] [--dpi <dpi>] [--skip-glyph-audit]

Примеры:
  python prepare_pdf.py "Гарри Поттер книга 1 глава 3.pdf" 3 8
  python prepare_pdf.py "Исходник_ГП_1_1.pdf" 20 25 --output-dir ./pages --dpi 200
"""

import argparse
import io
import os
import shutil
import subprocess
import sys
import unicodedata
import zipfile


def ensure_pymupdf():
    """Проверяет наличие pymupdf, устанавливает при необходимости."""
    try:
        import pymupdf  # noqa: F401
        return True
    except ImportError:
        print("pymupdf не найден, устанавливаю...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "pymupdf", "-q"]
        )
        return True


def ensure_fonttools():
    """Проверяет наличие fonttools, устанавливает при необходимости."""
    try:
        from fontTools.ttLib import TTFont  # noqa: F401
        return True
    except ImportError:
        print("fonttools не найден, устанавливаю для проверки PDF-глифов...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "fonttools", "-q"]
        )
        return True


class DavidGlyphAuditor:
    """Сверяет PDF ToUnicode с реальными контурами встроенного David."""

    def __init__(self, doc):
        self.doc = doc
        self.reference_indexes = []
        self.embedded_maps = {}
        self.available = False
        self.reason = ""
        self._initialize()

    @staticmethod
    def _glyph_fingerprint(font, glyph_name):
        try:
            glyf = font["glyf"]
            coords, end_points, flags = glyf[glyph_name].getCoordinates(glyf)
            return (
                tuple(map(tuple, coords)),
                tuple(end_points),
                tuple(flags),
            )
        except Exception:
            return None

    @classmethod
    def _reference_index(cls, font_path):
        from fontTools.ttLib import TTFont

        font = TTFont(font_path)
        glyph_to_codepoints = {}
        for table in font["cmap"].tables:
            if not table.isUnicode():
                continue
            for codepoint, glyph_name in table.cmap.items():
                glyph_to_codepoints.setdefault(glyph_name, set()).add(codepoint)

        result = {}
        for glyph_name, codepoints in glyph_to_codepoints.items():
            fingerprint = cls._glyph_fingerprint(font, glyph_name)
            if fingerprint is not None:
                result.setdefault(fingerprint, set()).update(codepoints)
        return result

    @staticmethod
    def _reference_font_paths():
        candidates = []
        custom = os.environ.get("HP_DAVID_FONT")
        if custom:
            candidates.append(custom)

        windows_dir = os.environ.get("WINDIR")
        if windows_dir:
            fonts_dir = os.path.join(windows_dir, "Fonts")
            candidates.extend(
                [
                    os.path.join(fonts_dir, "david.ttf"),
                    os.path.join(fonts_dir, "davidbd.ttf"),
                ]
            )
        return [path for path in candidates if os.path.isfile(path)]

    def _initialize(self):
        paths = self._reference_font_paths()
        if not paths:
            self.reason = (
                "эталонный David не найден; задайте HP_DAVID_FONT для проверки"
            )
            return
        try:
            ensure_fonttools()
            self.reference_indexes = [self._reference_index(path) for path in paths]
            self.available = bool(self.reference_indexes)
        except Exception as exc:
            self.reason = f"не удалось подготовить fonttools: {exc}"

    @staticmethod
    def _same_unicode(left, right):
        return unicodedata.normalize("NFKD", chr(left)) == unicodedata.normalize(
            "NFKD", chr(right)
        )

    @staticmethod
    def _preferred_codepoint(codepoints):
        groups = (
            [cp for cp in codepoints if 0x05D0 <= cp <= 0x05EA],
            [cp for cp in codepoints if 0x0590 <= cp <= 0x05FF],
            [cp for cp in codepoints if 0x20 <= cp <= 0x7E],
        )
        selected = next((group for group in groups if group), list(codepoints))
        codepoint = sorted(selected)[0]
        if codepoint in {0, 13, 160, 0x034F, 0x2009, 0x200A}:
            return 32
        return codepoint

    def _embedded_map(self, xref):
        if xref in self.embedded_maps:
            return self.embedded_maps[xref]

        from fontTools.ttLib import TTFont

        _, _, _, font_data = self.doc.extract_font(xref)
        embedded = TTFont(io.BytesIO(font_data))
        best_map = {}
        for reference in self.reference_indexes:
            candidate = {}
            for glyph_id, glyph_name in enumerate(embedded.getGlyphOrder()):
                fingerprint = self._glyph_fingerprint(embedded, glyph_name)
                codepoints = reference.get(fingerprint)
                if codepoints:
                    candidate[glyph_id] = codepoints
            if len(candidate) > len(best_map):
                best_map = candidate

        self.embedded_maps[xref] = best_map
        return best_map

    def audit_page(self, page):
        if not self.available:
            return {"status": "skipped", "reason": self.reason}

        page_maps = []
        for font_info in page.get_fonts(full=True):
            xref, _, _, base_font, *_ = font_info
            if "david" not in base_font.casefold():
                continue
            glyph_map = self._embedded_map(xref)
            if glyph_map:
                page_maps.append((xref, glyph_map))

        if not page_maps:
            return {"status": "skipped", "reason": "встроенный David не найден"}

        total = 0
        mapped = 0
        unmapped = []
        mismatches = []
        david_spans = [
            span
            for span in page.get_texttrace()
            if "david" in span.get("font", "").casefold()
        ]
        grouped_spans = {}
        for span in david_spans:
            key = (span["font"], round(span["size"], 3))
            grouped_spans.setdefault(key, []).append(span)

        group_maps = {}
        for key, spans in grouped_spans.items():
            scored_maps = []
            for xref, glyph_map in page_maps:
                score = 0
                coverage = 0
                for span in spans:
                    for current, glyph_id, *_ in span["chars"]:
                        candidates = glyph_map.get(glyph_id)
                        if not candidates:
                            continue
                        coverage += 1
                        if any(
                            self._same_unicode(current, candidate)
                            for candidate in candidates
                        ):
                            score += 1
                scored_maps.append((score, coverage, xref, glyph_map))
            group_maps[key] = max(
                scored_maps, key=lambda item: (item[0], item[1])
            )

        for span in david_spans:
            key = (span["font"], round(span["size"], 3))
            _, _, xref, glyph_map = group_maps[key]
            for current, glyph_id, origin, _ in span["chars"]:
                total += 1
                candidates = glyph_map.get(glyph_id)
                if not candidates:
                    unmapped.append(
                        {
                            "current": chr(current),
                            "x": round(origin[0], 1),
                            "y": round(origin[1], 1),
                            "font_xref": xref,
                        }
                    )
                    continue
                mapped += 1
                if any(
                    self._same_unicode(current, candidate)
                    for candidate in candidates
                ):
                    continue
                expected = self._preferred_codepoint(candidates)
                mismatches.append(
                    {
                        "current": chr(current),
                        "expected": chr(expected),
                        "x": round(origin[0], 1),
                        "y": round(origin[1], 1),
                        "font_xref": xref,
                    }
                )

        if not total:
            return {"status": "skipped", "reason": "текст David отсутствует"}
        if mapped != total:
            return {
                "status": "inconclusive",
                "mapped": mapped,
                "total": total,
                "unmapped": unmapped,
                "mismatches": mismatches,
            }
        return {
            "status": "verified" if not mismatches else "mismatch",
            "mapped": mapped,
            "total": total,
            "mismatches": mismatches,
        }


def print_glyph_audit(page_num, result):
    """Печатает компактный результат проверки глифов страницы."""
    status = result["status"]
    if status == "verified":
        print(
            f"  стр. {page_num}: ивритский слой David подтверждён по глифам "
            f"({result['mapped']}/{result['total']})"
        )
    elif status == "mismatch":
        print(
            f"  стр. {page_num}: ВНИМАНИЕ — найдено расхождений "
            f"ToUnicode/глифы: {len(result['mismatches'])}"
        )
        for item in result["mismatches"][:10]:
            print(
                "    "
                f"({item['x']}, {item['y']}): "
                f"{item['current']!r} → {item['expected']!r}"
            )
    elif status == "inconclusive":
        print(
            f"  стр. {page_num}: проверка глифов неполна "
            f"({result['mapped']}/{result['total']})"
        )
        for item in result.get("unmapped", [])[:10]:
            print(
                "    "
                f"не сопоставлен символ {item['current']!r} "
                f"в ({item['x']}, {item['y']})"
            )
    else:
        print(f"  стр. {page_num}: проверка глифов пропущена — {result['reason']}")


def detect_format(filepath):
    """Определяет формат файла: 'zip' или 'pdf'."""
    if zipfile.is_zipfile(filepath):
        return "zip"
    # Проверяем PDF-сигнатуру
    with open(filepath, "rb") as f:
        header = f.read(5)
    if header == b"%PDF-":
        return "pdf"
    raise ValueError(f"Неизвестный формат файла: {filepath}")


def process_zip(filepath, first_page, last_page, output_dir):
    """Извлекает страницы из ZIP-архива (формат с JPEG/TXT внутри)."""
    with zipfile.ZipFile(filepath, "r") as zf:
        names = zf.namelist()
        for page_num in range(first_page, last_page + 1):
            # Ищем изображение
            img_found = False
            for ext in ("jpeg", "jpg", "png"):
                img_name = f"{page_num}.{ext}"
                if img_name in names:
                    zf.extract(img_name, output_dir)
                    src = os.path.join(output_dir, img_name)
                    dst = os.path.join(output_dir, f"{page_num}.png")
                    if ext != "png":
                        # Переименовываем для единообразия
                        shutil.move(src, dst)
                    img_found = True
                    print(f"  стр. {page_num}: изображение извлечено ({img_name})")
                    break
            if not img_found:
                print(f"  стр. {page_num}: ВНИМАНИЕ — изображение не найдено в архиве")

            # Ищем текст
            txt_name = f"{page_num}.txt"
            if txt_name in names:
                zf.extract(txt_name, output_dir)
                print(f"  стр. {page_num}: текст извлечён ({txt_name})")
            else:
                print(f"  стр. {page_num}: текстовый файл отсутствует")


def process_pdf(filepath, first_page, last_page, output_dir, dpi, glyph_audit=True):
    """Конвертирует страницы из PDF в изображения и извлекает текст."""
    ensure_pymupdf()
    import pymupdf

    doc = pymupdf.open(filepath)
    total_pages = len(doc)
    auditor = DavidGlyphAuditor(doc) if glyph_audit else None

    for page_num in range(first_page, last_page + 1):
        page_idx = page_num - 1  # pymupdf использует 0-based индексы

        if page_idx < 0 or page_idx >= total_pages:
            print(f"  стр. {page_num}: ВНИМАНИЕ — страница за пределами документа ({total_pages} стр.)")
            continue

        page = doc[page_idx]

        # Изображение
        pix = page.get_pixmap(dpi=dpi)
        img_path = os.path.join(output_dir, f"{page_num}.png")
        pix.save(img_path)
        print(f"  стр. {page_num}: изображение сохранено ({dpi} dpi)")

        # Вспомогательный текст
        text = page.get_text()
        if text.strip():
            txt_path = os.path.join(output_dir, f"{page_num}.txt")
            with open(txt_path, "w", encoding="utf-8") as f:
                f.write(text)
            print(f"  стр. {page_num}: вспомогательный текст сохранён")
            if auditor:
                print_glyph_audit(page_num, auditor.audit_page(page))

    doc.close()


def main():
    parser = argparse.ArgumentParser(
        description="Подготовка страниц из исходника Гарри Поттера"
    )
    parser.add_argument("input", help="Путь к PDF/ZIP файлу")
    parser.add_argument("first_page", type=int, help="Первая страница")
    parser.add_argument("last_page", type=int, help="Последняя страница")
    parser.add_argument(
        "--output-dir", default=None,
        help="Папка для результатов (по умолчанию: ./tmp рядом с входным файлом)"
    )
    parser.add_argument(
        "--dpi", type=int, default=300,
        help="Разрешение для конвертации PDF (по умолчанию: 300)"
    )
    parser.add_argument(
        "--skip-glyph-audit", action="store_true",
        help="Не сверять ToUnicode с контурами встроенного шрифта David"
    )
    args = parser.parse_args()

    if not os.path.isfile(args.input):
        print(f"Ошибка: файл не найден: {args.input}")
        sys.exit(1)

    if args.first_page > args.last_page:
        print("Ошибка: первая страница больше последней")
        sys.exit(1)

    # Определяем выходную папку
    if args.output_dir:
        output_dir = args.output_dir
    else:
        output_dir = os.path.join(os.path.dirname(args.input) or ".", "tmp")

    os.makedirs(output_dir, exist_ok=True)

    # Определяем формат
    fmt = detect_format(args.input)
    print(f"Файл: {os.path.basename(args.input)}")
    print(f"Формат: {fmt.upper()}")
    print(f"Страницы: {args.first_page}–{args.last_page}")
    print(f"Выходная папка: {output_dir}")
    print()

    if fmt == "zip":
        process_zip(args.input, args.first_page, args.last_page, output_dir)
    else:
        process_pdf(
            args.input,
            args.first_page,
            args.last_page,
            output_dir,
            args.dpi,
            glyph_audit=not args.skip_glyph_audit,
        )

    print()
    print("Готово! Файлы:")
    for f in sorted(os.listdir(output_dir)):
        fpath = os.path.join(output_dir, f)
        size_kb = os.path.getsize(fpath) / 1024
        print(f"  {f} ({size_kb:.0f} KB)")


if __name__ == "__main__":
    main()
