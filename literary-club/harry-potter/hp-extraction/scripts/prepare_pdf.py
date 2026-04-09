#!/usr/bin/env python3
"""
Подготовка страниц из исходника Гарри Поттера для визуального чтения.

Поддерживает два формата входного файла:
  1. ZIP-архив с расширением .pdf (внутри N.jpeg / N.txt)
  2. Настоящий PDF-документ

В обоих случаях на выходе — единообразная структура:
  output_dir/
    N.png       — изображение страницы (300 dpi)
    N.txt       — вспомогательный текст (ненадёжен для иврита!)

Использование:
  python prepare_pdf.py <input.pdf> <first_page> <last_page> [--output-dir <dir>] [--dpi <dpi>]

Примеры:
  python prepare_pdf.py "Гарри Поттер книга 1 глава 3.pdf" 3 8
  python prepare_pdf.py "Исходник_ГП_1_1.pdf" 20 25 --output-dir ./pages --dpi 200
"""

import argparse
import os
import shutil
import subprocess
import sys
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


def process_pdf(filepath, first_page, last_page, output_dir, dpi):
    """Конвертирует страницы из PDF в изображения и извлекает текст."""
    ensure_pymupdf()
    import pymupdf

    doc = pymupdf.open(filepath)
    total_pages = len(doc)

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
        process_pdf(args.input, args.first_page, args.last_page, output_dir, args.dpi)

    print()
    print("Готово! Файлы:")
    for f in sorted(os.listdir(output_dir)):
        fpath = os.path.join(output_dir, f)
        size_kb = os.path.getsize(fpath) / 1024
        print(f"  {f} ({size_kb:.0f} KB)")


if __name__ == "__main__":
    main()
