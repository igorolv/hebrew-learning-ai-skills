#!/usr/bin/env python3
"""
Extract structured content from Hebrew lesson PPTX files.

Usage:
    python extract_pptx.py <input.pptx> <output_dir>

Outputs:
    <output_dir>/extracted.md   — structured markdown with all content
    <output_dir>/images/        — extracted images from slides
"""

import sys
import os
import json
import re
import base64
from pathlib import Path

try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
except ImportError:
    print("ERROR: python-pptx not installed. Run: pip install python-pptx --break-system-packages")
    sys.exit(1)


def classify_slide(texts, has_table, has_image, table_data):
    """
    Classify a slide into one of these categories:
    - grammar: Russian explanatory text with Hebrew examples
    - vocabulary: table with מילים חדשות header
    - conjugation: conjugation table (שורש, עבר, הווה, etc.)
    - verb_summary: summary table of verbs by gzarot (שלמים, פ״נ, etc.)
    - shem_peula: verbal noun table (שם פעולה)
    - exercise_fill: fill-in-the-blank exercises (______)
    - exercise_translate: translation exercises (numbered Russian sentences)
    - exercise_table: table-based exercises (fill table cells)
    - exercise_choose: choose the correct word
    - exercise_oged: אוגד exercises
    - reading: reading text (long Hebrew text, dialog, story)
    - image_task: slide with primarily an image (spot-the-difference, etc.)
    - joke: jokes / humorous texts (חידות, גולם, חושם, חלם)
    - other: unclassified
    """
    all_text = "\n".join(texts).strip()
    all_text_lower = all_text.lower()

    # Check for image-only slides (very little text, has image)
    if has_image and len(all_text.replace(" ", "").replace("\n", "")) < 30:
        return "image_task"

    # Check table headers for classification
    if has_table and table_data:
        all_headers = []
        all_cells_text = []
        for table in table_data:
            if table and len(table) > 0:
                all_headers.extend([cell.strip() for cell in table[0]])
                for row in table:
                    all_cells_text.extend([cell.strip() for cell in row])

        headers_joined = " ".join(all_headers)
        cells_joined = " ".join(all_cells_text)

        # Conjugation table: has שורש or יחיד/רבים or עבר/הווה/עתיד as structure
        if any(kw in headers_joined for kw in ["שורש", "יחיד זכר", "רבים זכר", "יחיד נקבה"]):
            return "conjugation"

        # Tense-based table (עבר, הווה, עתיד as headers)
        if any(kw in headers_joined for kw in ["עבר", "הווה", "עתיד"]):
            # Could be exercise (תמלאו) or conjugation reference
            if "תמלאו" in all_text or "ואלמת" in all_text:
                return "exercise_table"
            # If cells have blanks, it's an exercise
            if any("______" in c or "" == c.strip() for c in all_cells_text[len(table_data[0]):]):
                # Check: many empty cells = exercise
                empty_cells = sum(1 for c in all_cells_text if c.strip() == "")
                if empty_cells > 3:
                    return "exercise_table"
            return "conjugation"

        # Shem peula table: check both headers and surrounding text
        if any(kw in headers_joined for kw in ["שם פעולה", "פעולה", "תרגום"]):
            if "שם" in headers_joined or "פעולה" in all_text or "הפועלה" in all_text:
                return "shem_peula"

        # Verb summary by gzarot
        if any(kw in headers_joined for kw in ["שלמים", "גרזה", "גזרה"]):
            return "verb_summary"

        # Vocabulary table: has category columns for parts of speech
        vocab_markers = ["שם תואר", "פעלים", "מיליות", "ביטויים"]
        if any(kw in headers_joined for kw in vocab_markers):
            return "vocabulary"
        # "שם פועל" alone with other POS columns = vocabulary
        if "שם פועל" in headers_joined and any(kw in headers_joined for kw in ["תואר", "פעלים"]):
            return "vocabulary"

    # Exercise detection by text patterns
    if "______" in all_text or "_______" in all_text:
        # Table-based exercise
        if has_table:
            # Check if it's oged exercise
            if "אוגד" in all_text or "דגוא" in all_text:
                return "exercise_oged"
            return "exercise_table"

        if "תעשו משפטים" in all_text or "טפשמ ושעת" in all_text:
            return "exercise_fill"

        if "תמלאו" in all_text or "ואלמת" in all_text:
            return "exercise_table"

        if "תשלימו" in all_text or "ומילשת" in all_text:
            return "exercise_fill"

        if "הוסיפו" in all_text or "ופיסוה" in all_text or "אוגד" in all_text or "דגוא" in all_text:
            return "exercise_oged"

        if "תבחרו" in all_text or "ורחבת" in all_text:
            return "exercise_choose"

        return "exercise_fill"

    # Translation exercise: numbered Russian sentences
    if "תתרגמו" in all_text or "ומגרתת" in all_text:
        return "exercise_translate"

    russian_numbered = re.findall(r'^\d+[\.\)]\s*[А-Яа-яЁё]', all_text, re.MULTILINE)
    if len(russian_numbered) >= 3 and not any(kw in all_text for kw in ["Правил", "Дело в том", "Итак", "К данному", "Приведем"]):
        return "exercise_translate"

    # Choose correct word
    if "תבחרו" in all_text or "ורחבת" in all_text or "הנוכנה הלימב ורחב" in all_text:
        return "exercise_choose"

    # Aggressive/non-aggressive classification exercise
    if "תיביסרגא" in all_text or "агрессив" in all_text_lower:
        return "exercise_choose"

    # Grammar detection for longer texts (before reading check)
    # This catches explanatory slides about specific topics like אוגד
    grammar_keywords_ru = [
        "связк", "подлежащ", "сказуем", "именн", "копул", "огэд",
        "Правил", "спряжени", "глагол", "породы", "породе", "корн",
        "префикс", "огласов", "биньян", "гзар", "каузатив",
        "Итак", "Дело в том", "Приведем", "означает",
        "Признак", "Обратите вниман",
    ]
    grammar_keywords_he = [
        "אוגד", "דגוא", "משפט שמני", "ינמש טפשמ",
        "בניין", "ןיינב", "גזרת", "תרזג", "הפעיל", "ליעפה",
    ]
    russian_chars = len(re.findall(r'[А-Яа-яЁё]', all_text))
    hebrew_chars = len(re.findall(r'[\u0590-\u05FF]', all_text))

    if russian_chars > 50:
        if any(kw in all_text for kw in grammar_keywords_ru) or any(kw in all_text for kw in grammar_keywords_he):
            return "grammar"

    # Jokes detection
    joke_markers = ["גולם", "חושם", "חלם", "לבהא", "חידה", "חידות", "םלוג", "םשוח", "םלח"]
    if any(m in all_text for m in joke_markers):
        # Only if it looks like a story, not just a vocabulary word
        if len(all_text) > 200:
            return "joke"

    # Reading text detection: long text, primarily Hebrew or dialog
    if len(all_text) > 500:
        # Check for dialog markers
        if all_text.count("-") >= 3 or all_text.count("—") >= 3:
            return "reading"
        # Check for story-like content
        hebrew_chars = len(re.findall(r'[\u0590-\u05FF]', all_text))
        total_alpha = len(re.findall(r'[\w]', all_text))
        if total_alpha > 0 and hebrew_chars / total_alpha > 0.5:
            return "reading"

    # Grammar: Russian explanatory text with Hebrew (fallback for non-keyword matches)
    if russian_chars > 50 and hebrew_chars > 10:
        # Numbered Russian sentences without grammar keywords = translation exercise
        russian_numbered = re.findall(r'^\d+[\.\)]\s*[А-Яа-яЁё]', all_text, re.MULTILINE)
        if len(russian_numbered) >= 3:
            return "exercise_translate"
        # Substantial Russian text with Hebrew = likely grammar explanation
        if russian_chars > 200:
            return "grammar"

    # Shem peula section
    if "שם פעולה" in all_text or "הלעופ םש" in all_text:
        return "shem_peula"

    # If has table but not classified yet
    if has_table and table_data:
        # Check for exercise-like tables
        for table in table_data:
            for row in table:
                for cell in row:
                    if "______" in cell:
                        return "exercise_table"

    # Fallback for slides with some content
    if len(all_text) > 100:
        if russian_chars > hebrew_chars:
            return "grammar"
        return "reading"

    if has_image:
        return "image_task"

    return "other"


def extract_table(shape):
    """Extract table data as list of lists."""
    table = shape.table
    rows = []
    for row in table.rows:
        cells = [cell.text.strip() for cell in row.cells]
        rows.append(cells)
    return rows


def extract_images(slide, slide_idx, output_dir):
    """Extract images from a slide, save to files, return list of paths."""
    images_dir = os.path.join(output_dir, "images")
    os.makedirs(images_dir, exist_ok=True)

    image_paths = []
    img_count = 0
    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            img_count += 1
            image = shape.image
            ext = image.content_type.split("/")[-1]
            if ext == "jpeg":
                ext = "jpg"
            filename = f"slide_{slide_idx + 1}_img_{img_count}.{ext}"
            filepath = os.path.join(images_dir, filename)
            with open(filepath, "wb") as f:
                f.write(image.blob)
            image_paths.append(filepath)

    return image_paths


def extract_slide_content(slide, slide_idx, output_dir):
    """Extract all content from a single slide."""
    texts = []
    tables = []
    has_image = False
    image_paths = []

    for shape in slide.shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            has_image = True

        if shape.has_text_frame:
            # Skip slide number and footer placeholders
            shape_name = shape.name if hasattr(shape, 'name') else ""
            is_placeholder = ("מספר שקופית" in shape_name or
                            "שקופית" in shape_name and "מציין" in shape_name or
                            "כותרת תחתונה" in shape_name)
            if is_placeholder:
                continue

            for para in shape.text_frame.paragraphs:
                text = para.text.strip()
                if text:
                    texts.append(text)

        if shape.has_table:
            tables.append(extract_table(shape))

    if has_image:
        image_paths = extract_images(slide, slide_idx, output_dir)

    category = classify_slide(texts, bool(tables), has_image, tables)

    return {
        "slide_number": slide_idx + 1,
        "category": category,
        "texts": texts,
        "tables": tables,
        "has_image": has_image,
        "image_paths": image_paths,
    }


def table_to_markdown(table_data):
    """Convert table data to markdown table."""
    if not table_data:
        return ""

    lines = []
    # Header
    header = table_data[0]
    lines.append("| " + " | ".join(header) + " |")
    lines.append("| " + " | ".join(["---"] * len(header)) + " |")

    # Rows
    for row in table_data[1:]:
        # Pad row if shorter than header
        padded = row + [""] * (len(header) - len(row))
        lines.append("| " + " | ".join(padded[:len(header)]) + " |")

    return "\n".join(lines)


def format_slide_md(slide_data):
    """Format a single slide's data as markdown."""
    lines = []
    cat = slide_data["category"]
    num = slide_data["slide_number"]

    # Category label
    category_labels = {
        "grammar": "📖 Грамматика",
        "vocabulary": "📝 Новые слова",
        "conjugation": "📊 Таблица спряжения",
        "verb_summary": "📊 Сводная таблица глаголов",
        "shem_peula": "📊 Отглагольные существительные (שם פעולה)",
        "exercise_fill": "✏️ Упражнение: вставить слово",
        "exercise_translate": "✏️ Упражнение: перевод",
        "exercise_table": "✏️ Упражнение: заполнить таблицу",
        "exercise_choose": "✏️ Упражнение: выбрать слово",
        "exercise_oged": "✏️ Упражнение: אוגד",
        "reading": "📖 Текст для чтения",
        "image_task": "🖼️ Задание по картинке",
        "joke": "😄 Анекдоты / юмор",
        "other": "📄 Прочее",
    }

    label = category_labels.get(cat, cat)
    lines.append(f"## Слайд {num} — {label}")
    lines.append("")

    # Image placeholder
    if slide_data["has_image"] and slide_data["image_paths"]:
        for img_path in slide_data["image_paths"]:
            rel_path = os.path.basename(img_path)
            lines.append(f"![Изображение](images/{rel_path})")
            lines.append("")
        if cat == "image_task":
            lines.append("<!-- CLAUDE: Опиши изображение и найди различия между двумя частями картинки -->")
            lines.append("")

    # Text content
    if slide_data["texts"]:
        for text in slide_data["texts"]:
            lines.append(text)
            lines.append("")

    # Tables
    if slide_data["tables"]:
        for table in slide_data["tables"]:
            lines.append(table_to_markdown(table))
            lines.append("")

    lines.append("---")
    lines.append("")

    return "\n".join(lines)


def extract_lesson_info(slides_data, filename=None):
    """Try to determine lesson number and part from slide content or filename."""
    # First try from slide content
    for slide in slides_data:
        for text in slide["texts"]:
            match = re.search(r'שיעור\s+(\d+)\s+חלק\s+(\d+)', text)
            if match:
                return int(match.group(1)), int(match.group(2))
            match = re.search(r'(\d+)\s+קלח\s+(\d+)\s+רועיש', text)
            if match:
                return int(match.group(2)), int(match.group(1))

    # Fallback: try filename
    if filename:
        # Match patterns like שיעור_13_חלק_3 or שיעור 13 חלק 3
        match = re.search(r'שיעור[_\s]+(\d+)[_\s]+חלק[_\s]+(\d+)', filename)
        if match:
            return int(match.group(1)), int(match.group(2))
        # Try with just numbers
        match = re.search(r'(\d+)[_\s]+חלק[_\s]+(\d+)', filename)
        if match:
            return int(match.group(1)), int(match.group(2))

    return None, None


def main():
    if len(sys.argv) < 3:
        print(f"Usage: {sys.argv[0]} <input.pptx> <output_dir>")
        sys.exit(1)

    input_path = sys.argv[1]
    output_dir = sys.argv[2]

    if not os.path.exists(input_path):
        print(f"ERROR: File not found: {input_path}")
        sys.exit(1)

    os.makedirs(output_dir, exist_ok=True)

    print(f"Opening: {input_path}")
    prs = Presentation(input_path)

    slides_data = []
    for idx, slide in enumerate(prs.slides):
        print(f"  Processing slide {idx + 1}...")
        data = extract_slide_content(slide, idx, output_dir)
        slides_data.append(data)
        print(f"    → {data['category']} ({len(data['texts'])} text blocks, {len(data['tables'])} tables, images: {data['has_image']})")

    # Determine lesson info
    filename = os.path.basename(input_path)
    lesson_num, part_num = extract_lesson_info(slides_data, filename)

    # Build markdown
    md_lines = []

    if lesson_num and part_num:
        md_lines.append(f"# Извлечение из урока {lesson_num}, часть {part_num}")
        md_lines.append(f"## שיעור {lesson_num} חלק {part_num}")
    else:
        md_lines.append("# Извлечение из урока")

    md_lines.append("")
    md_lines.append(f"Всего слайдов: {len(slides_data)}")
    md_lines.append("")

    # Summary of slide categories
    md_lines.append("### Структура презентации")
    md_lines.append("")
    for slide in slides_data:
        cat = slide["category"]
        label_map = {
            "grammar": "грамматика",
            "vocabulary": "словарь",
            "conjugation": "спряжение",
            "verb_summary": "сводка глаголов",
            "shem_peula": "שם פעולה",
            "exercise_fill": "упражнение (вставить)",
            "exercise_translate": "упражнение (перевод)",
            "exercise_table": "упражнение (таблица)",
            "exercise_choose": "упражнение (выбор)",
            "exercise_oged": "упражнение (אוגד)",
            "reading": "текст",
            "image_task": "картинка",
            "joke": "анекдоты",
            "other": "прочее",
        }
        md_lines.append(f"- Слайд {slide['slide_number']}: {label_map.get(cat, cat)}")
    md_lines.append("")
    md_lines.append("---")
    md_lines.append("")

    # All slides
    for slide in slides_data:
        md_lines.append(format_slide_md(slide))

    # Write output
    md_path = os.path.join(output_dir, "extracted.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write("\n".join(md_lines))

    # Write structured JSON for programmatic access
    json_path = os.path.join(output_dir, "extracted.json")
    # Remove image paths for JSON (they're filesystem-specific)
    json_data = []
    for slide in slides_data:
        entry = dict(slide)
        entry["image_filenames"] = [os.path.basename(p) for p in slide.get("image_paths", [])]
        del entry["image_paths"]
        json_data.append(entry)

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(json_data, f, ensure_ascii=False, indent=2)

    print(f"\nDone!")
    print(f"  Markdown: {md_path}")
    print(f"  JSON:     {json_path}")
    if any(s["has_image"] for s in slides_data):
        print(f"  Images:   {os.path.join(output_dir, 'images/')}")

    # Print category summary
    from collections import Counter
    cats = Counter(s["category"] for s in slides_data)
    print(f"\nCategory summary:")
    for cat, count in cats.most_common():
        print(f"  {cat}: {count}")


if __name__ == "__main__":
    main()
