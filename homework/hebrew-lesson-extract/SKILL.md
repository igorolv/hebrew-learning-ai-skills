---
name: hebrew-lesson-extract
description: >
  Extract structured content from Hebrew lesson PPTX files into organized Markdown.
  Use this skill whenever the user uploads a Hebrew lesson presentation (.pptx) and asks to
  extract, parse, or convert its content. Triggers include: uploading a שיעור (lesson) file,
  asking to "извлечь урок", "разобрать презентацию", "сделать шпаргалку из урока",
  "extract lesson", or any request involving a Hebrew lesson PPTX file.
  This skill handles the EXTRACTION step only — it produces a structured "raw" markdown
  with classified slides. A separate enrichment skill turns this into a full grammar cheat sheet.
  Also use this skill when the user asks to process, read, or analyze a Hebrew lesson PPTX,
  even if they don't explicitly say "extract".
---

# Hebrew Lesson PPTX → Structured Markdown

This skill extracts content from Hebrew lesson PPTX presentations and produces
a structured Markdown file with classified slides.

## What it does

1. Parses the PPTX file using `python-pptx`
2. Classifies each slide (grammar, vocabulary, exercises, reading, images, jokes, etc.)
3. Extracts text with **nikud preserved** and tables as markdown
4. Extracts images and saves them as files
5. For image-based tasks (e.g. spot-the-difference), Claude describes the image directly

## Workflow

### Step 1: Install dependency and run extraction

```bash
pip install python-pptx --break-system-packages -q
python3 /path/to/scripts/extract_pptx.py "<uploaded_pptx_path>" /home/claude/lesson_extract
```

The script produces:
- `extracted.md` — structured markdown with all slide content
- `extracted.json` — same data as JSON (for programmatic use)
- `images/` — extracted slide images

### Step 2: Review the extraction result

Read `extracted.md` and check:
- Are slide categories correct? (see reference: `references/output_format.md`)
- Is all text content preserved?
- Are tables formatted properly?

### Step 3: Process image tasks

For each slide classified as `image_task`:
1. View the extracted image file(s) using the `view` tool
2. Describe what you see in the image
3. If it's a "spot the difference" task (two similar images), list the differences
4. Replace the `<!-- CLAUDE: ... -->` marker in the markdown with your description

Write descriptions in **Russian** with Hebrew terms where appropriate.

Example output for an image task:

```markdown
## Слайд 1 — 🖼️ Задание по картинке

![Изображение](images/slide_1_img_1.jpg)

**Описание:** Дедушка и мальчик сидят на диване и играют в видеоигры.
На столе перед ними — две чашки чая, книга и зелёная бутылка.

**Различия между верхней и нижней картинками:**
1. На верхней картинке мальчик хмурится, на нижней — улыбается
2. Подушки поменялись местами (красная слева/справа)
3. Картина на стене изменилась
...
```

### Step 4: Save and present the result

Copy the final `extracted.md` (with image descriptions filled in) to
`/mnt/user-data/outputs/` and present it to the user.

Also copy images if they exist:
```bash
cp -r /home/claude/lesson_extract/images/ /mnt/user-data/outputs/images/
```

## Slide categories

The script classifies slides into these categories:

| Category | What it detects |
|---|---|
| `grammar` | Russian explanatory text about grammar with Hebrew examples |
| `vocabulary` | Table of new words (מילים חדשות) |
| `conjugation` | Verb conjugation tables (with שורש, tense headers) |
| `verb_summary` | Summary table grouping verbs by גזרות |
| `shem_peula` | Verbal noun (שם פעולה) tables |
| `exercise_*` | Various exercise types (fill-in, translate, table, choose) |
| `reading` | Hebrew reading texts, dialogs, stories |
| `image_task` | Slides with primarily images (spot-the-difference, etc.) |
| `joke` | Jokes and humorous texts (גולם, חושם, חלם) |

## Important notes

- **Nikud preservation**: The PPTX format preserves nikud much better than PDF.
  The script extracts text as-is, keeping all vowel points intact.
- **RTL text**: Hebrew text may appear in different directions depending on the
  PPTX creation tool. The script extracts raw text without reordering.
- **Classification is heuristic**: The script uses keyword and pattern matching.
  Some slides may be misclassified. Review and correct if needed.
- **No enrichment**: This skill does NOT add grammar explanations, historical
  commentary, or additional examples. That's done by the enrichment skill.
