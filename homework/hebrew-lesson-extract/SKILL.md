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
2. Classifies each slide into one of 20 categories (grammar, exercises, reading, etc.)
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
- `урок_<N>_часть_<M>_extracted.md` — structured markdown with all slide content
- `images/` — extracted slide images

If lesson/part numbers cannot be determined, falls back to `extracted.md`.

### Output filename convention

The output filename is derived from the **input PPTX filename**.
The script parses the lesson number (`N`) and part number (`M`) from:
1. Slide content (regex: `שיעור <N> חלק <M>`)
2. Filename (fallback, e.g. `שיעור_13_חלק_3.pptx`)

| Input file | Output file |
|---|---|
| `שיעור 13 חלק 3.pptx` | `урок_13_часть_3_extracted.md` |
| `שיעור_2_חלק_1.pptx` | `урок_2_часть_1_extracted.md` |
| `unknown.pptx` | `extracted.md` |

### Step 2: Review the extraction result

Read `урок_<N>_часть_<M>_extracted.md` and check:
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

### Step 4: Package and present the result

Package the extracted markdown and images into a single ZIP archive.
The ZIP filename follows the same convention: `урок_<N>_часть_<M>_extracted.zip`.

```bash
cd /home/claude/lesson_extract
zip -r "/mnt/user-data/outputs/урок_<N>_часть_<M>_extracted.zip" "урок_<N>_часть_<M>_extracted.md" images/
```

The ZIP contains:
- `урок_<N>_часть_<M>_extracted.md` — the markdown file (with `![Изображение](images/...)` links)
- `images/` — all extracted slide images

This way the user gets a single downloadable file. When unpacked, relative image paths
in the markdown resolve correctly.

Present the ZIP to the user using `present_files`.

## Slide categories

The script classifies slides into these categories:

### Reference / Grammar
| Category | Emoji | What it detects |
|---|---|---|
| `grammar` | 📖 | Russian explanatory text about grammar with Hebrew examples |
| `vocabulary` | 📝 | Table of new words (מילים חדשות) |
| `conjugation` | 📊 | Verb conjugation tables (with שורש, tense headers) |
| `verb_summary` | 📊 | Summary table grouping verbs by גזרות |
| `shem_peula` | 📊 | Verbal noun (שם פעולה) tables |
| `numerals` | 🔢 | Numeral reference tables (שם מספר) from 10 to 90000 |
| `prepositions` | 📖 | Preposition declension tables (מילת יחס + suffixes) |

### Exercises
| Category | Emoji | What it detects |
|---|---|---|
| `exercise_fill` | ✏️ | Fill-in-the-blank exercises (blanks in text) |
| `exercise_translate` | ✏️ | Translation exercises (Russian sentences to translate) |
| `exercise_table` | ✏️ | Table-based exercises (fill empty table cells) |
| `exercise_choose` | ✏️ | Choose the correct word from options in parentheses |
| `exercise_oged` | ✏️ | אוגד exercises (copula, specifically) |
| `exercise_morphology` | ✏️ | Determine root/binyan/gzara from a verb form |
| `exercise_transform` | ✏️ | Rewrite sentences in a different tense (no blanks) |
| `exercise_writing` | ✏️ | Write a composition / essay (חיבור) |
| `exercise_questions` | ✏️ | Answer questions about a reading text |

### Texts / Other
| Category | Emoji | What it detects |
|---|---|---|
| `reading` | 📖 | Hebrew reading texts, dialogs, stories |
| `image_task` | 🖼️ | Slides with primarily images (spot-the-difference, etc.) |
| `joke` | 😄 | Jokes and humorous texts (גולם, חושם, חלם) |
| `other` | 📄 | Unclassified slides |

## Classification logic (summary)

The classifier works in phases, from most specific to most general:

1. **Image-only** — very little text + image + no table → `image_task`
2. **Instruction keywords** — early detection of exercise types by Hebrew instructions
   (תתרגמו, תשלימו, תכתבו חיבור, תעשו לפי הדוגמה, תבחרו, etc.)
3. **Table-based** — analyze table headers and empty cells to determine type
4. **Blanks in text** — `_{2,}` pattern detects fill-in exercises
5. **Translation detection** — numbered or unnumbered Russian sentences
6. **Transform detection** — instruction keyword + no blanks = rewrite exercise
7. **Grammar keywords** — Russian explanatory text with grammar terminology
8. **Reading** — long predominantly-Hebrew text
9. **Fallbacks** — by Russian/Hebrew character ratio

## Important notes

- **Nikud preservation**: The PPTX format preserves nikud much better than PDF.
  The script extracts text as-is, keeping all vowel points intact.
- **RTL text**: Hebrew text may appear in different directions depending on the
  PPTX creation tool. The script extracts raw text without reordering.
- **Underscore detection**: The script detects blanks as short as `__` (2+ underscores),
  not just long `______` sequences. This catches all common blank formats.
- **Classification is heuristic**: The script uses keyword and pattern matching.
  Some slides may be misclassified. Review and correct if needed.
- **No enrichment**: This skill does NOT add grammar explanations, historical
  commentary, or additional examples. That's done by the enrichment skill.
