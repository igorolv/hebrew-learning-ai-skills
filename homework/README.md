# Homework · Hebrew Learning Pipeline

Набор скиллов для обработки уроков иврита: от PPTX до грамматической шпаргалки и полностью
решённого домашнего задания в Markdown и DOCX.

## Режимы

### Полный автоматический цикл

`hebrew-homework-pipeline` принимает новый PPTX, запускает все этапы, решает все найденные задания
и складывает результаты в `урок_N_часть_M_result/`.

### Отдельные этапы

1. `hebrew-lesson-extract` — PPTX → extracted Markdown + images.
2. `hebrew-lesson-enrich` — extracted Markdown → грамматическая шпаргалка.
3. `hebrew-homework-solve` — выбранные или все задания → решённый Markdown.
4. `hebrew-homework-docx` — решённый Markdown → DOCX.
5. `hebrew-lesson-docx` — шпаргалка Markdown → DOCX.

Граф, стыки, validation gates и конвенции именования описаны в [PIPELINE.md](PIPELINE.md).

## Что проверяется

- Extract: полнота слайдов, классификация, таблицы, изображения и никуд источника.
- Enrich: грамматика, парадигмы, словарь, кумулятивные таблицы и никуд.
- Solve: полнота пунктов, грамматические формы, переводы и учебный никуд.
- DOCX: структурный контракт RTL/BiDi, шрифтов, таблиц и разрывов страниц.
- Pipeline: комплектность общей выходной директории и всех финальных артефактов.

## Структура

```text
homework/
├── README.md
├── PIPELINE.md
├── hebrew-homework-pipeline/
├── hebrew-lesson-extract/
├── hebrew-lesson-enrich/
├── hebrew-homework-solve/
├── hebrew-homework-docx/
└── hebrew-lesson-docx/
```
