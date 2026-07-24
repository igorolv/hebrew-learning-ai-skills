# hebrew-lesson-extract

Скилл для извлечения содержимого PPTX-уроков иврита в структурированный markdown.

## Что делает

- Парсит PPTX-файл урока с помощью `python-pptx`
- Извлекает текст с сохранением огласовок
- Выделяет таблицы и структуру слайдов
- Классифицирует слайды по типам: грамматика, словарь, упражнения, чтение, задания по картинкам и др.
- Сохраняет изображения из презентации отдельно
- Сохраняет результат в общей выходной директории для следующих этапов

## Как использовать

Подключите папку как skill в используемой AI-среде либо вызовите bundled Python-скрипт напрямую.

## Структура

```text
hebrew-lesson-extract/
├── SKILL.md                     — основная инструкция
├── README.md
├── references/
│   └── output_format.md         — формат extracted-markdown
└── scripts/
    └── extract_pptx.py          — вспомогательный скрипт извлечения
```

## Входные данные

PPTX-файл урока иврита, обычно вида `שיעור N חלק M.pptx`.

## Выходные данные

Директория `урок_N_часть_M_extracted/`:

- `урок_N_часть_M_extracted.md`;
- `images/` — извлечённые изображения из слайдов.

## Часть пайплайна

Этот skill — первый этап homework-пайплайна:

1. **hebrew-lesson-extract** ← вы здесь
2. [hebrew-lesson-enrich](../hebrew-lesson-enrich) — создание шпаргалки
3. [hebrew-homework-solve](../hebrew-homework-solve) — выполнение домашних заданий
4. [hebrew-homework-docx](../hebrew-homework-docx) — сборка DOCX

Полный автоматический цикл запускает [hebrew-homework-pipeline](../hebrew-homework-pipeline).
