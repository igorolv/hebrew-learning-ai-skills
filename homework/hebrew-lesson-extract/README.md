# hebrew-lesson-extract

Claude Skill для извлечения содержимого PPTX-уроков иврита в структурированный markdown.

## Что делает

- Парсит PPTX-файл урока с помощью `python-pptx`
- Извлекает текст с сохранением огласовок
- Выделяет таблицы и структуру слайдов
- Классифицирует слайды по типам: грамматика, словарь, упражнения, чтение, задания по картинкам и др.
- Сохраняет изображения из презентации отдельно
- Упаковывает результат в ZIP-архив

## Как использовать

### В Claude.ai (Projects)
Добавьте файл `SKILL.md` в Project Knowledge вашего проекта.

### В Claude Code
Скопируйте папку в директорию skills и подключите через конфигурацию.

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

ZIP-архив со структурированным результатом:
- `урок_N_часть_M_extracted.md`
- `images/` — извлечённые изображения из слайдов

## Часть пайплайна

Этот skill — первый этап homework-пайплайна:

1. **hebrew-lesson-extract** ← вы здесь
2. [hebrew-lesson-enrich](../hebrew-lesson-enrich) — создание шпаргалки
3. [hebrew-homework](../hebrew-homework) — выполнение домашних заданий
4. [hebrew-generate-docx](../hebrew-generate-docx) — сборка DOCX
