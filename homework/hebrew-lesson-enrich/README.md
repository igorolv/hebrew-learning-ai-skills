# hebrew-lesson-enrich

Claude Skill для превращения сырого extracted-markdown урока иврита в подробную грамматическую шпаргалку.

## Что делает

- Принимает результат работы `hebrew-lesson-extract`
- Выделяет ключевую грамматику урока
- Собирает новые слова, глаголы, биньяны, гзарот и отглагольные существительные
- Добавляет объяснения, таблицы и примеры
- Формирует единый markdown-файл-шпаргалку для дальнейшей работы и домашнего задания

## Как использовать

### В Claude.ai (Projects)
Добавьте файл `SKILL.md` в Project Knowledge вашего проекта.

### В Claude Code
Скопируйте папку в директорию skills и подключите через конфигурацию.

## Структура

```text
hebrew-lesson-enrich/
├── SKILL.md                         — основная инструкция
├── README.md
└── references/
    ├── enrichment_guidelines.md     — правила лингвистического обогащения
    └── output_structure.md          — структура итоговой шпаргалки
```

## Входные данные

Markdown-файл формата `урок_N_часть_M_extracted.md`, созданный skill'ом `hebrew-lesson-extract`.

## Выходной файл

`Шпаргалка_по_N_уроку_M_часть.md`

Пример:
- `Шпаргалка_по_13_уроку_3_часть.md`

## Часть пайплайна

Этот skill — второй этап в homework-пайплайне:

1. [hebrew-lesson-extract](../hebrew-lesson-extract) — извлечение урока из PPTX
2. **hebrew-lesson-enrich** ← вы здесь
3. [hebrew-homework-solve](../hebrew-homework-solve) — выполнение домашних заданий
4. [hebrew-homework-docx](../hebrew-homework-docx) — сборка DOCX
