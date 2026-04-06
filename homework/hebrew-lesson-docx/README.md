# hebrew-lesson-docx

Claude Skill для преобразования markdown-шпаргалок (грамматических справочников) по ивриту в форматированный Word-документ (DOCX).

## Что делает

- Принимает markdown-файл — результат работы скила `hebrew-lesson-enrich`
- Генерирует форматированный DOCX с иерархией заголовков (разделы → подразделы → подподразделы)
- Оформляет таблицы спряжений и лексики (David 18pt для иврита, Arial 12pt для русского)
- Поддерживает blockquotes с синей левой границей для ключевых правил
- Обрабатывает контрастные пары (иврит + русский перевод)
- Корректно работает с RTL/BiDi для иврита
- Вставляет разрывы страниц перед крупными разделами

## Как использовать

### В Claude.ai (Projects)
Добавьте файл `SKILL.md` в Project Knowledge вашего проекта.

### В Claude Code
Скопируйте папку в директорию skills и подключите через конфигурацию.

## Структура

```text
hebrew-lesson-docx/
├── SKILL.md                                        — основная инструкция
├── README.md
├── scripts/
│   └── build_lesson_docx.py                        — генерация DOCX из markdown
└── references/
    └── ...                                         — эталонные примеры входного markdown
```

## Входные данные

Markdown-файл — результат работы скила `hebrew-lesson-enrich`.

Формат имени: `Шпаргалка_по_{N}_уроку_{M}_часть.md`

## Выходной файл

`Шпаргалка_по_{N}_уроку_{M}_часть.docx`

## Отличия от hebrew-homework-docx

| | Домашнее задание | Шпаргалка |
|---|---|---|
| Скилл | hebrew-homework-docx | hebrew-lesson-docx |
| Структура | Плоская (слайды) | Иерархическая (разделы) |
| Blockquotes | Нет | Есть |
| Списки | Нумерованные упражнения | Маркированные |
| Разрывы страниц | Нет | Да (между разделами) |
| Иврит в таблицах | 14pt | 18pt |
| Поля страницы | 2.54 см | 2.0 см |

## Часть пайплайна

Этот skill — ответвление от основного homework-пайплайна для генерации справочных DOCX:

1. [hebrew-lesson-extract](../hebrew-lesson-extract) — извлечение урока из PPTX
2. [hebrew-lesson-enrich](../hebrew-lesson-enrich) — создание грамматической шпаргалки
3. **hebrew-lesson-docx** ← вы здесь (принимает результат шага 2)

Основная ветка пайплайна (домашние задания) продолжается через [hebrew-homework-solve](../hebrew-homework-solve) → [hebrew-homework-docx](../hebrew-homework-docx).

## Зависимости

- Python 3
- `python-docx`

## Запуск

```bash
python scripts/build_lesson_docx.py path/to/Шпаргалка_по_14_уроку_2_часть.md
```

Для одного входного файла можно указать явный выход:

```bash
python scripts/build_lesson_docx.py path/to/input.md -o path/to/output.docx
```
