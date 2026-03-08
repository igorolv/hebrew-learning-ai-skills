# Hebrew Learning AI Skills

Репозиторий Claude Skills для двух рабочих направлений по изучению иврита:

1. **Homework** — обработка уроков и выполнение домашних заданий
2. **Literary Club** — создание учебных материалов по ивритскому «Гарри Поттеру»

Все skill'ы оформлены как самостоятельные папки с `SKILL.md`, а для большинства пайплайнов и отдельных шагов есть README с назначением, входами, выходами и местом в общем процессе.

---

## Структура репозитория

```text
hebrew-learning-ai-skills/
├── README.md
├── homework/
│   ├── README.md
│   ├── hebrew-lesson-extract/
│   ├── hebrew-lesson-enrich/
│   ├── hebrew-homework/
│   └── hebrew-generate-docx/
└── literary-club/
    └── harry-potter/
        ├── README.md
        ├── hp-master-style/
        ├── hp-chapter-style/
        ├── hp-generate-image/
        ├── hp-extraction/
        ├── hp-source-texts/
        ├── hp-translate/
        └── hp-generate-docx/
```

---

## Раздел 1 · Homework

Папка `homework/` содержит pipeline для работы с обычными уроками иврита и домашними заданиями.

### Что внутри

- `hebrew-lesson-extract` — извлечение содержимого PPTX-урока в markdown
- `hebrew-lesson-enrich` — создание грамматической шпаргалки по extracted-файлу
- `hebrew-homework` — выполнение упражнений по выбранным слайдам
- `hebrew-generate-docx` — конвертация выполненного задания в DOCX

### Логика работы

```text
PPTX урока
   → hebrew-lesson-extract
   → hebrew-lesson-enrich
   → hebrew-homework
   → hebrew-generate-docx
```

Итог: готовый Word-документ с выполненным домашним заданием.

Подробности см. в [homework/README.md](homework/README.md).

---

## Раздел 2 · Literary Club / Harry Potter

Папка `literary-club/harry-potter/` содержит pipeline для подготовки учебных материалов по ивритскому изданию «Гарри Поттера».

### Что внутри

#### Текстовая ветка
- `hp-extraction` — извлечение ивритского текста из PDF-исходника
- `hp-source-texts` — поиск английского оригинала и перевода Росмэн
- `hp-translate` — добавление огласовок, подстрочного и литературного перевода, комментариев
- `hp-generate-docx` — сборка финального DOCX с иллюстрациями

#### Иллюстрационная ветка
- `hp-master-style` — управление master style framework проекта
- `hp-chapter-style` — поглавный визуальный стайлгайд
- `hp-generate-image` — генерация промтов для иллюстраций

### Логика работы

```text
Текстовая ветка:
PDF/ZIP → hp-extraction → hp-translate → hp-generate-docx

Иллюстрационная ветка:
hp-master-style → hp-chapter-style → hp-generate-image → hp-generate-docx
```

Итог: учебный DOCX по страницам книги с переводами, комментариями и иллюстрациями.

Подробности см. в [literary-club/harry-potter/README.md](literary-club/harry-potter/README.md).

---

## Как пользоваться репозиторием

### В Claude.ai (Projects)
Для нужного направления добавляйте в Project Knowledge:
- соответствующие `SKILL.md`
- reference-файлы, если они нужны конкретному skill'у
- входные материалы проекта

### В Claude Code
Папки skills можно подключать как локальные навыки и использовать по отдельности или как последовательный pipeline.

---

## Принцип организации

У каждого skill'а обычно есть:
- `SKILL.md` — основная инструкция
- `README.md` — краткое описание skill'а
- `references/` — образцы, шаблоны и вспомогательные правила
- дополнительные скрипты, если шаг требует локальной обработки файлов

Репозиторий организован не по техническому стеку, а по рабочим сценариям: отдельно домашние задания, отдельно литературный клуб.
