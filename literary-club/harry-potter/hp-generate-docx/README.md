# hp-generate-docx

Claude Skill для преобразования учебных markdown-файлов проекта Hebrew Harry Potter в DOCX.

## Что делает

- Принимает markdown-файл с переведёнными страницами и ZIP-архив с иллюстрациями
- Генерирует форматированный DOCX с правильной типографикой для иврита (David 18pt, RTL)
- Вставляет иллюстрации на каждую страницу
- Конвертирует markdown-таблицы в таблицы Word
- Проставляет разрывы страниц между секциями

## Как использовать

### В Claude.ai (Projects)
Добавьте файл `SKILL.md` в Project Knowledge вашего проекта.

### В Claude Code
Скопируйте папку в директорию skills и подключите через конфигурацию.

## Структура

```
hp-generate-docx/
├── SKILL.md                — основная инструкция
├── README.md
└── references/
    ├── HP_ch1_30_35_translate.md    — образец входного файла (стр. 30–35)
    └── HP_ch1_36_37_translate.md    — образец входного файла (стр. 36–37)
```

## Входные данные

1. Markdown-файл формата `HP_ch{CHAPTER}_{FROM}_{TO}_translate.md` (результат работы `hp-translate`)
2. ZIP-архив с иллюстрациями (png/jpg/jpeg/webp)

## Часть проекта

Этот skill — финальный этап подготовки печатных материалов:

1. [hp-extraction](../hp-extraction) — извлечение текста из PDF
2. [hp-translate](../hp-translate) — добавление огласовок и переводов
3. **hp-generate-docx** ← вы здесь
