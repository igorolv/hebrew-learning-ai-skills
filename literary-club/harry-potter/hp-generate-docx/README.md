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
├── scripts/
│   └── build_hp_docx.py    — скрипт генерации DOCX
└── references/
    ├── HP_ch1_30_35_translate.md    — образец входного файла (стр. 30–35)
    └── HP_ch1_36_37_translate.md    — образец входного файла (стр. 36–37)
```

## Входные данные

1. Markdown-файл формата `HP_ch{CHAPTER}_{FROM}_{TO}_translate.md` (результат работы `hp-translate`)
2. ZIP-архив с иллюстрациями — изображения, созданные на этапе `hp-generate-image` и собранные в архив (каждый файл содержит номер страницы в имени)

## Запуск

Зависимости:

```bash
pip install python-docx
```

Базовый запуск (с рендером через LibreOffice):

```bash
python3 scripts/build_hp_docx.py HP_ch3_1_2_translate.md Картинка_1-2.zip
```

С явным выходным файлом:

```bash
python3 scripts/build_hp_docx.py HP_ch3_1_2_translate.md Картинка_1-2.zip -o out.docx
```

Без рендера через LibreOffice:

```bash
python3 scripts/build_hp_docx.py HP_ch3_1_2_translate.md Картинка_1-2.zip --no-render
```

Имя выходного файла формируется автоматически: `HP_ch3_1_2_translate.md` → `Гарри Поттер глава 3 страницы 1-2.docx`.

## Часть проекта

Этот skill — финальный этап пайплайна, объединяющий результаты текстовой и иллюстрационной веток:

1. [hp-extraction](../hp-extraction) — извлечение текста из PDF
2. [hp-translate](../hp-translate) — добавление огласовок и переводов
3. [hp-chapter-style](../hp-chapter-style) — визуальный стайлгайд главы
4. [hp-generate-image](../hp-generate-image) — генерация промтов для иллюстраций
5. **hp-generate-docx** ← вы здесь (принимает результаты шагов 2 и 4)
