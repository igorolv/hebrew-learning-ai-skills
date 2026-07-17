# Harry Potter Pipeline

Формальное описание pipeline для AI-агентов. Человекочитаемое описание — в [README.md](README.md).

## Что такое pipeline

Pipeline — упорядоченный граф скилов, где выход одного скила является входом следующего. Скилы выполняются последовательно; между шагами AI показывает результат пользователю и ждёт подтверждения.

---

## Граф

Pipeline состоит из двух веток — текстовой и иллюстрационной — которые сходятся на финальном шаге.

```
 PDF-исходник
      |
      v
 [hp-extraction]
      |
      | HP_ch{N}_{FROM}_{TO}.md
      v
 [hp-translate] <-- [hp-source-texts]
      |
      | HP_ch{N}_{FROM}_{TO}_translate.md
      +------------------------+------------------------------+
                               |                              |
                               v                              |
                      [hp-generate-image]                     |
                               ^                              |
                               | BOOK_{N}_CHAPTER_{M}_STYLE.md|
                               |                              |
 [hp-master-style] ---> [hp-chapter-style]                   |
         |                     |                              |
         +---------------------+                              |
       master_style_framework.md                              |
                               |                              |
                               | HP_ch{CHAPTER}_page_{PAGE}.png|
                               +------------------------------+
                                              |
                                              v
                                     [hp-generate-docx]
                                              |
                                              v
                    Гарри Поттер глава {N} страницы {FROM}-{TO}.docx
```

## Ветки

### Текстовая ветка

```
extraction -> translate -> generate-docx
```

`hp-source-texts` — утилита, вызываемая автоматически из `hp-translate`.
Translated markdown передаётся одновременно в `hp-generate-image` и `hp-generate-docx`.

### Иллюстрационная ветка

```
master-style -> chapter-style ─┐
                              ├-> generate-image -> generate-docx
translate --------------------┘
```

`master-style` выполняется один раз на весь проект (setup). `chapter-style` — один раз на главу. `generate-image` — по одному разу на каждую страницу.

### Точка слияния

`hp-generate-docx` принимает выходы обеих веток:
- markdown из текстовой ветки
- список PNG-иллюстраций из иллюстрационной ветки

---

## Стыки: что передаётся между скилами

### extraction -> translate

| | |
|---|---|
| **Артефакт** | `HP_ch{N}_{FROM}_{TO}.md` |
| **Содержимое** | Ивритский текст с частичными огласовками + таблицы сложных слов, разбитые по страницам |

### hp-source-texts -> translate

| | |
|---|---|
| **Артефакт** | Фрагменты из `HP_en_ch{N}.md` и `HP_rosmen_ch{N}.md` |
| **Содержимое** | Английский оригинал и русский перевод Росмэн, соответствующие обрабатываемым страницам |
| **Где хранятся** | Project Knowledge |

### translate -> generate-docx

| | |
|---|---|
| **Артефакт** | `HP_ch{N}_{FROM}_{TO}_translate.md` |
| **Содержимое** | Полный учебный markdown: иврит с огласовками, подстрочник, литературный перевод, сложные слова, различия переводов |

### translate -> generate-image

| | |
|---|---|
| **Артефакт** | `HP_ch{N}_{FROM}_{TO}_translate.md` |
| **Дополнительный вход** | Целевой номер `PAGE` |
| **Содержимое** | Блок `# Страница PAGE`; события берутся из секций «Иврит» и «Подстрочный перевод» |
| **Правило** | Литературный перевод используется как контекст и не добавляет неподтверждённые визуальные детали |

### master-style -> chapter-style, generate-image

| | |
|---|---|
| **Артефакт** | `master_style_framework.md` |
| **Содержимое** | Неизменяемый каркас визуального стиля: философия, палитра, техника, запреты, шкала магии |
| **Где хранится** | Project Knowledge (создаётся один раз) |

### chapter-style -> generate-image

| | |
|---|---|
| **Артефакт** | `BOOK_{N}_CHAPTER_{M}_STYLE.md` |
| **Содержимое** | Персонажи (эталоны), локации, список сцен с ID, эмоциональные состояния |

### generate-image -> generate-docx

| | |
|---|---|
| **Артефакт** | Список отдельных PNG-иллюстраций |
| **Имя каждого файла** | `HP_ch{CHAPTER}_page_{PAGE}.png` |
| **Содержимое** | По одной готовой иллюстрации на каждую страницу markdown |
| **Особенность** | `hp-generate-image` сразу вызывает генератор изображений; ZIP и ручная упаковка не используются |

---

## Данные от пользователя vs между скилами

### От пользователя (внешние входы)

| Что | На каком шаге | Обязательно |
|-----|--------------|-------------|
| PDF-исходник (ZIP с JPEG) | extraction | да |
| Номер главы + диапазон страниц | extraction | да |
| Текст главы на русском (Росмэн) | chapter-style | да |
| Английский оригинал (`HP_en_ch{N}.md`) | translate (через source-texts) | да, в Project Knowledge |
| Русский Росмэн (`HP_rosmen_ch{N}.md`) | translate (через source-texts) | да, в Project Knowledge |
| Целевой номер страницы `PAGE` | generate-image | да |
| Список готовых PNG-иллюстраций | generate-docx | да |

### Между скилами (внутренние артефакты)

| Артефакт | Откуда | Куда |
|----------|--------|------|
| `HP_ch{N}_{FROM}_{TO}.md` | extraction | translate |
| `HP_ch{N}_{FROM}_{TO}_translate.md` | translate | generate-image, generate-docx |
| `master_style_framework.md` | master-style | chapter-style, generate-image |
| `BOOK_{N}_CHAPTER_{M}_STYLE.md` | chapter-style | generate-image |
| `HP_ch{CHAPTER}_page_{PAGE}.png` | generate-image | generate-docx |

---

## Порядок запуска

### Текстовая ветка

```
1. extraction     — первый шаг
2. translate      — после extraction (автоматически вызывает source-texts)
```

### Иллюстрационная ветка

```
0. master-style   — один раз на проект (setup)
1. chapter-style  — один раз на главу
2. generate-image — после translate и chapter-style, по одному готовому PNG на страницу
```

### Финальная сборка

```
3. generate-docx  — после завершения обеих веток
```

`translate` и `chapter-style` можно готовить параллельно. `generate-image`
запускается только после готовности обоих их выходов.

---

## Утилиты

### hp-source-texts

Вспомогательный скил. Не вызывается пользователем напрямую — используется внутри `hp-translate` для поиска английского оригинала и русского перевода Росмэн в Project Knowledge.

### hp-master-style

Скил управления жизненным циклом `master_style_framework.md`. Выполняется один раз при старте проекта. Команды: `create`, `validate`, `extend`, `show`.

---

## Что делать, если предыдущий скил не выполнен

Каждый скил проверяет наличие входных данных. Если артефакт предыдущего скила не найден:

1. Сообщить пользователю, какой артефакт отсутствует
2. Предложить запустить предыдущий скил
3. Не пытаться выполнить текущий скил без входных данных

Для `master_style_framework.md` — предложить запустить `hp-master-style` с командой `create`.

---

## Конвенции именования

| Артефакт | Формат имени |
|----------|-------------|
| Извлечённый иврит | `HP_ch{N}_{FROM}_{TO}.md` |
| Переведённый файл | `HP_ch{N}_{FROM}_{TO}_translate.md` |
| Master style | `master_style_framework.md` |
| Chapter style | `BOOK_{N}_CHAPTER_{M}_STYLE.md` |
| Иллюстрация страницы | `HP_ch{CHAPTER}_page_{PAGE}.png` |
| Английский оригинал | `HP_en_ch{N}.md` |
| Русский Росмэн | `HP_rosmen_ch{N}.md` |
| Готовый DOCX | `Гарри Поттер глава {N} страницы {FROM}-{TO}.docx` |
