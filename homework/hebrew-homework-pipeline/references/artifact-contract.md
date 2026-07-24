# Контракт артефактов полного конвейера

| Этап | Обязательный вход | Выход |
|---|---|---|
| extract | PPTX урока | `урок_N_часть_M_extracted/` (`.md` + `images/`) |
| enrich | extracted Markdown | `Шпаргалка_по_N_уроку_M_часть.md` |
| solve all | extracted Markdown + шпаргалка | `ДЗ_урок_N_часть_M_слайды_*.md` |
| lesson DOCX | шпаргалка Markdown | шпаргалка DOCX, прошедшая структурную проверку |
| homework DOCX | решённое ДЗ Markdown | ДЗ DOCX, прошедшее структурную проверку |
| collect | все доступные финальные файлы | `урок_N_часть_M_result/` |

В `all_tasks` входят категории `exercise_*`, `reading`, `joke`, `image_task` и вручную
подтверждённые задания, ошибочно классифицированные как `other`.
