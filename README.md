# Moodle XML Converter v3

Конвертер Word (.docx) в Moodle XML с графическим интерфейсом.

## Файлы

| Файл | Описание |
|------|----------|
| `converter_gui.py` | GUI-приложение (PyQt5) |
| `universal_moodle_converter_v3_stable.py` | Ядро конвертера (CLI + библиотека) |
| `table_compare.py` | Утилита сравнения с эталонными XML |
| `Шаблоны вопросов.docx` | Документация по разметке Word-файлов |

## Зависимости

```
pip install PyQt5 lxml python-docx docxlatex
```

## Запуск

### GUI
```
python converter_gui.py
```

### CLI (пакетная обработка)
```
python universal_moodle_converter_v3_stable.py <путь к docx или папке> --output-folder <папка>
```

### Сравнение с эталонами
```
python table_compare.py
```

---

## Архитектура

### Структура Word-файла

```
V1: Название учебного предмета (категория)
{маркер}V2: Название блока (подкатегория)

I:Задание N. Автор И.О., ТЗX-Y, b=N
S: Текст вопроса
+: Правильный ответ
-: Неправильный ответ
```

### Маркеры типов вопросов

Маркер ставится перед `V2:` в формате `{маркер}V2: Описание блока`.
Все вопросы внутри блока наследуют маркер до следующего `V2:`.

| Маркер | Тип Moodle | Описание |
|--------|-----------|----------|
| `{multichoice_one}` | multichoice (single=true) | Один правильный ответ. `+:` = 100%, `-:` = 0% |
| `{multichoice_many}` | multichoice (single=false) | Несколько правильных. Штраф **-100%** за каждый неправильный |
| `{shortanswer_phrase}` | shortanswer | Текстовый ввод. Несколько `+:` = несколько допустимых ответов |
| `{shortanswer_partial}` | shortanswer | Выбор нескольких ответов (нумерация 1)2)3)...). Все перестановки с partial scoring: 100%/50%/0% |
| `{shortanswer_numcombo}` | shortanswer | Выбор нескольких ответов. Все перестановки позиций = 100% |
| `{matching}` / `{match}` | matching | Соотношение. Формат `L1:` / `R1:`. Лишние R = дистракторы |
| `{match_123}` | matching | Последовательность. Формат `N: фраза` -> фраза сопоставляется с номером |
| `{ddmatch}` | ddmatch | Drag-and-drop. Формат `L1:` / `R1:` |
| `{gapselect}` | gapselect | Выпадающие списки. Текст с `(N)`, варианты `A)...D)`, ключ `+:ABCD` |
| `{cloze}` | cloze | Встроенные ответы `{1:SHORTANSWER:=answer}` |
| `{numerical}` | shortanswer | Числовой ответ. Генерирует два варианта: с `.` и с `,` |

Если маркер не указан, тип определяется эвристикой по содержимому.

### Форматы заголовков вопросов (7 вариантов)

Конвертер распознает 7 форматов начала вопроса:
1. `I:Задание N.` — стандартный
2. `I I:Задание N.` — двойной I (артефакт Word)
3. `I Задание N.` — пробел вместо двоеточия
4. `:Задание N.` — потерян символ I
5. `Задание N. Автор, ТЗX-Y, b=N` — без префикса I:
6. `Kn-=mЗадание N.` — мусор перед Задание
7. `Автор И.О., ТЗX-Y, b=N` — только автор (без слова Задание)

---

## GUI: converter_gui.py

### Возможности

1. **Выбор файла** — кнопка "Обзор" для .docx
2. **Выбор папки вывода** — куда сохранить XML
3. **Предпросмотр** (QTreeWidget):
   - Список всех вопросов с раскрывающимся содержимым
   - При клике на вопрос — раскрывается тело: текст (S:), правильные (+:, зелёные), неправильные (-:, красные), L/R пары
   - Комбобокс маркера — можно изменить маркер для блока
   - Цветовая кодировка по типу маркера
   - Подсветка ошибок красным
4. **Предобработка ошибок**:
   - Отсутствие правильного ответа
   - Пустой текст вопроса
   - Неизвестный маркер
5. **Конвертация** в отдельном потоке с прогресс-баром
6. **Постобработка XML**:
   - Проверка корневого элемента (`quiz`)
   - Проверка типов вопросов (только допустимые Moodle типы)
   - Проверка base64 картинок (не пустые)
   - Проверка маркеров `_IMAGE_` / `@@PLUGINFILE@@` без файлов
   - Проверка структуры matching (subquestion/answer)
   - Проверка gapselect (selectoption)
   - Проверка наличия ответов
7. **Разделение XML** на части до 1 МБ (чекбокс)

### Цветовая схема маркеров

| Цвет | Маркеры |
|------|---------|
| Голубой | multichoice_one, multichoice_many |
| Зелёный | shortanswer_phrase, shortanswer_partial, shortanswer_numcombo |
| Оранжевый | matching, match_123, match |
| Красноватый | ddmatch |
| Фиолетовый | gapselect |
| Жёлтый | cloze |
| Бирюзовый | numerical |

---

## Ядро: universal_moodle_converter_v3_stable.py

### Классы

- **`ImageProcessor`** — извлечение изображений из docx (base64)
- **`FormulaProcessor`** — конвертация LaTeX формул (`$...$` -> `\(...\)`)
- **`QuestionTypeDetector`** — определение типа вопроса (маркер приоритетнее эвристики)
- **`XMLGenerator`** — генерация Moodle XML:
  - `create_multichoice(single, penalty_wrong)` — single/multi choice
  - `create_shortanswer(subject)` — shortanswer + permutations + partial scoring
  - `create_shortanswer_numerical()` — числовой ответ (. и , варианты)
  - `create_matching()` — matching с дистракторами
  - `create_ddmatch()` — drag-and-drop
  - `create_gapselect()` — выпадающие списки
  - `create_cloze()` — встроенные ответы
  - `create_numerical()` — numerical (fallback)
- **`MoodleConverter`** — парсер docx + оркестратор

### Алгоритм partial scoring ({shortanswer_partial})

Для вопросов с несколькими правильными/неправильными ответами:
1. Текст вопроса нумеруется: `1)`, `2)`, `3)`... вместо `+:`/`-:`
2. Генерируются **ВСЕ перестановки** (1, 2, 3, ..., 12, 13, ..., 654321):
   - permutations('123456', 1) → 1,2,3,4,5,6
   - permutations('123456', 2) → 12,13,14,...,21,23,24...
   - permutations('123456', 6) → 654321
3. Fraction:
   - **100%**: все правильные цифры, ни одной неправильной
   - **50%**: ≥50% правильных и не более 1 неправильной ИЛИ все правильные + 1 неправильная
   - **0%**: остальные

### Алгоритм numcombo ({shortanswer_numcombo})

Для вопросов с несколькими правильными/неправильными ответами:
1. Текст вопроса нумеруется: `1)`, `2)`, `3)`... вместо `+:`/`-:`
2. Генерируются **ВСЕ перестановки** позиций правильных ответов:
   - 1 правильный → номер позиции (например "3")
   - несколько правильных → все перестановки (например "356", "365", "536"...)
3. Все ответы = 100%

### Алгоритм перестановок для текстовых ответов

Если ответ — строка цифр в shortanswer_phrase:
- Ограничение: максимум 7 цифр (7! = 5040 перестановок)
- 8+ цифр: только один ответ (8! = 40320 — слишком много)
- Фразы "в порядке возрастания/убывания" блокируют перестановки

---

## Логи конвертации

### Результат обработки (2026-04-08)

```
Файл                              Вопросов  Маркеры
вопросы-АЯ  10кл  ВИ ЛФУ.xml       615     multichoice_one, shortanswer_phrase, matching, gapselect
вопросы-АЯ  8кл  ВИ ЛФУ.xml        456     multichoice_one, matching
вопросы-ИСТ 10кл  ВИ ЛФУ.xml       131     match, multichoice_one, shortanswer_numcombo, match_123
вопросы-ИЯ  10кл  ВИ ЛФУ.xml        95     multichoice_one, matching, gapselect
вопросы-МАТ  10кл  ВИ ЛФУ.xml      422     numerical, multichoice_one
вопросы-МАТ  8кл  ВИ ЛФУ.xml       200     numerical, multichoice_one
вопросы-НЯ  10кл  ВИ ЛФУ.xml        95     multichoice_one, matching, gapselect
вопросы-ОБЩ 10кл  ВИ ЛФУ.xml       375     multichoice_many, shortanswer_partial, match
вопросы-РЯ  10кл  ВИ ЛФУ.xml        510     multichoice_many, shortanswer_phrase, ddmatch, shortanswer_numcombo
вопросы-ФИЗ  10кл  ВИ ЛФУ.xml       230     multichoice_one, numerical
вопросы-ФЯ  10кл  ВИ ЛФУ  2026.xml   95     multichoice_one, shortanswer_phrase, matching, gapselect
                                    -----
Итого:                             3224     Ошибок: 0
```
Файл                              Вопросов  Маркеры
вопросы-АЯ  10кл  ВИ ЛФУ.xml       615     multichoice_one, shortanswer_phrase, matching, gapselect
вопросы-АЯ  8кл  ВИ ЛФУ.xml        456     multichoice_one, matching
вопросы-ИСТ 10кл  ВИ ЛФУ.xml       131     match, multichoice_one, shortanswer_numcombo, match_123
вопросы-ИЯ  10кл  ВИ ЛФУ.xml        95     multichoice_one, matching, gapselect
вопросы-МАТ  10кл  ВИ ЛФУ.xml      422     numerical, multichoice_one
вопросы-МАТ  8кл  ВИ ЛФУ.xml       200     numerical, multichoice_one
вопросы-НЯ  10кл  ВИ ЛФУ.xml        95     multichoice_one, matching, gapselect
вопросы-ОБЩ 10кл  ВИ ЛФУ.xml       375     multichoice_many, shortanswer_partial, match
вопросы-РЯ  10кл  ВИ ЛФУ.xml       510     multichoice_many, shortanswer_phrase, ddmatch, shortanswer_numcombo
вопросы-ФИЗ  10кл  ВИ ЛФУ.xml      230     multichoice_one, numerical
вопросы-ФЯ  10кл  ВИ ЛФУ  2026.xml  95     multichoice_one, shortanswer_phrase, matching, gapselect
                                   -----
Итого:                             3224     Ошибок: 0
```

### Исправленные баги (из v3 baseline)

| # | Баг | Исправление |
|---|-----|-------------|
| 1 | `<answer>` снаружи `<subquestion>` в matching | Перемещён внутрь `<subquestion>` |
| 2 | Нет дистракторов в matching | Добавлены пустые `<subquestion>` для лишних R-элементов |
| 3 | gapselect не распознаётся для НЯ/ИЯ/ФЯ | Добавлен формат `+: ABCD`, исправлен regex для `1.A)` |
| 4 | Дублирование перестановок shortanswer | Убрано бессмысленное дублирование |
| 5 | `parse_answers_from_line` crash (3 vs 2 группы) | Исправлена распаковка кортежа |
| 6 | 8! = 40320 перестановок для 8-значных ответов | Лимит снижен до 7 цифр (7! = 5040) |
| 7 | "в порядке возрастания" генерировало перестановки | Добавлена проверка, блокирующая перестановки |
| 8 | `S:`, `I:`, `V1:`, `V2:`, `+:`, `-:` в XML questiontext | Добавлена функция `remove_service_markers()` при выводе |
| 9 | Переходы на новую строку не сохранялись | Использование `<br>` между частями вопроса |
| 10 | Подкатегории V2 работали только в начале файла | Добавлена обработка V1/V2 в любом месте файла |
| 11 | Дублирование категорий/подкатегорий при разделении XML | Добавлена защита от дубликатов + начало файла с последней категорией |
| 12 | {shortanswer_numcombo} не работал | Перемещён перед partial, добавлена нумерация + комбинации |
| 13 | {shortanswer_partial} не генерировал все варианты | Заменён combinations на permutations (как в оригинале) |
