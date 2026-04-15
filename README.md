# Moodle XML Converter v3

Word (.docx) to Moodle XML converter with graphical interface.

## Files

| File | Description |
|------|-------------|
| `converter_gui.py` | GUI application (PyQt5) |
| `universal_moodle_converter_v3_stable.py` | Converter core (CLI + library) |
| `table_compare.py` | Utility for comparing with reference XML |
| `Шаблоны вопросов_new.docx` | Documentation on Word file markup |
| `taskmanager/task-tracker.html` | Task history |

## Dependencies

```
pip install PyQt5 lxml python-docx docxlatex
```

## Running

### GUI
```
python converter_gui.py
```
Or double-click `Запустить конвертер.bat`

### CLI (batch processing)
```
python universal_moodle_converter_v3_stable.py <path to docx or folder> --output-folder <folder>
```

### Compare with reference
```
python table_compare.py
```

---

## Architecture

### Word File Structure

```
V1: Subject name (category)
{marker}V2: Block name (subcategory)

I:Task N. Author I.O., TZ X-Y, b=N
S: Question text
+: Correct answer
-: Incorrect answer
```

### Question Type Markers

Marker is placed before `V2:` in format `{marker}V2: Block description`.
All questions in the block inherit the marker until the next `V2:`.

| Marker | Moodle Type | Description |
|--------|-------------|-------------|
| `{multichoice_one}` | multichoice (single=true) | One correct answer. `+:` = 100%, `-:` = 0% |
| `{multichoice_many}` | multichoice (single=false) | Multiple correct answers. Penalty **-100%** for each incorrect |
| `{shortanswer_phrase}` | shortanswer | Text input. Multiple `+:` = multiple acceptable answers |
| `{numerical_partial}` | **numerical** | Multiple choice (numbered 1)2)3)...). All permutations with partial scoring: 100%/50%/0% |
| `{numerical_numcombo}` | **numerical** | Multiple choice. All position permutations = 100% |
| `{matching}` / `{match}` | matching | Matching. Format `L1:` / `R1:`. **All** L and R entries are used. Duplicate R = distractors |
| `{match_123}` | matching | Sequence. Format `N: phrase` -> phrase matched to number |
| `{ddmatch}` | ddmatch | Drag-and-drop. Format `L1:` / `R1:` |
| `{gapselect}` | gapselect | Dropdown lists. Text with `(N)`, options `A)...D)`, key `+:ABCD` |
| `{cloze}` | cloze | Embedded answers `{1:SHORTANSWER:=answer}` |
| `{numerical}` | shortanswer | Numeric answer. Generates two variants: with `.` and with `,` |

**Notes:** 
- `{numerical_partial}` and `{numerical_numcombo}` generate `type="numerical"` in XML output (not "shortanswer") 
- `{matching}` now uses **all** L and R entries (not just unique). Duplicate R entries become distractors

If marker is not specified, type is determined by heuristic based on content.

### Question Header Formats (7 variants)

The converter recognizes 7 formats of question beginning:
1. `I:Task N.` — standard
2. `I I:Task N.` — double I (Word artifact)
3. `I Task N.` — space instead of colon
4. `:Task N.` — missing I symbol
5. `Task N. Author, TZ X-Y, b=N` — without I: prefix
6. `Kn-=mTask N.` — garbage before Task
7. `Author I.O., TZ X-Y, b=N` — only author (without Task word)

---

## GUI: converter_gui.py

### Features

1. **File selection** — "Browse" button for .docx
2. **Output folder selection** — where to save XML
3. **Question selection** — checkboxes to select which questions to convert
   - "Выделить все" checkbox to select/deselect all
   - Counter shows "Выбрано: X / Y"
4. **Preview** (QTreeWidget):
   - List of all questions with expandable content
   - Clicking a question reveals: text (S:), correct (+:, green), incorrect (-:, red), L/R pairs
   - Marker combobox — can change marker for block
   - Color coding by marker type
   - Error highlighting in red
5. **Preprocessing errors**:
   - Missing correct answer
   - Empty question text
   - Unknown marker
6. **Conversion** in separate thread with progress bar
7. **XML post-processing**:
   - Root element check (`quiz`)
   - Question type check (only valid Moodle types)
   - Base64 images check (not empty)
   - Check for `_IMAGE_` / `@@PLUGINFILE@@` markers without files
   - Matching structure check (subquestion/answer)
   - Gapselect check (selectoption)
   - Answer presence check
8. **XML splitting** into parts up to 1 MB (checkbox)

### Marker Color Scheme

| Color | Markers |
|-------|---------|
| Blue | multichoice_one, multichoice_many |
| Green | numerical_partial, numerical_numcombo, shortanswer_phrase |
| Orange | matching, match_123, match |
| Reddish | ddmatch |
| Violet | gapselect |
| Yellow | cloze |
| Turquoise | numerical |

---

## Core: universal_moodle_converter_v3_stable.py

### Classes

- **`ImageProcessor`** — extract images from docx (base64)
- **`FormulaProcessor`** — convert LaTeX formulas (`$...$` -> `\(...\)`)
- **`QuestionTypeDetector`** — determine question type (marker has priority over heuristic)
- **`XMLGenerator`** — generate Moodle XML:
  - `create_multichoice(single, penalty_wrong)` — single/multi choice
  - `create_shortanswer(subject)` — shortanswer + permutations + partial scoring
  - `create_shortanswer_numerical()` — numeric answer (. and , variants)
  - `create_matching()` — matching with distractors
  - `create_ddmatch()` — drag-and-drop
  - `create_gapselect()` — dropdown lists
  - `create_cloze()` — embedded answers
  - `create_numerical()` — numerical (fallback)
- **`MoodleConverter`** — docx parser + orchestrator

### Partial Scoring Algorithm ({numerical_partial})

For questions with multiple correct/incorrect answers:
1. Question text is numbered: `1)`, `2)`, `3)`... instead of `+:`/`-:`
2. **ALL permutations** are generated (1, 2, 3, ..., 12, 13, ..., 654321):
   - permutations('123456', 1) → 1,2,3,4,5,6
   - permutations('123456', 2) → 12,13,14,...,21,23,24...
   - permutations('123456', 6) → 654321
3. Fraction:
   - **100%**: all correct digits, no incorrect
   - **50%**: ≥50% correct and no more than 1 incorrect OR all correct + 1 incorrect
   - **0%**: NOT included (only 100% and 50% answers are generated)

### Numcombo Algorithm ({numerical_numcombo})

For questions with multiple correct/incorrect answers:
1. Question text is numbered: `1)`, `2)`, `3)`... instead of `+:`/`-:`
2. **ALL permutations** of correct answer positions are generated:
   - 1 correct → position number (e.g., "3")
   - multiple correct → all permutations (e.g., "356", "365", "536"...)
3. All answers = 100%

### Permutation Algorithm for Text Answers

If answer is a string of digits in shortanswer_phrase:
- Limit: maximum 7 digits (7! = 5040 permutations)
- 8+ digits: only one answer (8! = 40320 — too many)
- Phrases "in ascending order"/"in descending order" block permutations

---

## Conversion Logs

### Processing Result (2026-04-09)

```
File                              Questions  Markers
questions-ИСТ 10кл               131        match, multichoice_one, numerical_partial, match_123
                                    -----
Total:                            131        Errors: 0
```

### Testing Results (2026-04-08)

```
File                              Questions  Markers
questions-AJ  10cl                615        multichoice_one, shortanswer_phrase, matching, gapselect
questions-AJ  8cl                 456        multichoice_one, matching
questions-HIST 10cl               131        match, multichoice_one, numerical_numcombo, match_123
questions-FL  10cl                 95        multichoice_one, matching, gapselect
questions-MATH 10cl              422        numerical, multichoice_one
questions-MATH 8cl               200        numerical, multichoice_one
questions-GER  10cl               95        multichoice_one, matching, gapselect
questions-SOC  10cl              375        multichoice_many, numerical_partial, match
questions-RU  10cl                510        multichoice_many, shortanswer_phrase, ddmatch, numerical_numcombo
questions-PHYS 10cl              230        multichoice_one, numerical
questions-LING 10cl               95        multichoice_one, shortanswer_phrase, matching, gapselect
                                    -----
Total:                            3224       Errors: 0
```

---

## Bug Fixes

| # | Bug | Fix |
|---|-----|-----|
| 1 | `<answer>` outside `<subquestion>` in matching | Moved inside `<subquestion>` |
| 2 | No distractors in matching | Added empty `<subquestion>` for extra R elements |
| 3 | gapselect not recognized for GER/FL/LING | Added format `+: ABCD`, fixed regex for `1.A)` |
| 4 | shortanswer permutations duplication | Removed meaningless duplication |
| 5 | `parse_answers_from_line` crash (3 vs 2 groups) | Fixed tuple unpacking |
| 6 | 8! = 40320 permutations for 8-digit answers | Limit reduced to 7 digits (7! = 5040) |
| 7 | "in ascending order" generated permutations | Added check blocking permutations |
| 8 | `S:`, `I:`, `V1:`, `V2:`, `+:`, `-:` in XML questiontext | Added `remove_service_markers()` function on output |
| 9 | Line breaks not preserved | Used `<br>` between question parts |
| 10 | V2 subcategories only worked at file start | Added V1/V2 processing anywhere in file |
| 11 | Category/subcategory duplication when splitting XML | Added duplicate protection + file starts with last category |
| 12 | {numerical_numcombo} not working | Moved before partial, added numbering + combinations |
| 13 | {numerical_partial} not generating all variants | Replaced combinations with permutations |
| 14 | GUI "Select all" checkbox too large | Fixed to single row, fixed height 32px |
| 15 | Missing feedback elements in NUMERICAL_TEMPLATE | Added correctfeedback, partiallycorrectfeedback, incorrectfeedback |
| 16 | XML splitting - duplicate XML declaration | Fixed: each part has single `<?xml>` and proper `</quiz>` closing |
| 17 | {numerical_numcombo} - answers not added to XML | Added missing `self.root.append(tree)` after return |
| 18 | {numerical_numcombo} - wrong answer logic | Rewrote: takes answer from `+:` and generates all permutations of digits |
| 19 | {numerical_numcombo} - question text split with `<br>` | Preserved original line breaks, keeps (1), (2), (3) intact |
| 20 | {numerical_partial} - duplicate text answers | Fixed: removes duplicate else branch, only numerical answers |
| 21 | {matching} - not all L and R entries used | Fixed: uses all entries, duplicate R = distractors |