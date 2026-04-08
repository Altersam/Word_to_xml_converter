# Moodle XML Converter v3

Word (.docx) to Moodle XML converter with graphical interface.

## Files

| File | Description |
|------|-------------|
| `converter_gui.py` | GUI application (PyQt5) |
| `universal_moodle_converter_v3_stable.py` | Converter core (CLI + library) |
| `table_compare.py` | Utility for comparing with reference XML |
| `Шаблоны вопросов.docx` | Documentation on Word file markup |

## Dependencies

```
pip install PyQt5 lxml python-docx docxlatex
```

## Running

### GUI
```
python converter_gui.py
```

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
| `{shortanswer_partial}` | shortanswer | Multiple choice (numbered 1)2)3)...). All permutations with partial scoring: 100%/50%/0% |
| `{shortanswer_numcombo}` | shortanswer | Multiple choice. All position permutations = 100% |
| `{matching}` / `{match}` | matching | Matching. Format `L1:` / `R1:`. Extra R = distractors |
| `{match_123}` | matching | Sequence. Format `N: phrase` -> phrase matched to number |
| `{ddmatch}` | ddmatch | Drag-and-drop. Format `L1:` / `R1:` |
| `{gapselect}` | gapselect | Dropdown lists. Text with `(N)`, options `A)...D)`, key `+:ABCD` |
| `{cloze}` | cloze | Embedded answers `{1:SHORTANSWER:=answer}` |
| `{numerical}` | shortanswer | Numeric answer. Generates two variants: with `.` and with `,` |

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
3. **Preview** (QTreeWidget):
   - List of all questions with expandable content
   - Clicking a question reveals: text (S:), correct (+:, green), incorrect (-:, red), L/R pairs
   - Marker combobox — can change marker for block
   - Color coding by marker type
   - Error highlighting in red
4. **Preprocessing errors**:
   - Missing correct answer
   - Empty question text
   - Unknown marker
5. **Conversion** in separate thread with progress bar
6. **XML post-processing**:
   - Root element check (`quiz`)
   - Question type check (only valid Moodle types)
   - Base64 images check (not empty)
   - Check for `_IMAGE_` / `@@PLUGINFILE@@` markers without files
   - Matching structure check (subquestion/answer)
   - Gapselect check (selectoption)
   - Answer presence check
7. **XML splitting** into parts up to 1 MB (checkbox)

### Marker Color Scheme

| Color | Markers |
|-------|---------|
| Blue | multichoice_one, multichoice_many |
| Green | shortanswer_phrase, shortanswer_partial, shortanswer_numcombo |
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

### Partial Scoring Algorithm ({shortanswer_partial})

For questions with multiple correct/incorrect answers:
1. Question text is numbered: `1)`, `2)`, `3)`... instead of `+:`/`-:`
2. **ALL permutations** are generated (1, 2, 3, ..., 12, 13, ..., 654321):
   - permutations('123456', 1) → 1,2,3,4,5,6
   - permutations('123456', 2) → 12,13,14,...,21,23,24...
   - permutations('123456', 6) → 654321
3. Fraction:
   - **100%**: all correct digits, no incorrect
   - **50%**: ≥50% correct and no more than 1 incorrect OR all correct + 1 incorrect
   - **0%**: rest

### Numcombo Algorithm ({shortanswer_numcombo})

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

### Processing Result (2026-04-08)

```
File                              Questions  Markers
questions-AJ  10cl                615        multichoice_one, shortanswer_phrase, matching, gapselect
questions-AJ  8cl                 456        multichoice_one, matching
questions-HIST 10cl               131        match, multichoice_one, shortanswer_numcombo, match_123
questions-FL  10cl                 95        multichoice_one, matching, gapselect
questions-MATH 10cl              422        numerical, multichoice_one
questions-MATH 8cl               200        numerical, multichoice_one
questions-GER  10cl               95        multichoice_one, matching, gapselect
questions-SOC  10cl              375        multichoice_many, shortanswer_partial, match
questions-RU  10cl                510        multichoice_many, shortanswer_phrase, ddmatch, shortanswer_numcombo
questions-PHYS 10cl              230        multichoice_one, numerical
questions-LING 10cl               95        multichoice_one, shortanswer_phrase, matching, gapselect
                                    -----
Total:                            3224       Errors: 0
```

### Bug Fixes (from v3 baseline)

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
| 12 | {shortanswer_numcombo} not working | Moved before partial, added numbering + combinations |
| 13 | {shortanswer_partial} not generating all variants | Replaced combinations with permutations (as in original) |