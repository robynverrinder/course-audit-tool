# Course Folder Audit Tool

A desktop GUI application for auditing EEE course folder structures at UCT against the departmental archive standard. Produces a plain-text log and an Excel review workbook saved directly into the scanned course folder.

---

## Requirements

- Python 3.10 or later (tested on 3.13.9 with Anaconda `base`)
- `openpyxl` — Excel workbook generation
- `Pillow` — optional, for logo display in the title bar

Install dependencies:

```bash
pip install openpyxl Pillow
```

Or with conda:

```bash
conda install openpyxl Pillow
```

---

## Running the tool

```bash
python3 course_audit_tool.py
```

From VS Code, set the interpreter to your Anaconda `base` environment and run directly. On macOS, `conda activate base` first if your terminal does not activate it automatically.

---

## What it does

Point the tool at a course-year folder (e.g. `2025 EEE2046F EEE2050F Abdul-Gaffar`) and click **Run Audit and Create Outputs**. The tool:

1. Auto-detects the folder structure profile (Legacy, Current, New, or EEE4022)
2. Compares every folder against the expected template for that profile
3. Checks that key folders contain the right file types (course handout, DP list, mark sheets, external moderator reports, sample hand-ins)
4. For GA courses, validates the `00_ga_moderation` folder and checks assessment forms completeness against the reference student count
5. Flags unsigned files, admin status markers, duplicates, NONE-marked folders with content, and unexpected folders
6. Saves a `.txt` log and `.xlsx` workbook into the scanned course root

---

## Folder structure profiles

| Profile | Years | Key markers |
|---|---|---|
| Legacy | 2023, 2024 | `09. Exam`, `13. Supplementary Exam`, `(h)` suffixes |
| Current | 2025 | `12. Exams Main (Admin)`, `13. Exams SUPPS (Admin)`, no `(h)` |
| New | 2026 | `lower_snake_case` folder names, `08_exams` two-level structure |
| EEE4022 | any | Course code `EEE4022` in folder name — project-based capstone course |

Profile is auto-detected from the year in the folder name. It can also be set manually from the dropdown.

---

## GA courses

The following course codes are treated as GA courses and require a `00_ga_moderation` folder:

```
EEE3088F  EEE3096S  EEE3097S  EEE3098S  EEE3099S  EEE3100S
EEE4022F  EEE4022S  EEE4113F  EEE4118F  EEE4119F  EEE4120F
EEE4121F  EEE4124C  EEE4125C  EEE4126F
```

For GA assessment forms completeness, the tool reads the reference student count from the mark sheets folder under the main exam folder. Drop a UCT PeopleSoft class list export (with an `Emplid` or `Campus ID` column) into `f. Mark sheets` / `d_marksheets` and it will be picked up automatically.

---

## Status flags

| Status | Meaning |
|---|---|
| `OK` | Folder found and passes all checks |
| `EMPTY - REVIEW` | Folder exists but contains no files |
| `MISSING` | Expected folder is absent |
| `UNEXPECTED` | Folder exists but is not in the template |
| `NONE - ACCEPTED` | Folder is marked NONE and is empty — accepted |
| `POPULATED DESPITE NONE` | Folder is marked NONE but contains files |
| `REVIEW - HAND-INS` | Sample hand-ins folder has fewer than 15 submissions in one or more groups |
| `REVIEW - GA INCOMPLETE` | GA assessment forms folder has fewer files than the reference student count |
| `DUPLICATE` | Two or more folders resolve to the same normalised name |
| `ADMIN FLAG` | Folder name contains an admin status marker (MISSING, UNSIGNED, etc.) |
| `UNSIGNED FILE` | A file name ends with `unsigned` |
| `MISSING CHILDREN` | Top-level folder found but expected subfolders are missing |

---

## Reference student count

The tool looks for the reference student count in the mark sheets folder under the main exam folder:

| Profile | Path |
|---|---|
| Legacy | `09. Exam/f. Mark sheets` |
| Current | `12. Exams Main (Admin)/f. Mark sheets` |
| New | `08_exams/01_main/d_marksheets` |
| EEE4022 | `3. Final Results` |

The first spreadsheet in that folder with an `Emplid` or `Campus ID` column is used. Drop a UCT PeopleSoft export there to enable student count checks.

---

## Output files

Both output files are saved into the scanned course root, named `YYYYMMDD_coursecode_folder_audit.txt` and `.xlsx`.

The workbook has five sheets:

- **Course Audit Summary** — scan metadata and headline counts
- **Expected Structure Check** — full comparison of expected vs actual folders with status and reviewer columns
- **Folder Details** — file counts, sizes, and types per folder
- **File Details** — every file with type, size, and modification date
- **Exceptions** — issues only, for review and sign-off

---

## NONE convention

Any folder whose name contains the word `NONE` (in any position, any separator) is treated as intentionally empty:

- Empty NONE folder → `NONE - ACCEPTED`
- NONE folder with content → `POPULATED DESPITE NONE`

Example: `b. Prescribed texts - NONE` is accepted as empty. If files are later added without removing `NONE` from the name, it is flagged.

---

## Admin status markers

Folders with a recognised admin status marker appended to their name are flagged as `ADMIN FLAG` rather than `MISSING` or `EMPTY`. The marker is stripped during normalisation so the folder still matches the expected template entry.

| Marker | Example |
|---|---|
| `MISSING` | `f. Mark sheets MISSING` |
| `INCOMPLETE` | `c. External moderator reports INCOMPLETE` |
| `URGENT` | `d. DP list final URGENT` |
| `TODO` | `a. Course handouts TODO` |
| `UNSIGNED` | `e. Departmental control sheet UNSIGNED` |
| `EMPTY` | `b. Prescribed texts EMPTY` |
| `TO BE SIGNED` | `f. Mark sheets TO BE SIGNED` |

The marker can be separated from the folder name by a space, hyphen, or underscore. Case is ignored.

---

## Adding course codes

To add a GA course code, edit the `GA_COURSES` set near the top of `course_audit_tool.py`:

```python
GA_COURSES: set[str] = {
    "eee4022s", "eee4022f",
    ...
}
```

Codes are lowercase and year-independent.

To add a TA name to the dropdown, edit the `TA_NAMES` list:

```python
TA_NAMES: list[str] = [
    "",
    "Jones",
    ...
]
```

---

## Repository

`robynverrinder/course-audit-tool`

Maintained by R.A. Verrinder, Department of Electrical and Electronic Engineering, UCT.
