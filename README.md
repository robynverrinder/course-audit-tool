# Course Folder Audit Tool

A desktop GUI application for auditing standardised EEE course folders at UCT.  
The tool compares a selected course folder against an expected structure profile,
flags issues, and exports a plain-text log and a colour-coded Excel workbook for
coordinated review.

---

## Features

- **Three structure profiles** — Current, Legacy, and Updated (auto-detected or manually selected)
- **NONE convention** — folders marked with `NONE` are treated as intentionally empty
- **Admin marker detection** — folders renamed by administrators with status words (e.g. `MISSING`, `INCOMPLETE`, `UNSIGNED`) are detected, matched to their template entry, and flagged `ADMIN FLAG` for review
- **Issue detection** — flags Missing, Empty, Unexpected, Duplicate, and "Populated Despite NONE" folders
- **Sample hand-ins validation** — checks that each submission group has ≥ 15 PDF or Word documents; other file types (e.g. `.yml`, `.zip`) are ignored and do not cause a flag
- **Per-group issue reporting** — when a submission group fails, the specific folder name is named in the detail message
- **Flat-file tolerance** — folders that contain files directly (without subfolders) are accepted as OK rather than flagging expected subfolders as missing
- **Six tabbed views** — Log Output, Issues, Expected Structure Check, Folder Details, File Details, ASCII Tree
- **Excel workbook output** — colour-coded status cells and blank reviewer columns for coordinated review
- **Plain-text log** — full audit record saved alongside the workbook
- **Recent directories** — last 10 used paths persisted between sessions

---

## Requirements

- Python 3.10 or later
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel workbook generation
- [Pillow](https://pillow.readthedocs.io/) — PNG logo rendering in the title bar
- tkinter — included with standard Python on Windows and macOS

Install dependencies:

```bash
pip install -r requirements.txt
```

> **macOS note:** If `tkinter` is missing, install it via Homebrew:
> ```bash
> brew install python-tk
> ```

> **Logo note:** Place `logo_uct.png` and `logo_eee.png` in the same folder as
> the script. If either file is missing, or if Pillow is not installed, the
> application falls back to a plain text label — no functionality is affected.
> The UCT logo is used as the window/taskbar icon.

---

## Installation

```bash
git clone https://github.com/robynverrinder/course-folder-audit.git
cd course-folder-audit
pip install -r requirements.txt
```

---

## Usage

```bash
python3 course_audit_tool.py
```

> **Windows users:** use `python course_audit_tool.py` or `python3 course_audit_tool.py` depending on your installation. Run `python --version` and `python3 --version` to check which is available.

1. **Select a course root** using the Browse button or pick from the recent-directories dropdown.
2. **Choose a profile** — leave on *Auto-detect* or manually select Current, Legacy, or Updated.
3. **Select an auditor** from the dropdown.
4. Click **Run Audit and Create Outputs**.

Two files are saved into the course root directory:

| File | Description |
|------|-------------|
| `YYYYMMDD_<course>_folder_audit.txt` | Plain-text audit log |
| `YYYYMMDD_<course>_folder_audit.xlsx` | Excel workbook with colour-coded status |

---

## Folder Structure Profiles

### Current
The standard 13-folder layout used from approximately 2022 onwards.  
Exam content is split across `12. Exams Main (Admin)` and `13. Exams SUPP (Admin)`.

### Legacy
An earlier 13-folder layout. Key differences:
- Lessons include a `Recordings` subfolder
- Many subfolders carry an `(h)` suffix (hidden/restricted content)
- Exam content lives in `09. Exam` and `13. Supplementary Exam`
- `11. Additional resources` has explicit subfolders

### Updated
A simplified 7-folder layout introduced with the new template. Key differences:
- `02. Teaching material` holds notes and prescribed textbooks
- `04. Practicals & Labs` replaces the standalone Practicals folder
- Subfolders use `Handouts` instead of `Instruction sheets`
- A single `07. Exams` folder contains `Main Exams` and `SUPP Exams` sub-groups, each holding the same five subfolders

### Auto-detection
When set to *Auto-detect*, the tool scores the top-level folder names on disk against unique markers for each profile and selects the best match. The folder count is also used as a tiebreaker — Updated courses have 7 top-level folders, while Current and Legacy have 13. If no markers match, Current is used as the safe default.

---

## NONE Convention

Any folder whose name contains the word `NONE` (in any position, with any
surrounding separator) is treated as intentionally empty for this offering.

| State | Status |
|-------|--------|
| NONE-marked folder, empty | `NONE - ACCEPTED` |
| NONE-marked folder, has files | `POPULATED DESPITE NONE` |

Examples that are recognised: `a. Slides NONE`, `a. Slides - NONE`, `NONE slides`

---

## Admin Marker Convention

Administrators sometimes rename folders to flag issues manually, appending a status word to the folder name rather than moving or deleting it. The tool detects these annotations, still matches the folder to its correct template entry, and flags it as `ADMIN FLAG` (amber) so it appears in the Issues tab for review.

Recognised marker words and phrases (trailing position only):

| Marker | Example folder name |
|--------|-------------------|
| `MISSING` | `f. Mark sheets MISSING` |
| `INCOMPLETE` | `05. Practicals INCOMPLETE` |
| `EMPTY` | `a. Exam paper EMPTY` |
| `UNSIGNED` | `c. External moderator reports UNSIGNED` |
| `TO BE SIGNED` | `f. Mark sheets TO BE SIGNED` |
| `URGENT` | `06. Tests URGENT` |
| `TODO` | `03. Lessons TODO` |

Compound qualifiers immediately before the marker are also stripped, for example `f. Mark sheets COR MISSING` and `c. External moderator reports SIGNATURES MISSING` both match their template entries correctly.

Words like `REVIEW`, `CHECK`, `ACTION`, and `PENDING` are intentionally excluded as they appear too commonly in legitimate folder names (e.g. `b. Review materials`).

---

Certain folders are checked not just for presence but for the correct file types.
If files exist but none match the required type, the folder is flagged `EMPTY - REVIEW`
with a detail message listing what was found.

| Folder | Required file types |
|--------|-------------------|
| Course Handout / Course Handouts | At least 1 × `.pdf`, `.doc`, or `.docx` |
| DP list / DP list final | At least 1 × `.pdf`, `.xls`, `.xlsx`, or `.csv` |
| Mark Sheets / Marks Sheets *(under Exams)* | At least 1 × `.pdf`, `.xls`, `.xlsx`, or `.csv` |
| External Moderator Reports *(under Exams)* | At least 1 × `.pdf`, `.doc`, or `.docx` |

### Submission folders (Sample hand-ins, Sample answers, Exam scripts)

Each submission group (e.g. `Practical 1 of 5`, `Tutorial 3`) is validated independently and must contain at least 15 PDF or Word documents (`.pdf`, `.doc`, `.docx`). Other file types present in the same folder are ignored and do not trigger a flag.

If a group fails, the detail message names it explicitly, for example:

> `2 groups with issues: "Practical 3 of 5": 8 PDF/Word docs (expected >=15) | "Practical 5 of 5": 12 PDF/Word docs (expected >=15, 1 other file type(s) ignored)`

Any loose files sitting directly alongside submission subfolders (e.g. a marking note PDF) are ignored — only the leaf subfolders are counted as groups.

---

## Status Reference

| Status | Colour | Meaning |
|--------|--------|---------|
| `OK` | Green | All checks passed |
| `EMPTY - REVIEW` | Yellow | Folder exists but contains no files |
| `MISSING` | Red | Expected folder is absent |
| `MISSING CHILDREN` | Red | Parent present but a child is missing |
| `UNEXPECTED` | Orange | Folder exists but is not in the template |
| `NONE - ACCEPTED` | Blue | Intentionally empty (NONE-marked) |
| `POPULATED DESPITE NONE` | Purple | NONE-marked folder contains files |
| `REVIEW - HAND-INS` | Orange | Submission folder has fewer than 15 PDF/Word docs in one or more groups |
| `DUPLICATE` | Red | Two folders resolve to the same normalised name |
| `ADMIN FLAG` | Amber | Folder name contains an administrator status marker — review and rename |

---

## Building a Standalone Executable

### macOS

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --icon=logo_uct.png \
  --add-data "logo_uct.png:." \
  --add-data "logo_eee.png:." \
  course_audit_tool.py
```

The executable is created at `dist/course_audit_tool`. Rename it as needed.  
If macOS blocks it with *"unidentified developer"*, right-click → Open the first time.

### Windows

Run the following on a Windows machine (PyInstaller builds for the OS it runs on):

```bat
pip install pyinstaller openpyxl pillow
pyinstaller --onefile --windowed --icon=logo_uct.ico ^
  --add-data "logo_uct.png;." ^
  --add-data "logo_eee.png;." ^
  course_audit_tool.py
```

The executable is created at `dist\course_audit_tool.exe`.  
Note: `--icon` requires a `.ico` file on Windows. Convert your PNG first:

```bash
python -c "from PIL import Image; img = Image.open('logo_uct.png'); img.save('logo_uct.ico', sizes=[(16,16),(32,32),(48,48),(256,256)])"
```

If Windows Defender blocks the exe, click *More info → Run anyway* in the SmartScreen dialog.

---

## Adding Auditors

Open `course_audit_tool.py` and add names to the `AUDIT_USERS` list near the top of the file:

```python
AUDIT_USERS: list[str] = [
    "Robyn Verrinder",
    "Yunus Abdul Gaffar",
    # Add new auditors here
]
```

No other code changes are required.

---

## Project Structure

```
course-folder-audit/
├── course_audit_tool.py  # Main application
├── logo_uct.png          # UCT logo (title bar right, window icon)
├── logo_eee.png          # EEE logo (title bar left)
├── requirements.txt      # Python dependencies
├── README.md             # This file
└── LICENSE               # MIT licence
```

---

## License

This project is licensed under the MIT License — see [LICENSE](LICENSE) for details.
