"""
Course Audit Tool
========================================
GUI application for auditing standardised EEE course folders at UCT.

Produces two output artefacts saved into the scanned course root:
  - A plain-text audit log  (<date>_<course>_folder_audit.txt)
  - An Excel review workbook (<date>_<course>_folder_audit.xlsx)

Supported folder-structure profiles
-------------------------------------
  Current  — 13-folder layout used from approximately 2022 onwards.
  Legacy   — Earlier 13-folder layout; includes Recordings subfolders and
             folders suffixed with "(h)" to denote hidden/restricted content.
  Updated  — Simplified 7-folder layout introduced for the new template,
             with a single "07. Exams" folder containing "Main exams" and
             "SUPP exams" sub-groups rather than two separate top-level
             exam folders.

Auto-detection
---------------
  Profile is detected by scoring top-level folder names against unique markers
  for each profile.  Folder count is used as a tiebreaker — Updated courses
  have 7 top-level folders; Current and Legacy have 13.

NONE convention
----------------
  Any folder whose name contains the word "NONE" (in any position, with any
  surrounding separator) is treated as intentionally empty.  The auditor has
  declared that no content is expected here for this offering.
    - Empty NONE folder      → "NONE - ACCEPTED"        (expected state)
    - NONE folder with files → "POPULATED DESPITE NONE" (flagged for review)

Admin marker convention
------------------------
  Administrators sometimes append status words to folder names to flag issues
  manually, e.g. "01. Administration - MISSING" or "05. Practicals INCOMPLETE".
  Recognised markers: MISSING, INCOMPLETE, EMPTY, REVIEW, PENDING, TODO,
  CHECK, URGENT, ACTION.
    - Folder is still matched against the template (marker is stripped first)
    - Folder is flagged "ADMIN FLAG" (amber) and appears in the Issues tab

Flat-file tolerance
--------------------
  If a lecturer places files directly inside a top-level folder (without
  creating the expected subfolders), the folder is marked OK rather than
  flagging all expected subfolders as MISSING.

Submission validation
----------------------
  Sample hand-ins, sample answers, and exam scripts folders are validated
  independently per submission group (e.g. each practical sub-folder).
  Each group must contain at least 15 PDF or Word documents (.pdf, .doc,
  .docx).  Other file types are ignored.  Loose files alongside submission
  sub-folders (e.g. a marking note) are also ignored.

Platform notes
---------------
  macOS .DS_Store files are silently excluded from all audit logic and counts.
  The root directory itself is excluded from folder/file listing rows; only
  its children and descendants are reported.
"""

# ---------------------------------------------------------------------------
# Standard-library imports
# ---------------------------------------------------------------------------
import json          # Persist recent-directories list between sessions
import os            # File-system traversal and path manipulation
import re            # Regular expressions for NONE detection and course-code extraction
import sys           # PyInstaller _MEIPASS detection for bundled resource paths
import traceback     # Full stack-trace logging on unexpected errors
from collections import Counter   # Frequency counts for file types and duplicate detection
from datetime import datetime     # Timestamps for log headers and file modification times
from typing import ClassVar       # Type annotation for class-level constants

# ---------------------------------------------------------------------------
# GUI toolkit
# ---------------------------------------------------------------------------
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

# ---------------------------------------------------------------------------
# Excel workbook generation (openpyxl)
# ---------------------------------------------------------------------------
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

# ---------------------------------------------------------------------------
# Image handling (Pillow)
# ---------------------------------------------------------------------------
# Pillow is used to load and resize PNG logos for the title bar.
# If Pillow is not installed the application starts normally with plain-text
# fallbacks in place of the logos — no functionality is lost.
try:
    from PIL import Image, ImageTk   # type: ignore
    _PILLOW_AVAILABLE = True
except ImportError:
    _PILLOW_AVAILABLE = False


# ===========================================================================
# Folder structure profile definitions
# ===========================================================================
#
# Each profile is a dict mapping top-level folder names to a list of their
# expected immediate subfolders.  An empty list means no subfolders are
# expected (the folder may contain files directly, or may be NONE-marked).
#
# Matching is performed case-insensitively, ignoring leading number/letter
# prefixes ("05. ") and NONE markers, so minor renumbering or NONE variants
# on disk do not cause false MISSING flags.

# ---------------------------------------------------------------------------
# Current structure  (default; ~2022 onwards)
# ---------------------------------------------------------------------------
CURRENT_STRUCTURE = {
    "01. Administration": [
        "a. Course handouts",
        "b. Prescribed texts",
        "c. Course evaluations",
        "d. DP list final",
    ],
    "02. Notes": [],
    "03. Lessons": [
        "a. Slides",
        "b. Worked examples",
        "c. Supporting notes",
        "d. Additional material",
    ],
    "04. Tutorials": [
        "a. Instruction sheets",
        "b. Solutions",
        "c. Sample hand-ins",
    ],
    "05. Practicals": [
        "a. Instruction sheets",
        "b. Solutions",
        "c. Sample hand-ins",
    ],
    "06. Assignments": [
        "a. Instruction sheets",
        "b. Solutions",
        "c. Sample hand-ins",
    ],
    "07. Projects": [
        "a. Instruction sheets",
        "b. Sample hand-ins",
    ],
    "08. Tests": [
        "a. Questions",
        "b. Model answers",
        "c. Sample answers from students",
    ],
    "09. Software": [],
    "10. Additional resources": [],
    "11. Other": [],
    "12. Exams main (Admin)": [
        "a. Exam paper",
        "b. Exam model answer",
        "c. External moderator reports",
        "d. Departmental control sheet",
        "e. Exam scripts",
        "f. Mark sheets",
    ],
    "13. Exams SUPP (Admin)": [
        "a. Exam paper",
        "b. Exam model answer",
        "c. External moderator reports",
        "d. Departmental control sheet",
        "e. Exam scripts",
        "f. Mark sheets",
    ],
}

# ---------------------------------------------------------------------------
# Legacy structure  (earlier template)
# ---------------------------------------------------------------------------
# Key differences from Current:
#   - "03. Lessons" has Recordings instead of Worked Examples/Supporting Notes
#   - Many subfolders carry an "(h)" suffix denoting hidden/restricted content
#   - Exam content lives in "09. Exam" and "13. Supplementary exam" (not Admin)
#   - "11. Additional resources" has explicit subfolders
#   - "12. Other" contains Images and Sick Notes
LEGACY_STRUCTURE = {
    "01. Administration": [
        "a. Course handouts",
        "b. Prescribed texts",
        "c. Course evaluations (h)",
        "d. DP list (h)",
    ],
    "02. Notes": [],
    "03. Lessons": [
        "a. Slides NONE",       # Slides were typically not kept; NONE-marked by default
        "b. Recordings",
        "c. Additional material",
    ],
    "04. Tutorials": [
        "a. Instruction sheets",
        "b. Recordings",
        "c. Solutions (h)",
        "d. Sample hand-ins (h)",
    ],
    "05. Practicals": [
        "a. Instruction sheets",
        "b. Recordings",
        "c. Solutions (h)",
        "d. Sample hand-ins (h)",
    ],
    "06. Assignments": [
        "a. Instruction sheets",
        "b. Solutions",         # Non-hidden solutions (e.g. published after deadline)
        "c. Solutions (h)",     # Hidden solutions (restricted access)
        "d. Sample hand-ins (h)",
    ],
    "07. Projects": [
        "a. Instruction sheets",
        "b. Sample hand-ins (h)",
    ],
    "08. Tests": [
        "a. Questions",
        "b. Model answers",
        "c. Sample answers from students (h)",
    ],
    "09. Exam": [
        "a. Exam paper",
        "b. Exam model answer",
        "c. Exam scripts",
        "d. External moderator reports",
        "e. Departmental control sheet",
        "f. Mark sheets",
    ],
    "10. Software": [],
    "11. Additional resources": [
        "a. Past tests and exams",
        "b. datasheets",
        "c. code",
    ],
    "12. Other": [
        "Images (h)",
        "Sick notes (h)",
    ],
    "13. Supplementary exam": [
        "a. Exam paper",
        "b. Exam model answer",
        "c. Exam scripts",
        "d. External moderators report",
        "e. Mark sheets",
    ],
}

# ---------------------------------------------------------------------------
# Updated structure  (new simplified template)
# ---------------------------------------------------------------------------
# Key differences from Current:
#   - Fewer top-level folders (7 vs 13)
#   - "02. Teaching material" replaces the separate Notes folder
#   - "04. Practicals & labs" replaces the standalone Practicals folder
#   - Subfolders renamed from "Instruction sheets" to "Handouts" throughout
#   - A single "07. Exams" folder replaces the two Admin exam folders;
#     it contains "Main exams" and "SUPP exams" sub-groups, each holding
#     the same five subfolders (see UPDATED_EXAM_SUBFOLDERS below).
UPDATED_STRUCTURE = {
    "01. Administration": [
        "a. Course handout",
        "b. DP list final",
        "c. Course evaluation",
    ],
    "02. Teaching material": [
        "a. Notes",
        "b. Prescribed textbooks",
    ],
    "03. Tutorials": [
        "a. Handouts",
        "b. Solutions",
        "c. Hand-ins",
    ],
    "04. Practicals & labs": [
        "a. Handouts",
        "b. Solutions",
        "c. Hand-ins",
    ],
    "05. Assignments": [
        "a. Handouts",
        "b. Solutions",
        "c. Hand-ins",
    ],
    "06. Tests": [
        "a. Handouts",
        "b. Solutions",
        "c. Hand-ins",
    ],
    # 07. Exams uses a two-level hierarchy handled by _evaluate_updated_exam_folder.
    # The list here names the two expected sub-groups; their contents are defined
    # separately in UPDATED_EXAM_SUBFOLDERS.
    "07. Exams": [
        "Main exams",
        "SUPP exams",
    ],
}

# Subfolders expected inside each exam group (Main exams, SUPP exams) in the
# Updated profile.  Both groups share the same five subfolders.
UPDATED_EXAM_SUBFOLDERS = [
    "a. Exam paper",
    "b. Exam model answers",
    "c. External moderator reports",
    "d. Marks sheets",
    "e. Exam scripts",
]

# Master registry — maps profile name (as shown in the UI) to its structure dict.
STRUCTURE_PROFILES = {
    "Current": CURRENT_STRUCTURE,
    "Legacy":  LEGACY_STRUCTURE,
    "Updated": UPDATED_STRUCTURE,
}

# Flat set of all top-level folder base names (number prefix stripped, lowercased,
# trailing-s removed) across every profile.  Used to detect folders that look like
# a renamed or renumbered template folder rather than a genuinely unknown folder.
_ALL_TEMPLATE_BASE_NAMES: set[str] = set()
for _profile in STRUCTURE_PROFILES.values():
    for _name in _profile:
        _base = re.sub(r'^[0-9a-z]+\.\s*', '', _name.strip(), flags=re.IGNORECASE).strip().lower()
        if _base.endswith("s") and len(_base) > 2:
            _base = _base[:-1]
        _ALL_TEMPLATE_BASE_NAMES.add(_base)

# ---------------------------------------------------------------------------
# Global constants
# ---------------------------------------------------------------------------

# File names (lowercased) that are silently ignored during all audit logic.
# Currently limited to the macOS Finder metadata file.
IGNORED_FILES: set[str] = {".ds_store"}

# Pattern that detects files whose stem ends with "unsigned" (case-insensitive),
# optionally preceded by a separator.  Matches before the file extension so
# "report_unsigned.pdf" and "report UNSIGNED.pdf" are both caught.
# Examples: "report_unsigned.pdf", "memo UNSIGNED.docx", "script-unsigned.pdf"
_UNSIGNED_FILE_PATTERN = re.compile(
    r'[\s\-_]unsigned$',
    re.IGNORECASE,
)

# Authorised auditors shown in the Auditor dropdown.
# Add new names here as required; no code changes are needed elsewhere.
AUDIT_USERS: list[str] = [
    "Robyn Verrinder",
    "Yunus Abdul Gaffar",
    "Joyce Mwangama",
    "Janine Buxey",
    "Rachmat Harris",
    "Verona Langenhoven",
]


# ===========================================================================
# Module-level helper functions
# (no GUI dependency — safe to unit-test in isolation)
# ===========================================================================

def _strip_number_prefix(name: str) -> str:
    """Remove a leading alphanumeric prefix and dot from a folder name.

    This normalises folder names so that minor renumbering on disk (e.g. a
    lecturer moving "05. Assignments" to "06. Assignments") does not cause
    false MISSING flags when matched against the template.

    Examples
    --------
    >>> _strip_number_prefix("05. Assignments")
    'Assignments'
    >>> _strip_number_prefix("a. Slides")
    'Slides'
    >>> _strip_number_prefix("No prefix here")
    'No prefix here'
    """
    return re.sub(r'^[0-9a-z]+\.\s*', '', name.strip(), flags=re.IGNORECASE).strip()


# Status-marker words/phrases that administrators sometimes append to folder
# names to flag issues manually.  Pattern only matches at the END of the
# string and requires at least one preceding word character (lookbehind) to
# avoid false positives on legitimate names like "a. Unsigned submissions".
# REVIEW, CHECK, ACTION, PENDING excluded — too common as real folder words.
_STATUS_MARKER_PATTERN = re.compile(
    r'(?<=\w)[\s\-_]+(MISSING|INCOMPLETE|URGENT|TODO|UNSIGNED|EMPTY|TO\s+BE\s+SIGNED)[\s\-_\S]*$',
    re.IGNORECASE,
)


def _strip_status_markers(name: str) -> str:
    """Remove any administrator status-marker annotation from *name*.

    Only strips markers at the end of the string (with a preceding word
    boundary) to avoid false positives on legitimate folder names.
    Compound qualifiers immediately before the marker (e.g. "COR MISSING",
    "SIGNATURES MISSING") are also stripped when doing so leaves a meaningful
    base name of ≥2 words.

    Safe marker words/phrases (trailing only):
      MISSING, INCOMPLETE, URGENT, TODO, UNSIGNED, TO BE SIGNED

    Examples
    --------
    >>> _strip_status_markers("01. Administration - MISSING")
    '01. Administration'
    >>> _strip_status_markers("f. Mark sheets COR MISSING")
    'f. Mark sheets'
    >>> _strip_status_markers("f. Mark sheets TO BE SIGNED")
    'f. Mark sheets'
    >>> _strip_status_markers("c. External moderator reports UNSIGNED")
    'c. External moderator reports'
    >>> _strip_status_markers("b. Review materials")
    'b. Review materials'
    >>> _strip_status_markers("a. Unsigned submissions")
    'a. Unsigned submissions'
    >>> _strip_status_markers("03. Tutorials")
    '03. Tutorials'
    """
    name = name.strip()
    marker_words = r'(MISSING|INCOMPLETE|URGENT|TODO|UNSIGNED|EMPTY|TO\s+BE\s+SIGNED)'
    # Try compound strip: <UPPERCASE-ONLY qualifier word> <marker> [rest] at end of string.
    # [A-Z]{2,} without IGNORECASE ensures only truly uppercase words like "COR",
    # "SIGNATURES" are treated as qualifiers — not natural lowercase words like "answer".
    compound = re.sub(
        rf'([\s\-_]+[A-Z]{{2,}}[\s\-_]+{marker_words})[\s\-_\S]*$',
        '',
        name,
        # No re.IGNORECASE here — qualifier must be uppercase
    ).strip()
    base_words = re.sub(r'^[0-9a-z]+\.\s*', '', compound, flags=re.IGNORECASE).split()
    if len(base_words) >= 2 and compound != name:
        return compound
    # Fallback: strip the marker and everything after it
    return re.sub(
        rf'(?<=\w)[\s\-_]+{marker_words}[\s\-_\S]*$',
        '',
        name,
        flags=re.IGNORECASE,
    ).strip()


def normalised_base_key(name: str, strip_none_fn) -> str:
    """Return a lowercase match key that is insensitive to number prefixes,
    NONE markers, administrator status markers, capitalisation, and
    singular/plural suffixes.

    Applied uniformly to both the template name and the disk name so that,
    for example, all of the following resolve to the same key and match each
    other correctly:

      Disk name                          Key produced
      ---------------------------------  -------------------------
      "a. Course handout"             -> "course handout"
      "a. Course Handout"             -> "course handout"
      "a. Course Handouts"            -> "course handout"
      "a. Course handouts"            -> "course handout"
      "b. Course handout NONE"        -> "course handout"
      "05. Assignments - NONE"        -> "assignment"
      "06. Assignments"               -> "assignment"
      "01. Administration - MISSING"  -> "administration"
      "05. Practicals INCOMPLETE"     -> "practical"

    The trailing-s strip is applied after all other normalisation so it only
    affects the final character of the resolved base name, not mid-word
    letters (e.g. "solutions" → "solution", "slides" → "slide").

    Parameters
    ----------
    name          : Raw folder name (from disk or from the template).
    strip_none_fn : Callable that removes NONE tokens from a name string.
                   Passed in to avoid a circular dependency on the class.
    """
    key = _strip_number_prefix(_strip_status_markers(strip_none_fn(name))).lower()
    # Strip parenthetical suffixes like "(admin)", "(h)" so they don't block
    # the trailing-s normalisation (e.g. "exams supps (admin)" → "exams supps"
    # → "exams supp" matching "exams supp (admin)" → "exams supp").
    key = re.sub(r'\s*\([^)]*\)\s*$', '', key).strip()
    # Strip a trailing 's' so plural/singular variants match each other.
    # Both sides of every comparison go through this function, so the
    # normalisation is symmetric and cannot cause false matches.
    if key.endswith("s") and len(key) > 2:
        key = key[:-1]
    return key


# ===========================================================================
# Main application class
# ===========================================================================

class CourseFolderAuditApp:
    """Tkinter GUI application for auditing EEE course folder structures.

    Responsibilities
    ----------------
    - Present a browse/recent-directory selector and profile/auditor controls.
    - Walk the selected course root and compare it against the active profile.
    - Display results across six tabbed views (log, issues, full check,
      folder details, file details, ASCII tree).
    - Write a plain-text log and a colour-coded Excel workbook to the course
      root directory.

    Colour palette (EEE department branding)
    -----------------------------------------
    Three GUI colours only; UCT dark blue is reserved for Excel headers.
      C_CHROME  #BFCCC2  EEE pale green — window chrome, frames, labels
      C_CONTENT #FFFFFF  white          — treeviews, text areas, entries
      C_INK     #6C9273  EEE mid green  — buttons, treeview headings, accents
    """

    # ------------------------------------------------------------------ #
    #  GUI colour constants                                                 #
    # ------------------------------------------------------------------ #
    C_CHROME          = "#BFCCC2"   # EEE pale green — main window chrome
    C_CONTENT         = "#FFFFFF"   # White          — content area backgrounds
    C_INK             = "#6C9273"   # EEE mid green  — buttons, headings, accents
    C_TEXT            = "#000000"   # Black          — all foreground text
    C_TEXT_ON_CONTENT = "#000000"   # Black          — text inside content areas
    C_UCT_DARK_BLUE   = "#003C69"   # UCT dark blue  — Excel headers only, never in GUI

    # ------------------------------------------------------------------ #
    #  Status colour map (used for Excel conditional formatting)           #
    # ------------------------------------------------------------------ #
    # Each entry maps a status string to (background_hex, font_hex).
    # Font is black throughout for readability against all fill colours.
    STATUS_COLOURS: ClassVar[dict[str, tuple[str, str]]] = {
        "OK":                     ("B6D7A8", "000000"),  # Bold green   — all checks passed
        "EMPTY - REVIEW":         ("FFE599", "000000"),  # Bold yellow  — folder exists but no files
        "MISSING":                ("EA9999", "000000"),  # Bold red     — expected folder absent
        "MISSING CHILDREN":       ("EA9999", "000000"),  # Bold red     — parent present, child missing
        "UNEXPECTED":             ("F9CB9C", "000000"),  # Bold orange  — not in template
        "NONE - ACCEPTED":        ("9FC5E8", "000000"),  # Bold blue    — intentionally empty
        "POPULATED DESPITE NONE": ("B4A7D6", "000000"),  # Bold purple  — NONE folder has content
        "REVIEW - HAND-INS":      ("F9CB9C", "000000"),  # Bold orange  — hand-in validation failed
        "DUPLICATE":              ("EA9999", "000000"),  # Bold red     — name collision on disk
        "ADMIN FLAG":             ("FFD966", "000000"),  # Bold amber   — administrator status marker in folder name
    }

    # ==================================================================== #
    #  Initialisation                                                        #
    # ==================================================================== #

    def __init__(self, root: tk.Tk) -> None:
        """Initialise application state, apply the EEE theme, and build the GUI.

        Parameters
        ----------
        root : The top-level Tk window supplied by ``main()``.
        """
        self.root = root
        self.root.title("Course Audit Tool")
        self.root.geometry("1420x860")
        self.root.minsize(1120, 720)
        self.root.configure(bg=self.C_CHROME)

        # Tkinter control variables bound to UI widgets
        self.selected_directory = tk.StringVar()
        self.profile_mode       = tk.StringVar(value="Auto-detect")
        self.selected_user      = tk.StringVar(value=AUDIT_USERS[0])

        # Path to the JSON file that persists the ten most-recently used directories
        self.recent_dirs_file = os.path.join(
            os.path.expanduser("~"),
            ".course_folder_audit_recent.json",
        )
        self.recent_directories = self.load_recent_directories()

        # Load title-bar logos.  References are kept on self so that Tkinter's
        # garbage collector does not destroy the PhotoImage objects while the
        # window is open (a common Tkinter pitfall with image variables that go
        # out of scope after widget creation).
        self._img_uct = self._load_logo("logo_uct.png", height=36)
        self._img_eee = self._load_logo("logo_eee.png", height=36)

        # Use the UCT logo as the window/taskbar icon if it loaded successfully.
        if self._img_uct:
            self.root.iconphoto(True, self._img_uct)

        self._apply_theme()
        self._build_gui()

    # ==================================================================== #
    #  Logo loading                                                          #
    # ==================================================================== #

    def _load_logo(self, filename: str, height: int):
        """Load *filename* from the same directory as this script, resize it to
        *height* pixels (preserving aspect ratio), and return a Tkinter-
        compatible PhotoImage object.

        Returns None if:
          - Pillow is not installed
          - The file does not exist alongside the script
          - The file cannot be opened as an image

        The caller stores the returned object on self so that Tkinter does
        not garbage-collect it while the window is alive.

        Parameters
        ----------
        filename : PNG filename (e.g. "logo_eee.png").
        height   : Target height in pixels; width is scaled proportionally.

        Returns
        -------
        ImageTk.PhotoImage | None
        """
        if not _PILLOW_AVAILABLE:
            return None

        # Resolve the path relative to this source file so the tool works
        # regardless of the working directory from which it is launched.
        # sys._MEIPASS is set by PyInstaller at runtime; fall back to the
        # script directory when running normally.
        base_dir  = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        logo_path = os.path.join(base_dir, filename)

        if not os.path.isfile(logo_path):
            return None   # File absent — caller will use text fallback

        try:
            img   = Image.open(logo_path).convert("RGBA")
            # Scale width proportionally to the requested height
            ratio = height / img.height
            img   = img.resize(
                (max(1, int(img.width * ratio)), height),
                Image.LANCZOS,
            )
            return ImageTk.PhotoImage(img)
        except Exception:
            return None   # Corrupt file or unsupported format — use fallback

    # ==================================================================== #
    #  Theme                                                                 #
    # ==================================================================== #

    def _apply_theme(self) -> None:
        """Configure ttk styles for the EEE 3-colour palette.

        Applies consistent styling to TNotebook tabs, Treeview rows and
        headings, Combobox fields, and Scrollbar troughs.  The "clam" base
        theme is used because it exposes the most style-configurable options
        across platforms.
        """
        style = ttk.Style(self.root)
        style.theme_use("clam")

        # --- Notebook tabs -----------------------------------------------
        style.configure("TNotebook", background=self.C_CHROME, borderwidth=0)
        style.configure(
            "TNotebook.Tab",
            background=self.C_CHROME,
            foreground="#000000",
            padding=[12, 6],
            font=("Montserrat", 10, "bold"),
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", self.C_CHROME)],
            foreground=[("selected", "#000000")],
        )

        # --- Treeview rows and column headings ---------------------------
        style.configure(
            "Treeview",
            background=self.C_CONTENT,
            foreground=self.C_TEXT,
            fieldbackground=self.C_CONTENT,
            rowheight=24,
            font=("Montserrat", 10),
        )
        style.configure(
            "Treeview.Heading",
            background=self.C_INK,
            foreground="#FFFFFF",
            font=("Montserrat", 10, "bold"),
            relief="flat",
        )
        style.map(
            "Treeview",
            background=[("selected", self.C_INK)],
            foreground=[("selected", self.C_TEXT)],
        )
        style.map("Treeview.Heading", background=[("active", self.C_INK)])

        # --- Combobox (read-only drop-down) ------------------------------
        style.configure(
            "TCombobox",
            fieldbackground=self.C_CONTENT,
            background=self.C_CONTENT,
            foreground=self.C_TEXT,
            arrowcolor=self.C_TEXT,
            selectbackground=self.C_INK,
            selectforeground=self.C_TEXT,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", self.C_CONTENT)],
            foreground=[("readonly", self.C_TEXT)],
        )

        # --- Scrollbars --------------------------------------------------
        style.configure(
            "TScrollbar",
            background=self.C_INK,
            troughcolor=self.C_CHROME,
            arrowcolor=self.C_TEXT,
            bordercolor=self.C_CHROME,
        )

    # ==================================================================== #
    #  GUI construction                                                      #
    # ==================================================================== #

    def _build_gui(self) -> None:
        """Construct the complete GUI layout.

        Layout (top to bottom):
          1. Title bar        — application name and department badge
          2. Controls frame   — recent-directory combo, path entry, profile
                               and auditor selectors, hint label
          3. Button bar       — Run Audit, Clear Output, summary label
          4. Notebook         — six tabs (see _build_*_tab methods)
        """
        # ── Title bar ─────────────────────────────────────────────────────
        # Height is set to 60px to accommodate the logos comfortably.
        title_bar = tk.Frame(self.root, bg=self.C_CHROME, height=60)
        title_bar.pack(fill="x")
        title_bar.pack_propagate(False)  # Prevent children from resizing the bar

        # EEE logo — left side.  Falls back to a plain text label when the
        # image could not be loaded (Pillow absent or file missing).
        if self._img_eee:
            tk.Label(
                title_bar,
                image=self._img_eee,
                bg=self.C_CHROME,
            ).pack(side="left", padx=(16, 8), pady=10)
        else:
            tk.Label(
                title_bar,
                text="EEE",
                font=("Montserrat", 10, "bold"),
                bg=self.C_CHROME, fg="#000000",
            ).pack(side="left", padx=(16, 8), pady=10)

        # Application title — centred between the two logos.
        tk.Label(
            title_bar,
            text="Course Audit Tool",
            font=("Montserrat", 16, "bold"),
            bg=self.C_CHROME, fg="#000000",
        ).pack(side="left", padx=8, pady=14)

        # UCT logo — right side.  Falls back to a plain text label.
        if self._img_uct:
            tk.Label(
                title_bar,
                image=self._img_uct,
                bg=self.C_CHROME,
            ).pack(side="right", padx=(8, 16), pady=10)
        else:
            tk.Label(
                title_bar,
                text="UCT",
                font=("Montserrat", 10, "bold"),
                bg=self.C_CHROME, fg="#000000",
            ).pack(side="right", padx=(8, 16), pady=10)

        # ── Controls frame ─────────────────────────────────────────────────
        main_frame = tk.Frame(self.root, bg=self.C_CHROME)
        main_frame.pack(fill="x", padx=15, pady=10)

        # Local factory functions keep widget creation concise
        def lbl(parent, text):
            """Return a themed Label widget."""
            return tk.Label(
                parent, text=text,
                bg=self.C_CHROME, fg=self.C_TEXT,
                font=("Montserrat", 10),
            )

        def btn(parent, text, command, width=15):
            """Return a themed Button widget."""
            return tk.Button(
                parent, text=text, command=command, width=width,
                bg=self.C_INK, fg="#000000",
                activebackground="#3A3536", activeforeground="#000000",
                relief="flat", cursor="hand2",
                font=("Montserrat", 10, "bold"),
                padx=6, pady=4,
            )

        # Recent directories row
        lbl(main_frame, "Recent Directories:").grid(
            row=0, column=0, sticky="w", pady=(4, 2))

        self.recent_combo = ttk.Combobox(
            main_frame, values=self.recent_directories,
            state="readonly", width=110,
        )
        self.recent_combo.grid(row=1, column=0, sticky="ew", padx=(0, 8))
        self.recent_combo.bind("<<ComboboxSelected>>", self.select_recent_directory)

        btn(main_frame, "Use Selected",
            self.use_selected_recent_directory).grid(row=1, column=1)

        # Selected path row
        lbl(main_frame, "Selected Course Root:").grid(
            row=2, column=0, sticky="w", pady=(12, 2))

        tk.Entry(
            main_frame,
            textvariable=self.selected_directory,
            width=110,
            bg=self.C_CHROME, fg=self.C_TEXT,
            insertbackground=self.C_TEXT,
            relief="flat",
            font=("Montserrat", 10),
        ).grid(row=3, column=0, sticky="ew", padx=(0, 8), ipady=4)

        btn(main_frame, "Browse...", self.browse_directory).grid(row=3, column=1)

        # Profile + auditor selectors (side by side)
        sub = tk.Frame(main_frame, bg=self.C_CHROME)
        sub.grid(row=4, column=0, columnspan=2, sticky="w", pady=(12, 0))

        lbl(sub, "Folder structure profile:").grid(
            row=0, column=0, sticky="w", padx=(0, 40))
        lbl(sub, "Auditor:").grid(row=0, column=1, sticky="w")

        ttk.Combobox(
            sub, textvariable=self.profile_mode,
            values=["Auto-detect", "Current", "Legacy", "Updated"],
            state="readonly", width=28,
        ).grid(row=1, column=0, sticky="w", padx=(0, 40), pady=(4, 0))

        ttk.Combobox(
            sub, textvariable=self.selected_user,
            values=AUDIT_USERS,
            state="readonly", width=28,
        ).grid(row=1, column=1, sticky="w", pady=(4, 0))

        # Hint label explaining the NONE convention to first-time users
        tk.Label(
            main_frame,
            text=(
                "Supports Current, Legacy, and Updated course-folder structures.  "
                "Folders containing 'NONE' (any position or separator) are treated "
                "as intentionally empty unless populated."
            ),
            bg=self.C_CHROME, fg=self.C_TEXT,
            font=("Montserrat", 9), anchor="w", justify="left",
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(8, 0))

        main_frame.columnconfigure(0, weight=1)   # Let the path entry stretch

        # ── Button bar ────────────────────────────────────────────────────
        button_frame = tk.Frame(self.root, bg=self.C_CHROME)
        button_frame.pack(fill="x", padx=15, pady=(0, 8))

        btn(button_frame, "Run Audit and Create Outputs",
            self.scan_and_export, width=28).pack(side="left", padx=(0, 8))
        btn(button_frame, "Clear Output",
            self.clear_output, width=14).pack(side="left", padx=(0, 8))

        # Inline summary updated after each successful audit run
        self.summary_label = tk.Label(
            button_frame, text="No audit run yet.",
            bg=self.C_CHROME, fg=self.C_TEXT,
            font=("Montserrat", 10), anchor="w",
        )
        self.summary_label.pack(side="left", padx=12)

        # ── Notebook (six tabs) ────────────────────────────────────────────
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=15, pady=(0, 12))

        # Create a plain Frame for each tab; content is added by the
        # corresponding _build_*_tab() helper below.
        self.output_tab   = tk.Frame(self.notebook, bg=self.C_CONTENT)
        self.issues_tab   = tk.Frame(self.notebook, bg=self.C_CONTENT)
        self.expected_tab = tk.Frame(self.notebook, bg=self.C_CONTENT)
        self.folder_tab   = tk.Frame(self.notebook, bg=self.C_CONTENT)
        self.file_tab     = tk.Frame(self.notebook, bg=self.C_CONTENT)
        self.tree_tab     = tk.Frame(self.notebook, bg=self.C_CONTENT)

        self.notebook.add(self.output_tab,   text="  Log Output  ")
        self.notebook.add(self.issues_tab,   text="  ⚠ Issues  ")
        self.notebook.add(self.expected_tab, text="  Expected Structure Check  ")
        self.notebook.add(self.folder_tab,   text="  Folder Details  ")
        self.notebook.add(self.file_tab,     text="  File Details  ")
        self.notebook.add(self.tree_tab,     text="  Tree Diagram  ")

        self._build_output_tab()
        self._build_issues_tab()
        self._build_expected_tab()
        self._build_folder_tab()
        self._build_file_tab()
        self._build_tree_tab()

    # ------------------------------------------------------------------ #
    #  Tab builders                                                         #
    # ------------------------------------------------------------------ #

    def _build_output_tab(self) -> None:
        """Populate the Log Output tab with a scrolled, word-wrapped text area."""
        self.output_text = scrolledtext.ScrolledText(
            self.output_tab, wrap=tk.WORD, width=140, height=32,
            bg=self.C_CONTENT, fg=self.C_TEXT,
            insertbackground=self.C_TEXT,
            selectbackground=self.C_INK, selectforeground=self.C_TEXT,
            font=("Montserrat", 10), relief="flat",
        )
        self.output_text.pack(fill="both", expand=True, padx=2, pady=2)

    def _build_structure_treeview(self, parent: tk.Frame) -> ttk.Treeview:
        """Build and return a Treeview with the standard expected-structure columns.

        Shared by both the Issues tab and the Expected Structure Check tab so
        that column definitions and widths stay consistent.  Horizontal and
        vertical scrollbars are attached automatically.

        Parameters
        ----------
        parent : The tab Frame that will contain the Treeview.

        Returns
        -------
        ttk.Treeview
            Fully configured Treeview widget, already packed into *parent*.
        """
        columns = (
            "relative_path", "level", "expected_name",
            "actual_name", "exists", "status", "details",
        )
        tv = ttk.Treeview(parent, columns=columns, show="headings")

        headings = {
            "relative_path": "Parent Path",
            "level":         "Level",
            "expected_name": "Expected Name",
            "actual_name":   "Actual Name",
            "exists":        "Exists",
            "status":        "Status",
            "details":       "Details",
        }
        widths = {
            "relative_path": 240,
            "level":         80,
            "expected_name": 250,
            "actual_name":   250,
            "exists":        70,
            "status":        170,
            "details":       340,
        }
        for col in columns:
            tv.heading(col, text=headings[col])
            tv.column(col, width=widths[col], anchor="w")

        y_scroll = ttk.Scrollbar(parent, orient="vertical",   command=tv.yview)
        x_scroll = ttk.Scrollbar(parent, orient="horizontal", command=tv.xview)
        tv.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)

        tv.pack(side="left",   fill="both", expand=True)
        y_scroll.pack(side="right",  fill="y")
        x_scroll.pack(side="bottom", fill="x")
        return tv

    def _build_issues_tab(self) -> None:
        """Populate the Issues tab with a structure Treeview (actionable rows only)."""
        self.issues_treeview = self._build_structure_treeview(self.issues_tab)

    def _build_expected_tab(self) -> None:
        """Populate the Expected Structure Check tab with a full structure Treeview."""
        self.expected_treeview = self._build_structure_treeview(self.expected_tab)

    def _build_folder_tab(self) -> None:
        """Populate the Folder Details tab with a Treeview showing per-folder metrics."""
        columns = (
            "folder", "depth", "subfolder_count", "file_count",
            "total_size", "last_modified", "type_counts", "status",
        )
        self.folder_treeview = ttk.Treeview(self.folder_tab, columns=columns, show="headings")

        headings = {
            "folder":          "Folder",
            "depth":           "Depth",
            "subfolder_count": "Subfolders",
            "file_count":      "Files",
            "total_size":      "Folder File Size",
            "last_modified":   "Latest Modified",
            "type_counts":     "File Type Counts",
            "status":          "Status",
        }
        widths = {
            "folder": 320, "depth": 60, "subfolder_count": 80, "file_count": 60,
            "total_size": 120, "last_modified": 150, "type_counts": 340, "status": 180,
        }
        for col in columns:
            self.folder_treeview.heading(col, text=headings[col])
            self.folder_treeview.column(col, width=widths[col], anchor="w")

        y_scroll = ttk.Scrollbar(
            self.folder_tab, orient="vertical",   command=self.folder_treeview.yview)
        x_scroll = ttk.Scrollbar(
            self.folder_tab, orient="horizontal", command=self.folder_treeview.xview)
        self.folder_treeview.configure(
            yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.folder_treeview.pack(side="left", fill="both", expand=True)
        y_scroll.pack(side="right",  fill="y")
        x_scroll.pack(side="bottom", fill="x")

    def _build_file_tab(self) -> None:
        """Populate the File Details tab with a Treeview listing every file found."""
        columns = ("directory", "name", "type", "size", "modified")
        self.file_treeview = ttk.Treeview(self.file_tab, columns=columns, show="headings")

        for col, title, width in [
            ("directory", "Directory", 320),
            ("name",      "File Name", 300),
            ("type",      "Type",      100),
            ("size",      "Size",      100),
            ("modified",  "Modified",  160),
        ]:
            self.file_treeview.heading(col, text=title)
            self.file_treeview.column(col, width=width, anchor="w")

        y_scroll = ttk.Scrollbar(
            self.file_tab, orient="vertical",   command=self.file_treeview.yview)
        x_scroll = ttk.Scrollbar(
            self.file_tab, orient="horizontal", command=self.file_treeview.xview)
        self.file_treeview.configure(
            yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.file_treeview.pack(side="left", fill="both", expand=True)
        y_scroll.pack(side="right",  fill="y")
        x_scroll.pack(side="bottom", fill="x")

    def _build_tree_tab(self) -> None:
        """Populate the Tree Diagram tab with a horizontally-scrollable text area."""
        self.tree_text = scrolledtext.ScrolledText(
            self.tree_tab,
            wrap=tk.NONE,        # No line-wrapping: long paths scroll horizontally
            width=140, height=32,
            bg=self.C_CONTENT, fg=self.C_TEXT_ON_CONTENT,
            insertbackground=self.C_TEXT_ON_CONTENT,
            selectbackground=self.C_CHROME, selectforeground=self.C_TEXT,
            font=("Montserrat", 10), relief="flat",
        )
        self.tree_text.pack(fill="both", expand=True, padx=2, pady=2)

    # ------------------------------------------------------------------ #
    #  Logging                                                              #
    # ------------------------------------------------------------------ #

    def log_message(self, message: str) -> None:
        """Append *message* to the Log Output tab and scroll to the bottom.

        Calls ``update_idletasks()`` so progress is visible during a long scan
        without the main event loop being blocked.
        """
        self.output_text.insert(tk.END, message + "\n")
        self.output_text.see(tk.END)
        self.root.update_idletasks()

    # ==================================================================== #
    #  Name and path utilities                                               #
    # ==================================================================== #

    def normalise_for_match(self, name: str) -> str:
        """Return *name* stripped of leading/trailing whitespace and lowercased.

        Used for simple case-insensitive comparisons that do not need prefix
        or NONE stripping.
        """
        return name.strip().lower()

    def has_none_suffix(self, name: str) -> bool:
        """Return True if *name* contains the word NONE as a discrete token.

        The match is case-insensitive and accepts NONE in any position,
        separated by any combination of spaces, dashes, or underscores.

        Examples that match:     "a. Slides NONE", "a. Slides - NONE",
                                 "NONE slides", "a. Slides--NONE"
        Examples that do not:    "nonevent" (adjacent letters block the match)
        """
        return bool(
            re.search(r'(?<![A-Za-z])NONE(?![A-Za-z])', name.strip(), re.IGNORECASE)
        )

    def strip_none_suffix(self, name: str) -> str:
        """Remove the NONE token and any surrounding separators from *name*.

        Works regardless of the position of NONE within the string and the
        style of surrounding separators.

        Examples
        --------
        >>> strip_none_suffix("a. Slides - NONE")
        'a. Slides'
        >>> strip_none_suffix("NONE slides")
        'slides'
        >>> strip_none_suffix("a. Slides")
        'a. Slides'
        """
        cleaned = re.sub(
            r'[\s\-_]*(?<![A-Za-z])NONE(?![A-Za-z])[\s\-_]*',
            ' ',
            name.strip(),
            flags=re.IGNORECASE,
        )
        return cleaned.strip()

    def _nbk(self, name: str) -> str:
        """Shorthand: return the normalised base key for *name*.

        Delegates to the module-level ``normalised_base_key`` function,
        passing ``self.strip_none_suffix`` as the NONE-stripping callback.
        This avoids repeating the callback argument at every call site.
        """
        return normalised_base_key(name, self.strip_none_suffix)

    def _is_under_submission_folder(self, relative_path: str) -> bool:
        """Return True if any segment of *relative_path* is a submission folder.

        Used to apply the ≥15-files-of-the-same-type validation rule to
        deeply nested subdirectories (e.g. practical_1 inside
        "c. Sample hand-ins", or a student sub-folder inside "e. Exam scripts").

        Parameters
        ----------
        relative_path : Forward-slash-separated path relative to the course root.
        """
        parts = relative_path.replace("\\", "/").split("/")
        return any(self.is_submission_folder(part) for part in parts)

    def is_submission_folder(self, folder_name: str) -> bool:
        """Return True if *folder_name* is a submission folder requiring ≥15 files.

        Two categories are matched (case-insensitive):

        1. Sample hand-ins / sample answers — names containing "sample" AND
           either "hand" or "answer":
             - "c. Sample hand-ins"
             - "Sample hand-ins (h)"
             - "c. Sample answers from students"
             - "c. Sample answers from students (h)"

        2. Scripts folders — names containing "script":
             - "e. Exam scripts"
             - "c. Exam scripts"
             - "Exam Scripts (h)"

        Both categories enforce the same rule: each leaf sub-group must contain
        at least 15 files, all of the same file type.
        """
        name_lower = folder_name.lower()
        is_sample  = "sample" in name_lower and ("hand" in name_lower or "answer" in name_lower)
        is_scripts = "script" in name_lower
        return is_sample or is_scripts

    # ------------------------------------------------------------------ #
    #  Folder-type classifiers and targeted content checks                 #
    # ------------------------------------------------------------------ #
    #
    # Each classifier identifies a specific folder category by name.
    # The corresponding checker validates the files inside it, returning
    # (status, detail) in the same format as check_sample_handins so that
    # _evaluate_children can call all checks uniformly.
    #
    # Acceptable-type sets are defined as module-level constants below the
    # class; they are referenced here by name for clarity.

    def is_course_handout_folder(self, folder_name: str) -> bool:
        """Return True if *folder_name* is the course handout folder.

        Matches names containing "course handout" (singular or plural),
        covering all profile variants:
          - "a. Course handouts"  (Current / Legacy)
          - "a. Course handout"   (Updated)
        """
        return "course handout" in folder_name.lower()

    def check_course_handout(self, folder_path: str) -> tuple[str, str]:
        """Validate that the course handout folder contains at least one
        PDF or Word document (.pdf, .doc, .docx).

        Returns
        -------
        tuple[str, str]
            (status, detail) — status is "OK" or "EMPTY - REVIEW".
        """
        return self._check_required_types(
            folder_path,
            required_extensions={".pdf", ".doc", ".docx"},
            friendly_types="PDF or Word document (.pdf, .doc, .docx)",
            folder_label="Course handout",
        )

    def is_dp_list_folder(self, folder_name: str) -> bool:
        """Return True if *folder_name* is the DP list folder.

        Matches names containing "dp list", covering all profile variants:
          - "d. DP list final"  (Current)
          - "d. DP list (h)"    (Legacy)
          - "b. DP list"        (Updated)
        """
        return "dp list" in folder_name.lower()

    def check_dp_list(self, folder_path: str) -> tuple[str, str]:
        """Validate that the DP list folder contains at least one
        PDF or spreadsheet file (.pdf, .xls, .xlsx, .csv).

        Returns
        -------
        tuple[str, str]
            (status, detail) — status is "OK" or "EMPTY - REVIEW".
        """
        return self._check_required_types(
            folder_path,
            required_extensions={".pdf", ".xls", ".xlsx", ".csv"},
            friendly_types="PDF or spreadsheet (.pdf, .xls, .xlsx, .csv)",
            folder_label="DP list",
        )

    def is_mark_sheets_folder(self, folder_name: str) -> bool:
        """Return True if *folder_name* is the mark sheets folder under Exams.

        Matches names containing "mark sheet" or "marks sheet" (singular or
        plural), covering all profile variants:
          - "f. Mark sheets"   (Current / Legacy)
          - "d. Marks sheets"  (Updated)
        """
        name_lower = folder_name.lower()
        return "mark sheet" in name_lower or "marks sheet" in name_lower

    def check_mark_sheets(self, folder_path: str) -> tuple[str, str]:
        """Validate that the mark sheets folder contains at least one
        PDF or spreadsheet file (.pdf, .xls, .xlsx, .csv).

        Returns
        -------
        tuple[str, str]
            (status, detail) — status is "OK" or "EMPTY - REVIEW".
        """
        return self._check_required_types(
            folder_path,
            required_extensions={".pdf", ".xls", ".xlsx", ".csv"},
            friendly_types="PDF or spreadsheet (.pdf, .xls, .xlsx, .csv)",
            folder_label="Mark sheets",
        )

    def is_external_moderator_folder(self, folder_name: str) -> bool:
        """Return True if *folder_name* is the external moderator reports folder.

        Matches names containing "external moderator", covering all profile
        variants:
          - "c. External moderator reports"   (Current / Legacy)
          - "c. External moderator reports"   (Updated)
          - "d. External moderators report"   (Legacy Supplementary Exam)
        """
        return "external moderator" in folder_name.lower()

    def check_external_moderator(self, folder_path: str) -> tuple[str, str]:
        """Validate that the external moderator reports folder contains at least
        one PDF or Word document (.pdf, .doc, .docx).

        Returns
        -------
        tuple[str, str]
            (status, detail) — status is "OK" or "EMPTY - REVIEW".
        """
        return self._check_required_types(
            folder_path,
            required_extensions={".pdf", ".doc", ".docx"},
            friendly_types="PDF or Word document (.pdf, .doc, .docx)",
            folder_label="External moderator reports",
        )

    def _check_required_types(
        self,
        folder_path: str,
        required_extensions: set[str],
        friendly_types: str,
        folder_label: str,
    ) -> tuple[str, str]:
        """Generic validator: check that *folder_path* contains at least one
        file whose extension is in *required_extensions*.

        Searches recursively so files nested in sub-folders are also found.
        Files in IGNORED_FILES are excluded.

        Returns "EMPTY - REVIEW" in three cases:
          1. The folder is completely empty.
          2. The folder has files, but none match the required extensions.

        Parameters
        ----------
        folder_path         : Absolute path to the folder to validate.
        required_extensions : Set of lowercase extensions including the dot
                              (e.g. {".pdf", ".doc", ".docx"}).
        friendly_types      : Human-readable description of acceptable types,
                              used in the detail message shown to auditors.
        folder_label        : Short folder name used in the OK detail message.

        Returns
        -------
        tuple[str, str]
            (status, detail) where status is "OK" or "EMPTY - REVIEW".
        """
        all_files: list[str] = []
        matching:  list[str] = []

        for dirpath, _, filenames in os.walk(folder_path):
            for fname in filenames:
                if fname.lower() in IGNORED_FILES:
                    continue
                all_files.append(fname)
                ext = os.path.splitext(fname)[1].lower()
                if ext in required_extensions:
                    matching.append(fname)

        if not all_files:
            return (
                "EMPTY - REVIEW",
                f"{folder_label}: folder is empty",
            )
        if not matching:
            found_types = sorted({os.path.splitext(f)[1].lower() for f in all_files})
            return (
                "EMPTY - REVIEW",
                f"{folder_label}: no {friendly_types} found "
                f"(found: {', '.join(found_types) or 'no ext'})",
            )
        return (
            "OK",
            f"{folder_label}: {len(matching)} qualifying file"
            f"{'s' if len(matching) != 1 else ''} found",
        )

    def get_relative_directory(self, root_path: str, dirpath: str) -> str:
        """Return the path of *dirpath* relative to *root_path*, using
        forward slashes as separators on all platforms."""
        return os.path.relpath(dirpath, root_path).replace("\\", "/")

    def get_file_extension(self, filename: str) -> str:
        """Return the lower-cased file extension of *filename* (including the
        leading dot), or the sentinel string '[no ext]' when absent."""
        _, ext = os.path.splitext(filename)
        return ext.lower() if ext else "[no ext]"

    def get_depth(self, relative_path: str) -> int:
        """Return the folder depth where 1 = direct child of the course root.

        Parameters
        ----------
        relative_path : Forward-slash-separated path returned by
                        ``get_relative_directory``.
        """
        parts = relative_path.replace("\\", "/").split("/")
        return len([p for p in parts if p and p != "."])

    # ==================================================================== #
    #  Folder content utilities                                              #
    # ==================================================================== #

    def folder_has_content(self, folder_path: str) -> bool:
        """Return True if *folder_path* (or any descendant) contains at least
        one real file, ignoring IGNORED_FILES.

        Empty subfolders do NOT count as content.  This means a NONE-marked
        folder whose only children are empty subdirectories is correctly
        treated as "NONE - ACCEPTED" rather than "POPULATED DESPITE NONE".

        The function is recursive and silently skips directories it cannot
        read (e.g. due to permissions errors).
        """
        try:
            for entry in os.scandir(folder_path):
                if entry.name.lower() in IGNORED_FILES:
                    continue
                if entry.is_file():
                    return True
                if entry.is_dir() and self.folder_has_content(entry.path):
                    return True
        except OSError:
            pass
        return False

    def folder_has_direct_files(self, folder_path: str) -> bool:
        """Return True if *folder_path* contains at least one file directly
        (not inside a subfolder), ignoring IGNORED_FILES.

        Used to detect the case where a lecturer has placed files directly
        inside a folder that would normally contain subfolders (e.g. dropping
        slides directly into "03. Lessons" rather than "03. Lessons/a. Slides").
        In that situation the folder should be treated as OK rather than flagging
        all expected subfolders as MISSING.
        """
        try:
            for entry in os.scandir(folder_path):
                if entry.is_file() and entry.name.lower() not in IGNORED_FILES:
                    return True
        except OSError:
            pass
        return False

    def latest_modified_in_folder(self, folder_path: str) -> str:
        """Return the most recent modification timestamp of any direct file
        inside *folder_path* as a "YYYY-MM-DD HH:MM:SS" string.

        Returns "N/A" when the folder is empty or unreadable.  Only direct
        children are examined (not recursive descendants) to keep the value
        meaningful as a per-folder metric in the Folder Details tab.
        """
        latest = None
        try:
            for entry in os.scandir(folder_path):
                if entry.is_file() and entry.name.lower() not in IGNORED_FILES:
                    try:
                        mtime = entry.stat().st_mtime
                        if latest is None or mtime > latest:
                            latest = mtime
                    except OSError:
                        pass
        except OSError:
            pass
        return (
            datetime.fromtimestamp(latest).strftime("%Y-%m-%d %H:%M:%S")
            if latest else "N/A"
        )

    def format_file_size(self, size_bytes) -> str:
        """Convert *size_bytes* to a human-readable string with the appropriate
        binary unit (B, KB, MB, GB, TB, or PB).

        Returns "Unknown" when *size_bytes* is None (e.g. after an OSError
        during stat collection).
        """
        if size_bytes is None:
            return "Unknown"
        for unit in ("B", "KB", "MB", "GB", "TB"):
            if size_bytes < 1024:
                return f"{size_bytes:.1f} {unit}"
            size_bytes /= 1024
        return f"{size_bytes:.1f} PB"

    def _collect_leaf_file_groups(self, folder_path: str) -> list[tuple[str, list[str]]]:
        """Recursively collect file lists from every leaf directory under
        *folder_path*.

        A "leaf directory" is any directory that contains at least one file.
        Each such directory contributes one (folder_name, file_names) tuple to
        the returned collection.  This allows ``check_sample_handins`` to
        validate each submission sub-group (e.g. practical_1, practical_2)
        independently and report exactly which group has an issue.

        Parameters
        ----------
        folder_path : Absolute path to the folder to scan.

        Returns
        -------
        list[tuple[str, list[str]]]
            One (folder_name, file_names) tuple per leaf directory found.
        """
        groups: list[tuple[str, list[str]]] = []

        def _walk(path: str, name: str) -> None:
            try:
                entries = [e for e in os.scandir(path) if e.name.lower() not in IGNORED_FILES]
            except OSError:
                return
            files   = [e.name for e in entries if e.is_file()]
            subdirs = [e for e in entries if e.is_dir()]
            if files and not subdirs:
                # True leaf: files with no subfolders — count this as a group
                groups.append((name, files))
            for sd in subdirs:
                # If subfolders exist, walk into them and ignore any direct files
                # (e.g. a readme or marking note sitting alongside prac subfolders)
                _walk(sd.path, sd.name)

        _walk(folder_path, os.path.basename(folder_path))
        return groups

    # Extensions counted as valid student submissions.
    # Mixed types are allowed — only the count of these qualifying files matters.
    SUBMISSION_EXTENSIONS: ClassVar[set[str]] = {".pdf", ".doc", ".docx"}

    def check_sample_handins(self, folder_path: str) -> tuple[str, str]:
        """Validate that a submission folder meets the minimum-file rules.

        Applies to any folder identified by ``is_submission_folder`` — i.e.
        sample hand-ins, sample answers, and scripts folders
        (e.g. "e. Exam scripts").

        Each leaf group (e.g. one per student sub-folder or practical task)
        must independently contain at least 15 PDF or Word documents
        (.pdf, .doc, .docx).  Other file types (e.g. .yml, .zip, .png) are
        ignored and do NOT cause a mixed-type flag.

        Groups are evaluated independently so a passing group cannot mask a
        failing sibling.

        Parameters
        ----------
        folder_path : Absolute path to the submission folder.

        Returns
        -------
        tuple[str, str]
            (status_string, detail_string) where status is one of
            "OK", "EMPTY - REVIEW", or "REVIEW - HAND-INS".
        """
        groups = self._collect_leaf_file_groups(folder_path)

        if not groups:
            return "EMPTY - REVIEW", "Subfolder exists but appears empty"

        group_issues: list[str] = []
        for group_name, files in groups:
            qualifying = [
                f for f in files
                if os.path.splitext(f)[1].lower() in self.SUBMISSION_EXTENSIONS
            ]
            count = len(qualifying)
            if count < 15:
                other_count = len(files) - count
                other_note  = f", {other_count} other file type(s) ignored" if other_count else ""
                group_issues.append(
                    f'"{group_name}": {count} PDF/Word doc{"s" if count != 1 else ""} (expected >=15{other_note})'
                )

        if not group_issues:
            total_qualifying = sum(
                len([f for f in files if os.path.splitext(f)[1].lower() in self.SUBMISSION_EXTENSIONS])
                for _, files in groups
            )
            return (
                "OK",
                f"Submissions: {total_qualifying} PDF/Word docs across "
                f"{len(groups)} group{'s' if len(groups) != 1 else ''}",
            )

        return (
            "REVIEW - HAND-INS",
            f"{len(group_issues)} group{'s' if len(group_issues) != 1 else ''} "
            f"with issues: " + " | ".join(group_issues),
        )

    def build_ascii_tree(self, root_path: str) -> str:
        """Build and return a Unicode box-drawing ASCII tree of the directory
        rooted at *root_path*.

        Directories are listed before files at each level; entries are sorted
        case-insensitively.  IGNORED_FILES are excluded.  The root folder
        itself appears as the first line.
        """
        lines: list[str] = []

        def _walk(path: str, prefix: str) -> None:
            try:
                # Sort: directories before files, both case-insensitive
                entries = sorted(
                    os.scandir(path),
                    key=lambda e: (e.is_file(), e.name.lower()),
                )
            except OSError:
                return
            entries = [e for e in entries if e.name.lower() not in IGNORED_FILES]
            for index, entry in enumerate(entries):
                is_last   = index == len(entries) - 1
                connector = "└── " if is_last else "├── "
                lines.append(prefix + connector + entry.name)
                if entry.is_dir():
                    # Indent continuation: blank space under last entry,
                    # vertical bar under all others
                    extension = "    " if is_last else "│   "
                    _walk(entry.path, prefix + extension)

        lines.append(os.path.basename(root_path) or root_path)
        _walk(root_path, "")
        return "\n".join(lines)

    # ==================================================================== #
    #  Recent-directory persistence                                          #
    # ==================================================================== #

    def load_recent_directories(self) -> list[str]:
        """Load the persisted recent-directories list from the JSON cache file.

        Returns an empty list if the file does not exist or cannot be parsed,
        so the application starts cleanly on first run or after corruption.
        """
        if os.path.exists(self.recent_dirs_file):
            try:
                with open(self.recent_dirs_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, list):
                        return data
            except (OSError, json.JSONDecodeError):
                pass
        return []

    def save_recent_directories(self) -> None:
        """Persist the current recent-directories list to the JSON cache file.

        Errors are surfaced as log messages rather than exceptions so that a
        write failure never interrupts an audit run.
        """
        try:
            with open(self.recent_dirs_file, "w", encoding="utf-8") as f:
                json.dump(self.recent_directories, f, indent=4)
        except OSError as exc:
            self.log_message(f"Warning: Could not save recent directories: {exc}")

    def update_recent_directories(self, directory: str) -> None:
        """Prepend *directory* to the recent list, capping it at ten entries.

        If *directory* is already present it is moved to the front (most
        recently used) rather than duplicated.  The Combobox values and the
        JSON cache are updated immediately.
        """
        if directory in self.recent_directories:
            self.recent_directories.remove(directory)
        self.recent_directories.insert(0, directory)
        self.recent_directories = self.recent_directories[:10]
        self.recent_combo["values"] = self.recent_directories
        self.save_recent_directories()

    # ------------------------------------------------------------------ #
    #  Directory selection callbacks                                        #
    # ------------------------------------------------------------------ #

    def browse_directory(self) -> None:
        """Open a native directory-picker dialog and apply the chosen path."""
        directory = filedialog.askdirectory()
        if directory:
            self.selected_directory.set(directory)
            self.update_recent_directories(directory)
            self.log_message(f"Selected directory: {directory}")

    def select_recent_directory(self, event=None) -> None:
        """Copy the Combobox selection into the path entry field.

        Bound to the ``<<ComboboxSelected>>`` virtual event so the entry
        updates as soon as the user picks from the dropdown.
        """
        selected = self.recent_combo.get()
        if selected:
            self.selected_directory.set(selected)

    def use_selected_recent_directory(self) -> None:
        """Apply the current Combobox selection and log the choice.

        Triggered by the "Use Selected" button so the user can explicitly
        confirm the recent-directory selection before running an audit.
        """
        selected = self.recent_combo.get()
        if selected:
            self.selected_directory.set(selected)
            self.log_message(f"Selected recent directory: {selected}")
        else:
            messagebox.showwarning("No Selection", "Please select a recent directory first.")

    def clear_output(self) -> None:
        """Reset all output widgets to their initial empty state.

        Clears the log text area, tree diagram, all four Treeview tables,
        and the summary label in the button bar.
        """
        self.output_text.delete("1.0", tk.END)
        self.tree_text.delete("1.0", tk.END)
        for tree in (
            self.issues_treeview,
            self.expected_treeview,
            self.folder_treeview,
            self.file_treeview,
        ):
            for item in tree.get_children():
                tree.delete(item)
        self.summary_label.config(text="No audit run yet.")

    # ==================================================================== #
    #  Profile detection                                                     #
    # ==================================================================== #

    def detect_profile(self, root_path: str) -> str:
        """Determine which folder-structure profile best matches *root_path*.

        If the user has manually selected a profile in the dropdown (i.e. any
        value other than "Auto-detect"), that selection is returned immediately
        without inspecting the disk.

        Auto-detection works by scoring the top-level directory names against
        a set of unique marker folders for each profile.  The profile with the
        highest score wins; ties are broken in favour of Updated > Legacy >
        Current.  "Current" is returned as the safe default when no markers
        match at all.

        Parameters
        ----------
        root_path : Absolute path to the course root directory.

        Returns
        -------
        str
            One of "Current", "Legacy", or "Updated".
        """
        mode = self.profile_mode.get()
        if mode in STRUCTURE_PROFILES:
            return mode   # User has overridden auto-detection

        try:
            top_level_dirs = [
                name for name in os.listdir(root_path)
                if os.path.isdir(os.path.join(root_path, name))
            ]
        except OSError:
            return "Current"

        # Normalised base keys for all top-level folders actually on disk
        names = {self._nbk(name) for name in top_level_dirs}

        # Marker folders unique to each profile (after normalisation).
        # NOTE: "exams" is NOT used as an Updated marker because _nbk reduces
        # it to "exam", which collides with the Legacy marker and causes ties.
        # Instead we use folder count: Updated has 7 top-level folders,
        # Current/Legacy have 13.  Count <= 9 gives Updated a bonus point.
        current_markers = {
            "exams main (admin)",   # 12. Exams main (Admin)
            "exams supp (admin)",   # 13. Exams SUPP (Admin)
        }
        legacy_markers = {
            "exam",                 # 09. Exam  (NOT "exams" — that's Updated)
            "supplementary exam",   # 13. Supplementary exam
        }
        updated_markers = {
            "teaching material",    # 02. Teaching material
            "practicals & lab",     # 04. Practicals & labs (already _nbk-normalised)
        }

        current_score = sum(1 for m in current_markers if m in names)
        # Legacy "exam" marker must not fire on "exams" (Updated). Only count
        # it when there is no folder whose _nbk is exactly "exam" AND the disk
        # also has a folder that _nbk's to "supplementary exam" — or use the
        # raw marker match but guard against the Updated "exams" folder.
        legacy_score  = sum(1 for m in legacy_markers  if m in names)
        updated_score = sum(1 for m in updated_markers if m in names)

        # Folder-count tiebreaker: Updated uses 7 top-level folders; Current
        # and Legacy use 13.  A short folder count is a strong Updated signal.
        if len(top_level_dirs) <= 9:
            updated_score += 1

        best = max(current_score, legacy_score, updated_score)
        if best == 0:
            return "Current"        # No markers found — default to Current
        if updated_score == best:
            return "Updated"
        if legacy_score == best:
            return "Legacy"
        return "Current"

    # ==================================================================== #
    #  Structure evaluation helpers                                          #
    # ==================================================================== #

    def _make_result(
        self,
        relative_path: str,
        level: str,
        expected_name: str,
        actual_name: str,
        exists: str,
        status: str,
        details: str,
    ) -> dict:
        """Return a standardised result-row dictionary.

        Centralises the key names used throughout the evaluation pipeline so
        that a typo in one place does not cause a silent KeyError elsewhere.

        Parameters
        ----------
        relative_path : Parent folder path (empty string for top-level items).
        level         : "Top level" or "Subfolder".
        expected_name : Name from the template (empty when flagging UNEXPECTED).
        actual_name   : Name found on disk (empty when flagging MISSING).
        exists        : "Yes" or "No".
        status        : One of the STATUS_COLOURS keys.
        details       : Human-readable explanation shown in the Details column.
        """
        return {
            "relative_path": relative_path,
            "level":         level,
            "expected_name": expected_name,
            "actual_name":   actual_name,
            "exists":        exists,
            "status":        status,
            "details":       details,
        }

    def _evaluate_children(
        self,
        top_path: str,
        actual_top_display: str,
        expected_children: list[str],
        actual_children: list[str],
        results: list[dict],
    ) -> None:
        """Compare one level of subfolders against the template's expected list.

        For each expected child the following checks are applied in order:
          - MISSING                if absent from disk
          - NONE - ACCEPTED        if NONE-marked and empty
          - POPULATED DESPITE NONE if NONE-marked but has content
          - EMPTY - REVIEW         if present, not NONE-marked, and empty;
                                   also used when a typed folder has files but
                                   none of the required file type
          - REVIEW - HAND-INS      if a submission folder (sample hand-ins,
                                   sample answers, or scripts) fails the
                                   ≥15-files / same-type check
          - OK (with type detail)  for course handout, DP list, mark sheets,
                                   and external moderator folders that pass
                                   their respective file-type checks
          - OK                     for all other folders with content

        Any disk child not in the expected list is flagged as UNEXPECTED.

        Parameters
        ----------
        top_path           : Absolute path to the parent folder being checked.
        actual_top_display : Display name for the parent (NONE tokens removed).
        expected_children  : Ordered list of expected subfolder names from the profile.
        actual_children    : Sorted list of subfolder names actually on disk.
        results            : Accumulator list; result dicts are appended here.
        """
        # Build a lookup from normalised key → raw disk name for O(1) matching
        actual_child_map    = {self._nbk(n): n for n in actual_children}
        expected_child_keys = set()

        for expected_child in expected_children:
            key = self._nbk(expected_child)
            expected_child_keys.add(key)
            actual_child = actual_child_map.get(key)

            if actual_child is None:
                # Template entry has no matching folder on disk
                results.append(self._make_result(
                    actual_top_display, "Subfolder", expected_child, "",
                    "No", "MISSING", "Expected subfolder is missing",
                ))
                continue

            child_path       = os.path.join(top_path, actual_child)
            has_content      = self.folder_has_content(child_path)
            is_none          = self.has_none_suffix(actual_child)
            has_admin_marker = bool(_STATUS_MARKER_PATTERN.search(actual_child))

            if is_none:
                status = "NONE - ACCEPTED" if not has_content else "POPULATED DESPITE NONE"
                detail = "Subfolder is marked NONE" + (
                    " but contains files or subfolders" if has_content else ""
                )
            elif has_admin_marker:
                status = "ADMIN FLAG"
                detail = (
                    f"Administrator status marker detected in subfolder name — "
                    f"{'subfolder has content' if has_content else 'subfolder is empty'}, please review and rename"
                )
            elif not has_content:
                status, detail = "EMPTY - REVIEW", "Subfolder exists but appears empty"
            elif self.is_submission_folder(expected_child):
                # Runs ≥15-files-of-same-type check for sample hand-ins,
                # sample answers, and scripts folders
                status, detail = self.check_sample_handins(child_path)
            elif self.is_course_handout_folder(expected_child):
                # Must contain at least one PDF or Word document
                status, detail = self.check_course_handout(child_path)
            elif self.is_dp_list_folder(expected_child):
                # Must contain at least one PDF or spreadsheet
                status, detail = self.check_dp_list(child_path)
            elif self.is_mark_sheets_folder(expected_child):
                # Must contain at least one PDF or spreadsheet
                status, detail = self.check_mark_sheets(child_path)
            elif self.is_external_moderator_folder(expected_child):
                # Must contain at least one PDF or Word document
                status, detail = self.check_external_moderator(child_path)
            else:
                status, detail = "OK", "Subfolder found"

            results.append(self._make_result(
                actual_top_display, "Subfolder", expected_child, actual_child,
                "Yes", status, detail,
            ))

        # Flag disk folders that have no matching template entry
        for actual_child in actual_children:
            if self._nbk(actual_child) not in expected_child_keys:
                base = self._nbk(actual_child)
                if base in _ALL_TEMPLATE_BASE_NAMES:
                    detail = (
                        "Subfolder name matches a known template folder but may be "
                        "renumbered or renamed — check against the expected structure"
                    )
                else:
                    detail = (
                        "Subfolder exists but is not in the template and does not match "
                        "any known folder name — check naming convention"
                    )
                results.append(self._make_result(
                    actual_top_display, "Subfolder", "", actual_child,
                    "Yes", "UNEXPECTED", detail,
                ))

    def _check_duplicates(
        self,
        entries: list[str],
        relative_path: str,
        level: str,
        results: list[dict],
    ) -> None:
        """Detect folders at the same level whose normalised base key collides.

        A collision means two or more folders map to the same template slot
        after stripping number prefixes and NONE markers (e.g. "a. Solutions"
        and "b. Solutions" both resolve to "solutions").  Each collision group
        produces one DUPLICATE result row listing all involved folder names.

        Parameters
        ----------
        entries       : List of raw folder names to check (disk names).
        relative_path : Parent path used in the result's relative_path field.
        level         : "Top level" or "Subfolder".
        results       : Accumulator list; DUPLICATE dicts are appended here.
        """
        key_counts = Counter(self._nbk(n) for n in entries)
        for key, count in key_counts.items():
            if count > 1:
                dupes = [n for n in entries if self._nbk(n) == key]
                results.append(self._make_result(
                    relative_path, level, "", ", ".join(dupes),
                    "Yes", "DUPLICATE",
                    f"Multiple folders resolve to the same name: {', '.join(dupes)}",
                ))

    # ==================================================================== #
    #  Updated-profile exam folder evaluation                               #
    # ==================================================================== #

    def _evaluate_updated_exam_folder(
        self, exam_top_path: str, exam_display_name: str, results: list[dict]
    ) -> None:
        """Validate the two-level exam hierarchy specific to the Updated profile.

        The Updated profile uses a single "07. Exams" top-level folder that
        contains two sub-groups — "Main exams" and "SUPP exams" — each of which
        must contain the five subfolders listed in UPDATED_EXAM_SUBFOLDERS.
        This three-level structure is deeper than the standard two-level
        evaluation handled by ``_evaluate_children``, so it requires its own
        dedicated method.

        Expected layout::

            07. Exams
            ├── Main Exams
            │   ├── a. Exam Paper
            │   ├── b. Exam Model Answers
            │   ├── c. External moderator reports
            │   ├── d. Marks Sheets
            │   └── e. Exam Scripts
            └── SUPP exams
                └── (same five subfolders)

        Parameters
        ----------
        exam_top_path    : Absolute path to the "07. Exams" folder on disk.
        exam_display_name: NONE-stripped display name for the exam folder
                           (used as relative_path in result rows).
        results          : Accumulator list; result dicts are appended here.
        """
        expected_exam_groups = UPDATED_STRUCTURE["07. Exams"]   # ["Main exams", "SUPP exams"]

        try:
            actual_children = [
                n for n in sorted(os.listdir(exam_top_path), key=str.lower)
                if os.path.isdir(os.path.join(exam_top_path, n))
            ]
        except OSError:
            actual_children = []

        actual_child_map    = {self._nbk(n): n for n in actual_children}
        expected_child_keys = set()

        for expected_group in expected_exam_groups:
            key = self._nbk(expected_group)
            expected_child_keys.add(key)
            actual_group = actual_child_map.get(key)

            if actual_group is None:
                results.append(self._make_result(
                    exam_display_name, "Subfolder", expected_group, "",
                    "No", "MISSING", "Expected exam group folder is missing",
                ))
                continue

            group_path  = os.path.join(exam_top_path, actual_group)
            has_content = self.folder_has_content(group_path)
            is_none     = self.has_none_suffix(actual_group)

            if is_none:
                # NONE-marked exam group — report and skip sub-subfolder check
                status = "NONE - ACCEPTED" if not has_content else "POPULATED DESPITE NONE"
                detail = "Exam group folder is marked NONE" + (
                    " but contains files" if has_content else ""
                )
                results.append(self._make_result(
                    exam_display_name, "Subfolder", expected_group, actual_group,
                    "Yes", status, detail,
                ))
                continue

            # Report the exam group folder itself
            results.append(self._make_result(
                exam_display_name, "Subfolder", expected_group, actual_group,
                "Yes",
                "OK" if has_content else "EMPTY - REVIEW",
                "Exam group folder found" if has_content else "Exam group folder exists but is empty",
            ))

            # Validate the five sub-subfolders (a. Exam Paper, b. Exam Model Answers, etc.)
            group_display = f"{exam_display_name}/{self.strip_none_suffix(actual_group)}"
            try:
                grandchildren = [
                    n for n in sorted(os.listdir(group_path), key=str.lower)
                    if os.path.isdir(os.path.join(group_path, n))
                ]
            except OSError:
                grandchildren = []

            self._evaluate_children(
                group_path, group_display,
                UPDATED_EXAM_SUBFOLDERS, grandchildren, results,
            )
            self._check_duplicates(grandchildren, group_display, "Subfolder", results)

        # Flag any exam-level folders on disk not in the expected group list
        for actual_child in actual_children:
            if self._nbk(actual_child) not in expected_child_keys:
                results.append(self._make_result(
                    exam_display_name, "Subfolder", "", actual_child,
                    "Yes", "UNEXPECTED", "Subfolder exists but is not in the template",
                ))

        # Duplicate check at the exam-group level
        self._check_duplicates(actual_children, exam_display_name, "Subfolder", results)

    # ==================================================================== #
    #  Main structure evaluation                                             #
    # ==================================================================== #

    def evaluate_expected_structure(self, root_path: str, profile_name: str) -> list[dict]:
        """Compare the course folder on disk against the chosen profile template.

        Iterates over every top-level folder defined in the profile, matches it
        to a disk folder (using normalised base keys), and evaluates both the
        top-level entry and its expected subfolders.

        Special cases handled:
          - NONE-marked top-level: all children inherit NONE status; normal
            child evaluation is skipped.
          - Updated profile "07. Exams": delegated to
            ``_evaluate_updated_exam_folder`` for three-level validation.
          - Disk folders with no matching template entry: flagged UNEXPECTED.
          - Duplicate folder names (after normalisation): flagged DUPLICATE.

        Parameters
        ----------
        root_path    : Absolute path to the course root directory.
        profile_name : One of "Current", "Legacy", or "Updated".

        Returns
        -------
        list[dict]
            One dict per template entry and per unexpected/duplicate finding,
            each conforming to the schema defined by ``_make_result``.
        """
        expected_structure = STRUCTURE_PROFILES[profile_name]
        results: list[dict] = []

        try:
            top_level_entries = [
                name for name in sorted(os.listdir(root_path), key=str.lower)
                if os.path.isdir(os.path.join(root_path, name))
            ]
        except OSError:
            return results   # Root is unreadable — return empty results gracefully

        # Build a lookup from normalised key → raw disk name for top-level folders
        top_level_map     = {self._nbk(n): n for n in top_level_entries}
        expected_top_keys = set()

        for expected_top, expected_children in expected_structure.items():
            expected_key = self._nbk(expected_top)
            expected_top_keys.add(expected_key)
            actual_top   = top_level_map.get(expected_key)

            # ── Top-level folder missing from disk ────────────────────────
            if actual_top is None:
                results.append(self._make_result(
                    "", "Top level", expected_top, "",
                    "No", "MISSING", "Expected top-level folder is missing",
                ))
                continue

            # ── Evaluate the top-level folder itself ──────────────────────
            top_path    = os.path.join(root_path, actual_top)
            is_none_top = self.has_none_suffix(actual_top)
            has_content = self.folder_has_content(top_path)
            has_admin_marker = bool(_STATUS_MARKER_PATTERN.search(actual_top))

            if is_none_top:
                status = "NONE - ACCEPTED" if not has_content else "POPULATED DESPITE NONE"
                detail = "Folder is marked NONE" + (
                    " but contains files or subfolders" if has_content else ""
                )
            elif has_admin_marker:
                status = "ADMIN FLAG"
                detail = (
                    f"Administrator status marker detected in folder name — "
                    f"{'folder has content' if has_content else 'folder is empty'}, please review and rename"
                )
            elif has_content:
                status, detail = "OK", "Top-level folder found"
            else:
                status, detail = "EMPTY - REVIEW", "Top-level folder exists but contains no files"

            results.append(self._make_result(
                "", "Top level", expected_top, actual_top, "Yes", status, detail,
            ))

            # ── NONE-marked top-level: report children without template ───
            # Skip normal child validation; surface each child with an
            # appropriate NONE status so auditors can see what is on disk.
            if is_none_top:
                try:
                    for child_name in sorted(os.listdir(top_path), key=str.lower):
                        if os.path.isdir(os.path.join(top_path, child_name)):
                            child_has = self.folder_has_content(
                                os.path.join(top_path, child_name)
                            )
                            results.append(self._make_result(
                                self.strip_none_suffix(actual_top), "Subfolder",
                                "", child_name, "Yes",
                                "NONE - ACCEPTED" if not child_has else "POPULATED DESPITE NONE",
                                "Child of a NONE-marked folder",
                            ))
                except OSError:
                    pass
                continue   # No further template matching for NONE top-level folders

            # ── Updated profile: 07. Exams requires three-level evaluation ─
            if profile_name == "Updated" and self._nbk(expected_top) == self._nbk("07. Exams"):
                self._evaluate_updated_exam_folder(
                    top_path, self.strip_none_suffix(actual_top), results
                )
                continue

            # ── Standard two-level child evaluation ───────────────────────
            # If files sit directly inside this folder (no subfolders created),
            # treat it as OK rather than flagging all expected subfolders MISSING.
            if expected_children and self.folder_has_direct_files(top_path):
                # Update the top-level row status to reflect flat file layout
                results[-1]["status"] = "OK"
                results[-1]["details"] = "Top-level folder contains files directly (no subfolders)"
                continue

            try:
                actual_children = [
                    n for n in sorted(os.listdir(top_path), key=str.lower)
                    if os.path.isdir(os.path.join(top_path, n))
                ]
            except OSError:
                actual_children = []

            self._evaluate_children(
                top_path, self.strip_none_suffix(actual_top),
                expected_children, actual_children, results,
            )
            self._check_duplicates(
                actual_children, self.strip_none_suffix(actual_top), "Subfolder", results,
            )

        # ── Flag unexpected top-level folders ─────────────────────────────
        for actual_top in top_level_entries:
            if self._nbk(actual_top) not in expected_top_keys:
                base = self._nbk(actual_top)
                if base in _ALL_TEMPLATE_BASE_NAMES:
                    detail = (
                        "Folder name matches a known template folder but may be "
                        "renumbered or renamed — check against the expected structure"
                    )
                else:
                    detail = (
                        "Folder exists but is not in the template and does not match "
                        "any known folder name — check naming convention"
                    )
                results.append(self._make_result(
                    "", "Top level", "", actual_top,
                    "Yes", "UNEXPECTED", detail,
                ))

        # ── Duplicate check at the top level ──────────────────────────────
        self._check_duplicates(top_level_entries, "", "Top level", results)
        return results

    # ==================================================================== #
    #  Folder status derivation                                              #
    # ==================================================================== #

    def folder_status_from_expected(
        self,
        relative_path: str,
        expected_results: list[dict],
        folder_abs_path: str = "",
    ) -> str:
        """Derive the display status for a row in the Folder Details tab.

        Looks up the folder in *expected_results* by matching its name and
        parent path, then returns the most severe status found.  If no
        matching row exists (e.g. a deeply nested submission sub-folder not
        explicitly covered by the template), the folder is evaluated directly
        via content checks and, where applicable, the submission validation rule.

        Status priority order (most → least severe)
        ---------------------------------------------
        DUPLICATE > POPULATED DESPITE NONE > REVIEW - HAND-INS >
        MISSING CHILDREN > EMPTY - REVIEW > NONE - ACCEPTED > UNEXPECTED > OK

        Parameters
        ----------
        relative_path    : Forward-slash path relative to the course root.
        expected_results : Full list of result dicts from evaluate_expected_structure.
        folder_abs_path  : Absolute path used for direct content checks on deep
                           subfolders not covered by the template.

        Returns
        -------
        str
            One of the STATUS_COLOURS keys.
        """
        folder_name = os.path.basename(relative_path)
        parent_path = "/".join(relative_path.replace("\\", "/").split("/")[:-1])

        # Match rows that correspond to this exact folder.  Three conditions
        # cover top-level, direct subfolder, and deep subfolder cases.
        relevant = [
            row for row in expected_results
            if row["actual_name"] == folder_name and (
                (row["relative_path"] == "" and parent_path == "")
                or (row["relative_path"] != "" and row["relative_path"] == parent_path)
                or (row["relative_path"] != "" and
                    relative_path.startswith(row["relative_path"] + "/"))
            )
        ]
        statuses = {row["status"] for row in relevant}

        # Return the highest-priority status present
        priority = [
            "DUPLICATE",
            "ADMIN FLAG",
            "POPULATED DESPITE NONE",
            "REVIEW - HAND-INS",
            "MISSING",           # Translated to "MISSING CHILDREN" for Folder Details
            "EMPTY - REVIEW",
            "NONE - ACCEPTED",
            "UNEXPECTED",
        ]
        for p in priority:
            if p in statuses:
                return "MISSING CHILDREN" if p == "MISSING" else p
        if statuses == {"OK"}:
            return "OK"

        # ── Deep subfolder not in expected_results ─────────────────────────
        # Evaluate directly; NONE applies to the folder itself only and is
        # never inherited from ancestors.
        if folder_abs_path:
            if self.has_none_suffix(folder_name):
                return (
                    "NONE - ACCEPTED"
                    if not self.folder_has_content(folder_abs_path)
                    else "POPULATED DESPITE NONE"
                )
            if not self.folder_has_content(folder_abs_path):
                return "EMPTY - REVIEW"
            if self._is_under_submission_folder(relative_path):
                # Apply the submission validation rule to any folder nested
                # within a sample hand-ins, sample answers, or scripts ancestor
                status, _ = self.check_sample_handins(folder_abs_path)
                return status
            return "OK"

        return "OK"   # No absolute path supplied — cannot evaluate; assume OK

    # ==================================================================== #
    #  Core analysis                                                         #
    # ==================================================================== #

    def analyse_folder_tree(self, root_path: str) -> dict:
        """Walk the entire course folder tree and collect all audit data.

        Performs structure validation followed by a single os.walk traversal
        that simultaneously collects:
          - Folder metrics  (depth, file count, size, mtime, type breakdown, status)
          - File details    (name, extension, size, mtime per file)

        The course root itself is excluded from all folder/file rows.  Each
        directory visited by os.walk is counted exactly once — incrementing
        ``len(dirnames)`` at each step would double-count every sub-directory.

        Parameters
        ----------
        root_path : Absolute path to the course root directory.

        Returns
        -------
        dict
            Keys: profile_name, total_files, total_folders, total_size_bytes,
                  folder_data, file_details, overall_file_type_counts,
                  ascii_tree, expected_results, issues.
        """
        total_files      = 0
        total_folders    = 0   # Incremented once per visited dirpath (not via dirnames)
        total_size_bytes = 0
        folder_data:  list[dict] = []
        file_details: list[dict] = []
        overall_file_types       = Counter()

        # Run structure validation before the walk so status lookup is ready
        detected_profile = self.detect_profile(root_path)
        expected_results = self.evaluate_expected_structure(root_path, detected_profile)

        for dirpath, dirnames, filenames in os.walk(root_path):
            # Sort in-place so os.walk descends in a consistent alphabetical order
            dirnames.sort()
            filenames.sort()

            # Prune ignored directories so os.walk does not descend into them
            dirnames[:] = [d for d in dirnames if d.lower() not in IGNORED_FILES]
            filtered_filenames = [f for f in filenames if f.lower() not in IGNORED_FILES]

            if dirpath == root_path:
                continue   # Skip the root — report only its descendants

            total_folders += 1   # Each visited dirpath is exactly one folder

            relative_directory     = self.get_relative_directory(root_path, dirpath)
            folder_file_total_size = 0
            file_type_counter      = Counter()

            # ── Per-file data collection ───────────────────────────────────
            for filename in filtered_filenames:
                total_files   += 1
                file_extension = self.get_file_extension(filename)
                full_path      = os.path.join(dirpath, filename)

                try:
                    file_size = os.path.getsize(full_path)
                except OSError:
                    file_size = None   # Stat failed; size shown as "Unknown"

                try:
                    modified_time = datetime.fromtimestamp(
                        os.path.getmtime(full_path)
                    ).strftime("%Y-%m-%d %H:%M:%S")
                except OSError:
                    modified_time = "Unavailable"

                if file_size is not None:
                    total_size_bytes       += file_size
                    folder_file_total_size += file_size

                file_type_counter[file_extension]  += 1
                overall_file_types[file_extension] += 1

                file_details.append({
                    "directory": relative_directory,
                    "name":      filename,
                    "extension": file_extension,
                    "size":      file_size,
                    "modified":  modified_time,
                })

                # ── Unsigned file check ────────────────────────────────────
                stem = os.path.splitext(filename)[0]
                if _UNSIGNED_FILE_PATTERN.search(stem):
                    unsigned_file_issues.append(self._make_result(
                        relative_directory, "File",
                        "", filename,
                        "Yes", "UNSIGNED FILE",
                        "File name indicates it has not been signed — please review",
                    ))

            # ── Per-folder summary row ─────────────────────────────────────
            folder_data.append({
                "directory":              relative_directory,
                "depth":                  self.get_depth(relative_directory),
                "subfolder_count":        len(dirnames),
                "file_count":             len(filtered_filenames),
                "folder_file_total_size": folder_file_total_size,
                "latest_modified":        self.latest_modified_in_folder(dirpath),
                "file_type_counts":       dict(sorted(file_type_counter.items())),
                "status":                 self.folder_status_from_expected(
                                              relative_directory, expected_results, dirpath
                                          ),
            })

        # Extract only the rows that warrant auditor attention
        issue_statuses = {
            "MISSING", "EMPTY - REVIEW", "UNEXPECTED",
            "POPULATED DESPITE NONE", "REVIEW - HAND-INS", "DUPLICATE",
            "ADMIN FLAG",
        }
        issue_rows = [row for row in expected_results if row["status"] in issue_statuses]

        return {
            "profile_name":             detected_profile,
            "total_files":              total_files,
            "total_folders":            total_folders,
            "total_size_bytes":         total_size_bytes,
            "folder_data":              folder_data,
            "file_details":             file_details,
            "overall_file_type_counts": dict(sorted(overall_file_types.items())),
            "ascii_tree":               self.build_ascii_tree(root_path),
            "expected_results":         expected_results,
            "issues":                   issue_rows,
        }

    # ==================================================================== #
    #  Populate GUI tables                                                   #
    # ==================================================================== #

    def _populate_structure_treeview(self, tv: ttk.Treeview, rows: list[dict]) -> None:
        """Clear *tv* and insert *rows* into it.

        The column order matches the schema returned by ``_make_result``:
        relative_path, level, expected_name, actual_name, exists, status, details.

        Parameters
        ----------
        tv   : Treeview to refresh (Issues or Expected Structure Check).
        rows : List of result dicts conforming to the _make_result schema.
        """
        for item in tv.get_children():
            tv.delete(item)
        for row in rows:
            tv.insert("", tk.END, values=(
                row["relative_path"], row["level"], row["expected_name"],
                row["actual_name"],   row["exists"], row["status"], row["details"],
            ))

    def populate_issues_table(self, analysis_data: dict) -> None:
        """Refresh the Issues tab Treeview with actionable-only result rows."""
        self._populate_structure_treeview(self.issues_treeview, analysis_data["issues"])

    def populate_expected_table(self, analysis_data: dict) -> None:
        """Refresh the Expected Structure Check tab Treeview with all result rows."""
        self._populate_structure_treeview(
            self.expected_treeview, analysis_data["expected_results"])

    def populate_folder_table(self, analysis_data: dict) -> None:
        """Refresh the Folder Details tab Treeview with per-folder metric rows."""
        for item in self.folder_treeview.get_children():
            self.folder_treeview.delete(item)
        for folder in analysis_data["folder_data"]:
            type_counts_text = (
                ", ".join(f"{k}: {v}" for k, v in folder["file_type_counts"].items())
                if folder["file_type_counts"] else "None"
            )
            self.folder_treeview.insert("", tk.END, values=(
                folder["directory"],
                folder["depth"],
                folder["subfolder_count"],
                folder["file_count"],
                self.format_file_size(folder["folder_file_total_size"]),
                folder["latest_modified"],
                type_counts_text,
                folder["status"],
            ))

    def populate_file_table(self, analysis_data: dict) -> None:
        """Refresh the File Details tab Treeview with one row per file found."""
        for item in self.file_treeview.get_children():
            self.file_treeview.delete(item)
        for file_info in analysis_data["file_details"]:
            self.file_treeview.insert("", tk.END, values=(
                file_info["directory"],
                file_info["name"],
                file_info["extension"],
                self.format_file_size(file_info["size"]),
                file_info["modified"],
            ))

    def populate_tree_tab(self, analysis_data: dict) -> None:
        """Replace the Tree Diagram tab content with the ASCII tree string."""
        self.tree_text.delete("1.0", tk.END)
        self.tree_text.insert(tk.END, analysis_data["ascii_tree"])

    # ==================================================================== #
    #  File output                                                           #
    # ==================================================================== #

    def _output_stem(self, base_path: str) -> str:
        """Return the shared filename stem for both output files.

        Format: YYYYMMDD_<course_code>  (lowercased)

        The course code is extracted using a regex matching UCT convention:
        2+ letters, 3+ digits, optional trailing letter (EEE3097S, CSC2001F).
        If no code is found, the full folder name is sanitised and used instead.

        Computing the stem once and passing it to both filename generators
        ensures both files carry an identical timestamp even when the audit
        runs near a clock-tick boundary.

        Parameters
        ----------
        base_path : Absolute path to the course root directory.
        """
        date_str    = datetime.now().strftime("%Y%m%d")
        folder_name = os.path.basename(base_path).strip()
        match       = re.search(r'[A-Za-z]{2,}[0-9]{3,}[A-Za-z]?', folder_name)
        course_code = (
            match.group(0)
            if match
            else re.sub(r'[^\w\-.]', '_', folder_name).strip('_')
        )
        return f"{date_str}_{course_code}".lower()

    def generate_log_filename(self, base_path: str, stem: str) -> str:
        """Return the absolute path for the plain-text audit log file."""
        return os.path.join(base_path, f"{stem}_folder_audit.txt")

    def generate_workbook_filename(self, base_path: str, stem: str) -> str:
        """Return the absolute path for the Excel audit workbook."""
        return os.path.join(base_path, f"{stem}_folder_audit.xlsx")

    def write_log_file(self, log_path: str, analysis_data: dict, root_path: str) -> None:
        """Write a plain-text audit log to *log_path*.

        The log contains six sections:
          1. Scan details            — root path, auditor, profile, timestamp
          2. Summary                 — counts and issue total
          3. Tree structure diagram  — ASCII tree of the full folder hierarchy
          4. Expected structure check — one line per template entry
          5. Folder details          — per-folder metrics
          6. File details            — per-file metadata

        Parameters
        ----------
        log_path      : Destination file path.
        analysis_data : Dict returned by ``analyse_folder_tree``.
        root_path     : Course root path (included in the log header).
        """
        with open(log_path, "w", encoding="utf-8") as log_file:
            log_file.write("COURSE FOLDER AUDIT LOG\n")
            log_file.write("=" * 100 + "\n\n")

            # Section 1 — Scan details
            log_file.write("SCAN DETAILS\n")
            log_file.write("-" * 100 + "\n")
            log_file.write(f"Root directory:   {root_path}\n")
            log_file.write(f"Auditor:          {self.selected_user.get()}\n")
            log_file.write(f"Detected profile: {analysis_data['profile_name']}\n")
            log_file.write(f"Scan date:        {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")

            # Section 2 — Summary
            log_file.write("SUMMARY\n")
            log_file.write("-" * 100 + "\n")
            log_file.write(f"Total folders: {analysis_data['total_folders']}\n")
            log_file.write(f"Total files:   {analysis_data['total_files']}\n")
            log_file.write(f"Total size:    {self.format_file_size(analysis_data['total_size_bytes'])}\n")
            log_file.write(f"Issues found:  {len(analysis_data['issues'])}\n\n")

            # Section 3 — ASCII tree
            log_file.write("TREE STRUCTURE DIAGRAM\n")
            log_file.write("-" * 100 + "\n")
            log_file.write(analysis_data["ascii_tree"])
            log_file.write("\n\n")

            # Section 4 — Expected structure check
            log_file.write("EXPECTED STRUCTURE CHECK\n")
            log_file.write("-" * 100 + "\n")
            for row in analysis_data["expected_results"]:
                log_file.write(
                    f"Parent: {row['relative_path'] or '[root]'} | "
                    f"Level: {row['level']} | "
                    f"Expected: {row['expected_name'] or '-'} | "
                    f"Actual: {row['actual_name'] or '-'} | "
                    f"Exists: {row['exists']} | "
                    f"Status: {row['status']} | "
                    f"Details: {row['details']}\n"
                )

            # Section 5 — Folder details
            log_file.write("\nFOLDER DETAILS\n")
            log_file.write("-" * 100 + "\n")
            for folder in analysis_data["folder_data"]:
                type_counts_text = (
                    ", ".join(f"{k}: {v}" for k, v in folder["file_type_counts"].items())
                    if folder["file_type_counts"] else "None"
                )
                log_file.write(
                    f"Folder: {folder['directory']} | "
                    f"Depth: {folder['depth']} | "
                    f"Subfolders: {folder['subfolder_count']} | "
                    f"Files: {folder['file_count']} | "
                    f"Size: {self.format_file_size(folder['folder_file_total_size'])} | "
                    f"Latest modified: {folder['latest_modified']} | "
                    f"Types: {type_counts_text} | "
                    f"Status: {folder['status']}\n"
                )

            # Section 6 — File details
            log_file.write("\nFILE DETAILS\n")
            log_file.write("-" * 100 + "\n")
            if analysis_data["file_details"]:
                for file_info in analysis_data["file_details"]:
                    log_file.write(
                        f"Directory: {file_info['directory']} | "
                        f"Name: {file_info['name']} | "
                        f"Type: {file_info['extension']} | "
                        f"Size: {self.format_file_size(file_info['size'])} | "
                        f"Modified: {file_info['modified']}\n"
                    )
            else:
                log_file.write("No files found.\n")

    # ==================================================================== #
    #  Excel workbook helpers                                                #
    # ==================================================================== #

    def apply_sheet_style(self, ws) -> None:
        """Apply standard header styling and auto-fit column widths to *ws*.

        Header row treatment:
          - UCT dark-blue fill (#003C69) with white bold text
          - Horizontally and vertically centred
          - Thin bottom border separating header from data
          - Row 1 frozen so headers stay visible while scrolling

        Column widths are auto-sized from content length, capped at 60
        characters to prevent excessively wide columns for long path strings.
        """
        header_fill = PatternFill("solid", fgColor="003C69")
        header_font = Font(color="FFFFFF", bold=True)
        thin_border = Border(bottom=Side(style="thin", color="D9D9D9"))

        for cell in ws[1]:
            cell.fill      = header_fill
            cell.font      = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border    = thin_border

        ws.freeze_panes = "A2"   # Keep header visible while scrolling

        for column_cells in ws.columns:
            max_length = max(
                (len(str(cell.value or "")) for cell in column_cells),
                default=0,
            )
            col_letter = get_column_letter(column_cells[0].column)
            ws.column_dimensions[col_letter].width = min(max_length + 2, 60)

    def apply_status_conditional_formatting(
        self, ws, status_col_letter: str, total_col_count: int
    ) -> None:
        """Add Excel conditional formatting rules to colour the Status column.

        One CellIsRule is added per status value using an exact-match equality
        test.  Only the Status cell itself is coloured — not the entire row —
        to keep the workbook readable.

        Parameters
        ----------
        ws                : Worksheet to format.
        status_col_letter : Column letter of the Status column (e.g. "F", "H").
        total_col_count   : Accepted for API compatibility; currently unused.
        """
        max_row = ws.max_row
        if max_row < 2:
            return   # Nothing to format beyond the header row

        cell_range = f"{status_col_letter}2:{status_col_letter}{max_row}"

        for status_text, (bg_hex, font_hex) in self.STATUS_COLOURS.items():
            fill = PatternFill(start_color=bg_hex, end_color=bg_hex, fill_type="solid")
            font = Font(color=font_hex, bold=True)
            rule = CellIsRule(
                operator="equal",
                formula=[f'"{status_text}"'],
                fill=fill,
                font=font,
            )
            ws.conditional_formatting.add(cell_range, rule)

    def create_audit_workbook(
        self, workbook_path: str, analysis_data: dict, root_path: str
    ) -> None:
        """Build and save the Excel audit workbook to *workbook_path*.

        The workbook contains five sheets:
          1. Course Audit Summary      — key metrics (root, auditor, counts)
          2. Expected Structure Check  — all template vs disk comparison rows
          3. Folder Details            — per-folder metrics with status
          4. File Details              — per-file metadata (no status column)
          5. Exceptions                — actionable issue rows only

        Sheets 2, 3, and 5 include four blank reviewer columns (Reviewer,
        Checked, Comment, Action Needed) for use during coordinated review.
        Conditional formatting is applied to the Status column in sheets 2, 3,
        and 5.

        Parameters
        ----------
        workbook_path : Destination file path (.xlsx).
        analysis_data : Dict returned by ``analyse_folder_tree``.
        root_path     : Course root path (included in the Summary sheet).
        """
        wb = Workbook()

        # ── Sheet 1: Course Audit Summary ──────────────────────────────────
        ws_summary       = wb.active
        ws_summary.title = "Course Audit Summary"
        for row in [
            ["Item",             "Value"],
            ["Root directory",   root_path],
            ["Auditor",          self.selected_user.get()],
            ["Detected profile", analysis_data["profile_name"]],
            ["Scan date",        datetime.now().strftime("%Y-%m-%d %H:%M:%S")],
            ["Total folders",    analysis_data["total_folders"]],
            ["Total files",      analysis_data["total_files"]],
            ["Total size",       self.format_file_size(analysis_data["total_size_bytes"])],
            ["Issues found",     len(analysis_data["issues"])],
        ]:
            ws_summary.append(row)

        # ── Sheet 2: Expected Structure Check  (Status column = F) ─────────
        ws_expected = wb.create_sheet("Expected Structure Check")
        ws_expected.append([
            "Parent Path", "Level", "Expected Name", "Actual Name", "Exists",
            "Status", "Details", "Reviewer", "Checked", "Comment", "Action Needed",
        ])
        for row in analysis_data["expected_results"]:
            ws_expected.append([
                row["relative_path"], row["level"], row["expected_name"],
                row["actual_name"],   row["exists"], row["status"], row["details"],
                "", "", "", "",   # Blank reviewer columns for coordinated review
            ])

        # ── Sheet 3: Folder Details  (Status column = H) ───────────────────
        ws_folders = wb.create_sheet("Folder Details")
        ws_folders.append([
            "Folder", "Depth", "Subfolder Count", "File Count", "Folder File Size",
            "Latest Modified", "File Type Counts", "Status",
            "Reviewer", "Checked", "Comment", "Action Needed",
        ])
        for folder in analysis_data["folder_data"]:
            type_counts_text = (
                ", ".join(f"{k}: {v}" for k, v in folder["file_type_counts"].items())
                if folder["file_type_counts"] else "None"
            )
            ws_folders.append([
                folder["directory"],       folder["depth"],
                folder["subfolder_count"], folder["file_count"],
                self.format_file_size(folder["folder_file_total_size"]),
                folder["latest_modified"], type_counts_text, folder["status"],
                "", "", "", "",   # Blank reviewer columns
            ])

        # ── Sheet 4: File Details  (no Status column) ──────────────────────
        ws_files = wb.create_sheet("File Details")
        ws_files.append(["Directory", "File Name", "Type", "Size", "Modified"])
        for file_info in analysis_data["file_details"]:
            ws_files.append([
                file_info["directory"], file_info["name"], file_info["extension"],
                self.format_file_size(file_info["size"]), file_info["modified"],
            ])

        # ── Sheet 5: Exceptions  (Status column = F, same layout as sheet 2) ─
        ws_issues = wb.create_sheet("Exceptions")
        ws_issues.append([
            "Parent Path", "Level", "Expected Name", "Actual Name", "Exists",
            "Status", "Details", "Reviewer", "Checked", "Comment", "Action Needed",
        ])
        for row in analysis_data["issues"]:
            ws_issues.append([
                row["relative_path"], row["level"], row["expected_name"],
                row["actual_name"],   row["exists"], row["status"], row["details"],
                "", "", "", "",   # Blank reviewer columns
            ])

        # ── Apply base styling to all sheets ───────────────────────────────
        for ws in (ws_summary, ws_expected, ws_folders, ws_files, ws_issues):
            self.apply_sheet_style(ws)

        # ── Apply conditional formatting to Status columns ─────────────────
        self.apply_status_conditional_formatting(ws_expected, "F", 11)  # Sheet 2
        self.apply_status_conditional_formatting(ws_folders,  "H", 12)  # Sheet 3
        self.apply_status_conditional_formatting(ws_issues,   "F", 11)  # Sheet 5

        wb.save(workbook_path)

    # ==================================================================== #
    #  Main scan entry point                                                 #
    # ==================================================================== #

    def scan_and_export(self) -> None:
        """Validate the selected directory and orchestrate the full audit run.

        Sequence of operations:
          1. Validate that a directory has been selected and that it exists.
          2. Update the recent-directories list.
          3. Run ``analyse_folder_tree`` to collect all audit data.
          4. Refresh all six GUI tabs with the results.
          5. Switch to the Issues tab so the auditor sees problems first.
          6. Compute a shared filename stem (single timestamp, consistent names).
          7. Write the plain-text log file.
          8. Write the Excel workbook.
          9. Update the summary label and show a completion dialog.

        All exceptions are caught, logged with a full traceback, and shown to
        the user as an error dialog so the application never crashes silently.
        """
        root_path = self.selected_directory.get().strip()

        # ── Pre-flight validation ──────────────────────────────────────────
        if not root_path:
            messagebox.showwarning(
                "No Directory Selected", "Please select a directory first."
            )
            return
        if not os.path.isdir(root_path):
            messagebox.showerror(
                "Invalid Directory", "The selected path is not a valid directory."
            )
            return

        try:
            self.update_recent_directories(root_path)
            self.log_message("Starting course folder audit...")
            self.log_message(f"Auditor: {self.selected_user.get()}")

            # ── Core audit ─────────────────────────────────────────────────
            analysis_data = self.analyse_folder_tree(root_path)

            self.log_message(f"Using profile: {analysis_data['profile_name']}")
            self.log_message("Audit complete.")
            self.log_message(f"Total folders found: {analysis_data['total_folders']}")
            self.log_message(f"Total files found:   {analysis_data['total_files']}")
            self.log_message(f"Issues found:        {len(analysis_data['issues'])}")

            if analysis_data["overall_file_type_counts"]:
                self.log_message("Overall file types:")
                for extension, count in analysis_data["overall_file_type_counts"].items():
                    self.log_message(f"  {extension}: {count}")

            # ── Refresh GUI tabs ───────────────────────────────────────────
            self.populate_issues_table(analysis_data)
            self.populate_expected_table(analysis_data)
            self.populate_folder_table(analysis_data)
            self.populate_file_table(analysis_data)
            self.populate_tree_tab(analysis_data)
            self.notebook.select(self.issues_tab)   # Surface issues immediately

            # ── Write output files ─────────────────────────────────────────
            # Compute the stem once so both files share the same timestamp
            stem          = self._output_stem(root_path)
            log_path      = self.generate_log_filename(root_path, stem)
            workbook_path = self.generate_workbook_filename(root_path, stem)

            self.write_log_file(log_path, analysis_data, root_path)
            self.create_audit_workbook(workbook_path, analysis_data, root_path)

            self.log_message(f"Log file created:  {log_path}")
            self.log_message(f"Workbook created:  {workbook_path}")

            # ── Update summary label ───────────────────────────────────────
            self.summary_label.config(
                text=(
                    f"Profile: {analysis_data['profile_name']}   "
                    f"Folders: {analysis_data['total_folders']}   "
                    f"Files: {analysis_data['total_files']}   "
                    f"Issues: {len(analysis_data['issues'])}"
                )
            )

            messagebox.showinfo(
                "Success",
                (
                    "Course folder audit complete.\n\n"
                    f"Profile:  {analysis_data['profile_name']}\n"
                    f"Folders:  {analysis_data['total_folders']}\n"
                    f"Files:    {analysis_data['total_files']}\n"
                    f"Issues:   {len(analysis_data['issues'])}\n\n"
                    f"Log saved to:\n{log_path}\n\n"
                    f"Workbook saved to:\n{workbook_path}"
                ),
            )

        except Exception as exc:
            # Log the full traceback for developer diagnostics while showing
            # a concise message in the dialog for the end user.
            error_message = f"An error occurred: {exc}"
            self.log_message(error_message)
            self.log_message(traceback.format_exc())
            messagebox.showerror("Error", error_message)


# ===========================================================================
# Entry point
# ===========================================================================

def main() -> None:
    """Create the Tk root window and start the application event loop."""
    root = tk.Tk()
    CourseFolderAuditApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()