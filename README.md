# Word VBA Patch Automator

Local automation loop for testing and repairing VBA macro code in Microsoft Word.

Runs inside **Windows** (e.g. Parallels VM). Uses COM automation to drive real Word, a VBA test harness for ground-truth testing, and generates repair prompts for Claude Code CLI.

## Architecture

```
controller.py (parent)
  │
  ├── check locked paths (git status --porcelain)
  ├── discover fixtures from tests/fixtures/
  │
  └── spawn word_runner.py (child subprocess)
        │
        ├── open host/MacroHost.docm
        ├── import VBA source from src/
        ├── import harness from harness/
        ├── run TestHarness.RunAllTests
        │     ├── smoke tests (no fixture)
        │     ├── logic tests (no fixture)
        │     └── for each fixture in tests/fixtures/
        │           ├── open fixture read-only
        │           ├── run TestFixture_* subs
        │           ├── record per-fixture results
        │           └── close fixture
        └── write JSON results

  controller reads results
  ├── write report_NNN.json (detailed)
  ├── write summary_NNN.txt (concise)
  └── if failed: write repair_prompt_NNN.md
        └── human runs: claude -p repair_prompt_NNN.md

  if timeout: controller kills child + Word via taskkill
```

The subprocess boundary is what makes timeout enforcement real. If Word hangs inside a VBA macro, the controller can still kill the child process and Word.

## Folder Structure

```
word-code-patch-automator/
├── project.json                # All configuration (single source of truth)
├── CLAUDE.md                   # Instructions for Claude Code
├── .gitignore
│
├── host/                       # Mutable macro container
│   └── MacroHost.docm          #   ← you create this
│
├── src/                        # VBA source (.bas, .cls, .frm)
│   └── MyModule.bas
│
├── harness/                    # VBA test harness
│   ├── TestHarness.bas         #   test runner framework
│   └── Test_Smoke.bas          #   example tests
│
├── tests/
│   ├── fixtures/               # Fixture documents — protected
│   └── expected/               # Expected outputs — protected
│
├── controller/
│   ├── controller.py           # Parent orchestrator
│   ├── word_runner.py          # Child COM worker
│   ├── vba_io.py               # Manual import/export tool
│   └── templates/
│       └── repair_prompt.md
│
└── results/                    # Generated each run (gitignored)
```

## Host Document vs Fixture Documents

| | Host (`host/MacroHost.docm`) | Fixtures (`tests/fixtures/`) |
|--|--|--|
| Purpose | Mutable container for VBA code | Immutable test inputs |
| VBA imported? | Yes — source + harness | No |
| Modified at runtime? | Yes | Opened read-only |
| Protected from Claude? | No | Yes |

## Configuration (project.json)

Single source of truth for all settings:

```json
{
    "project_name": "MyVBAProject",
    "host_doc": "host/MacroHost.docm",
    "src_dir": "src",
    "fixtures_dir": "tests/fixtures",
    "expected_dir": "tests/expected",
    "results_dir": "results",
    "mutable_paths": ["src/", "controller/", "harness/"],
    "locked_paths": ["tests/fixtures/", "tests/expected/", "project.json", "CLAUDE.md"],
    "controller": {
        "max_iterations": 5,
        "timeout_seconds": 60,
        "word_visible": false,
        "kill_on_timeout": true
    },
    "harness": {
        "entry_sub": "RunAllTests"
    }
}
```

`locked_paths` defines what the controller enforces before each run. `mutable_paths` defines what the repair prompt tells Claude it may edit.

## Setup (Windows VM)

1. Python 3.8+ with `pip install pywin32`
2. Microsoft Word with "Trust access to VBA project object model" enabled
3. Git
4. Clone this repo
5. Create `host/MacroHost.docm` — empty macro-enabled document
6. Place VBA source in `src/`, fixtures in `tests/fixtures/`, test modules in `harness/`
7. Run: `python controller\controller.py`

## Usage

```powershell
# Run tests
python controller\controller.py

# If tests fail, run Claude against the repair prompt
claude -p results\repair_prompt_001.md

# Re-run to verify
python controller\controller.py

# Manual import/export
python controller\vba_io.py import
python controller\vba_io.py export
python controller\vba_io.py import --visible
python controller\vba_io.py export --doc path\to\other.docm
```

## Test Layers

| Layer | Naming | Fixture? |
|-------|--------|----------|
| Smoke | Built into TestHarness | No |
| Logic | `Sub Test_Logic_*()` | No |
| Regression | `Sub TestFixture_Regression_*(doc)` | Yes |
| Integration | `Sub TestFixture_Integration_*(doc)` | Yes |

Non-fixture tests run once. Fixture tests run once per fixture document.

## Writing Tests

```vba
Attribute VB_Name = "Test_MyTests"
Option Explicit

' Non-fixture test — runs once
Public Sub Test_MyTests_ParseDate()
    AssertEqual 2026, ParseYear("2026-03-13"), "Should parse year"
End Sub

' Fixture test — runs once per fixture document
Public Sub TestFixture_MyTests_CheckParagraph(doc As Document)
    AssertContains doc.Paragraphs(1).Range.Text, "Expected", "Should contain expected text"
End Sub
```

Assertions: `AssertTrue`, `AssertFalse`, `AssertEqual`, `AssertNotEqual`, `AssertContains`, `Fail`.

## Reporting

Each iteration produces:

**`report_NNN.json`** — detailed, fixture-centric:
```json
{
  "project": "MyVBAProject",
  "iteration": 1,
  "timestamp": "2026-03-13T14:30:00",
  "elapsed_seconds": 2.34,
  "timed_out": false,
  "word_killed": false,
  "fixtures": [
    {"name": "doc1.docm", "status": "pass", "elapsed": 1.2, "errors": []},
    {"name": "doc2.docm", "status": "fail", "elapsed": 0.8, "errors": ["..."]}
  ],
  "tests": [...],
  "summary": {
    "fixtures_total": 2, "fixtures_passed": 1, "fixtures_failed": 1,
    "tests_total": 8, "tests_passed": 6, "tests_failed": 2,
    "timed_out": 0, "all_passed": false
  }
}
```

**`summary_NNN.txt`** — concise:
```
Run summary
-----------
iteration: 1
fixtures:  2
tests:     8
passed:    6
failed:    2
timeouts:  0
elapsed:   2.3s

Failed fixtures:
  - doc2.docm

Failed tests:
  - Test_Logic.Test_ParseDate — Expected [2026] but got [2025]
  - Test_Regression.TestFixture_Bug42 [doc2.docm] — Missing paragraph

result: FAIL
```

Timeout case:
```
Run summary
-----------
iteration: 2
fixtures:  0
tests:     0
passed:    0
failed:    0
timeouts:  1
elapsed:   60.1s
word:      killed
error:     TIMEOUT: test run exceeded 60s limit (Word killed)

result: FAIL
```

## Locked Path Enforcement

The controller runs `git status --porcelain` before each test pass and checks for changes under `locked_paths` (from `project.json`). It catches:
- Modified but unstaged files
- Staged changes
- Untracked files in protected directories

If violations are found, the run is aborted with a clear message.

## .frm / .frx Limitations

- `.frm` files import/export via COM
- `.frx` binary companions must travel alongside `.frm` in `src/`
- UserForm functional testing is not covered in the MVP

## MVP vs Future

**MVP (this version):** Semi-automatic repair flow, real timeout via subprocess, working-tree locked-path enforcement, fixture-based test loop with per-fixture reporting, flat `src/` layout, dual-format reporting.

**Future:** Fully automatic (controller invokes Claude CLI), per-rule timing, expected output diffing, UserForm testing, test tagging, parallel execution.
