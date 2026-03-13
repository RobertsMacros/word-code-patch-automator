# Word VBA Patch Automator

A local-first automation loop for testing and repairing VBA macro code in Microsoft Word.

Runs inside **Windows** (e.g. Parallels VM on Mac). Uses COM automation to drive real Word, a VBA test harness for ground-truth testing, and generates repair prompts for Claude Code CLI.

## Architecture

```
┌──────────────────────────────────────────────────────────────────┐
│  Controller (Python) — parent process                            │
│                                                                  │
│  1. Check locked paths (working tree + staged)                   │
│  2. Discover fixture documents from tests/fixtures/              │
│  3. Spawn word_runner.py as a child process                      │
│  4. Monitor with real timeout — kill Word if it hangs            │
│  5. Read JSON results from the harness                           │
│  6. Generate detailed report + concise summary                   │
│  7. If tests fail: generate repair_prompt.md                     │
│  8. Human runs Claude Code CLI against the prompt                │
│  9. Claude patches only src/ files                               │
│ 10. Re-run controller (back to step 1)                           │
│ 11. Stop on: all pass OR max iterations reached                  │
│                                                                  │
│  word_runner.py — child process (subprocess boundary)            │
│  ├── Opens host/MacroHost.docm                                   │
│  ├── Imports VBA source (src/) + harness modules                 │
│  ├── Runs TestHarness.RunAllTests via COM                        │
│  └── If Word hangs, controller kills this process + Word         │
└──────────────────────────────────────────────────────────────────┘

Key points:
- Subprocess boundary enables real timeout enforcement
- host/MacroHost.docm = mutable macro container (NOT a test fixture)
- tests/fixtures/*.docm = immutable test documents
- Controller decides when to stop, not Claude
- Claude may only modify files under src/
```

## Folder Structure

```
word-code-patch-automator/
├── CLAUDE.md                   # Instructions for Claude (immutable paths, rules)
├── project.json                # Project config
├── locked_paths.json           # Paths Claude must not modify
├── .gitignore
│
├── host/                       # Mutable macro container
│   └── MacroHost.docm          #   ← create this yourself
│
├── src/                        # VBA source files (.bas, .cls, .frm)
│   └── MyModule.bas            #   ← your VBA modules go here
│
├── harness/                    # VBA test harness (imported into host doc)
│   ├── TestHarness.bas         #   ← test runner framework
│   └── Test_Smoke.bas          #   ← smoke tests
│
├── tests/
│   ├── fixtures/               # Fixture documents (.docm) — immutable
│   │   └── SampleDoc.docm
│   └── expected/               # Expected outputs for regression tests
│
├── controller/
│   ├── controller.py           # Main controller (parent orchestrator)
│   ├── word_runner.py          # COM worker (child process)
│   ├── vba_io.py               # Standalone import/export tool
│   └── templates/
│       └── repair_prompt.md    # Repair prompt template
│
└── results/                    # Generated each run (gitignored)
    ├── report_001.json         #   detailed machine-readable report
    ├── summary_001.txt         #   concise human-readable summary
    ├── repair_prompt_001.md    #   repair prompt for Claude
    └── run_log.json            #   combined iteration log
```

### Host Document vs Fixture Documents

| | Host (`host/MacroHost.docm`) | Fixtures (`tests/fixtures/`) |
|--|--|--|
| **Purpose** | Mutable container for VBA code | Immutable test documents |
| **VBA imported?** | Yes — source + harness | No |
| **Modified at runtime?** | Yes | Opened read-only |
| **Claude may edit?** | No (locked) | No (locked) |
| **You create it** | Once, manually | Per test scenario |

## Test Layers

| Layer | Naming Pattern | Fixture Needed? | Speed |
|-------|---------------|-----------------|-------|
| **Smoke** | Built into harness | No | <1s |
| **Fast logic** | `Sub Test_Logic_*()` | No | Fast |
| **Regression** | `Sub TestFixture_Regression_*(doc)` | Yes | Medium |
| **Integration** | `Sub TestFixture_Integration_*(doc)` | Yes | Slower |

- **Non-fixture tests** (`Test_*`): Run once, no document argument
- **Fixture tests** (`TestFixture_*`): Run once per fixture document, receive `Document` parameter

## Setup (Windows VM)

### Prerequisites
- Python 3.8+
- `pip install pywin32`
- Microsoft Word
- Trust access to VBA project object model enabled in Word
  (File > Options > Trust Center > Trust Center Settings > Macro Settings)
- Git

### Quick Start

1. Clone this repo inside the Windows VM
2. Create `host/MacroHost.docm` — an empty macro-enabled document
3. Place your VBA source files (`.bas`, `.cls`) in `src/`
4. Place fixture documents in `tests/fixtures/`
5. Add test modules (`Test_*.bas`) to `harness/`
6. Run:

```powershell
python controller\controller.py
```

7. If tests fail, the controller generates a repair prompt:

```powershell
claude -p results\repair_prompt_001.md
```

8. Re-run the controller to verify:

```powershell
python controller\controller.py
```

## Import / Export

```powershell
# Import src/ modules into host document
python controller\vba_io.py import

# Export from host document to src/
python controller\vba_io.py export

# With Word visible
python controller\vba_io.py import --visible

# Use a different document
python controller\vba_io.py export --doc path/to/other.docm
```

## Configuration (project.json)

```json
{
    "project_name": "MyVBAProject",
    "host_doc": "host/MacroHost.docm",
    "src_dir": "src",
    "fixtures_dir": "tests/fixtures",
    "expected_dir": "tests/expected",
    "results_dir": "results",
    "mutable_paths": ["src/"],
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

## Timeout Enforcement

The controller uses a **subprocess boundary** for real timeout enforcement:

1. `controller.py` spawns `word_runner.py` as a child process
2. `word_runner.py` does all COM work (open Word, import VBA, run harness)
3. Controller monitors with `subprocess.communicate(timeout=N)`
4. If timeout expires: kills the child process, then `taskkill /F /IM WINWORD.EXE`
5. Reports the timeout clearly in both JSON and summary output

This works even if Word is completely hung inside a VBA macro.

## Reporting

Each iteration produces two files:

**`report_NNN.json`** — detailed machine-readable report:
- Per-test name, pass/fail, message, duration, fixture
- Per-fixture pass/fail counts and timing
- Timeout and Word-killed flags
- Full harness output

**`summary_NNN.txt`** — concise human-readable:
```
=== Test Run Summary (iteration 1) ===
Time:    2026-03-13 14:30:00
Elapsed: 2.34s
Fixtures: TestDoc1.docm, TestDoc2.docm
Tests: 8 total, 6 passed, 2 failed

Failed tests:
  FAIL: Test_Logic.Test_ParseDate (12.3ms) — Expected [2026] but got [2025]
  FAIL: Test_Regression.TestFixture_Bug42 [TestDoc1.docm] (45.1ms) — Missing paragraph

Result: FAIL
================================================
```

Timeout/crash cases are reported clearly:
```
=== Test Run Summary (iteration 2) ===
Time:    2026-03-13 14:32:00
Elapsed: 60.12s
TIMEOUT: run exceeded limit
Word was force-killed
ERROR: TIMEOUT: test run exceeded 60s limit (Word killed)
Tests: 0 total, 0 passed, 0 failed

Result: FAIL
================================================
```

## Locked Path Enforcement

Three layers:

1. **CLAUDE.md** — Claude Code reads this and is instructed not to touch locked paths
2. **locked_paths.json** — Machine-readable list
3. **Controller check** — runs `git status --porcelain` before each run to catch:
   - Modified but unstaged files in locked paths
   - Staged changes to locked files
   - Untracked files in locked directories

If violations are found, the controller prints them and aborts.

## Writing Tests

### Non-fixture tests (logic, smoke)

```vba
Attribute VB_Name = "Test_Logic"
Option Explicit

Public Sub Test_Logic_ParseDate()
    AssertEqual 2026, ParseYear("2026-03-13"), "Should parse year"
End Sub

Public Sub Test_Logic_ValidateInput()
    AssertTrue IsValidInput("hello"), "Should accept valid input"
    AssertFalse IsValidInput(""), "Should reject empty input"
End Sub
```

### Fixture-based tests (regression, integration)

```vba
Attribute VB_Name = "Test_Regression"
Option Explicit

Public Sub TestFixture_Regression_Bug42(doc As Document)
    ' doc is a fixture document opened read-only by the harness
    Dim para As Paragraph
    Set para = doc.Paragraphs(1)
    AssertContains para.Range.Text, "Expected text", "First paragraph should contain expected text"
End Sub

Public Sub TestFixture_Integration_FormatCheck(doc As Document)
    ' Run the macro against the fixture
    doc.Range.Select
    Application.Run "MyFormatter"
    AssertEqual "Formatted", doc.Paragraphs(1).Style.NameLocal, "Should apply Formatted style"
End Sub
```

Assertions: `AssertTrue`, `AssertFalse`, `AssertEqual`, `AssertNotEqual`, `AssertContains`, `Fail`.

## .frm / .frx Limitations

- `.frm` (UserForm) files can be imported/exported via COM
- `.frx` (binary companion files for forms with images/controls) must travel alongside the `.frm`
- UserForm testing beyond import/export is not covered in the MVP
- If a form has embedded images or ActiveX controls, the `.frx` must be present in `src/` next to the `.frm`

## MVP vs Future

### MVP (this version)
- Semi-automatic: controller generates repair prompt, human runs Claude
- Real timeout enforcement via subprocess boundary
- Working-tree locked-path checking (modified, staged, untracked)
- Host document / fixture document separation
- Fixture-based test loop with per-fixture reporting
- Flat `src/` directory (all .bas/.cls/.frm in one folder)
- JSON + text dual-format reporting
- Timeout/crash/killed-Word clearly reported

### Future Upgrades
- Fully automatic mode (controller invokes Claude CLI directly)
- Per-rule timing for document checker macros
- Expected output comparison for regression tests
- UserForm functional testing
- Test tagging and selective runs
- Parallel test execution
- CI integration (optional)
