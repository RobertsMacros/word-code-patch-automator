# Word VBA Patch Automator

A local-first automation loop for testing and repairing VBA macro code in Microsoft Word.

Runs inside **Windows** (e.g. Parallels VM on Mac). Uses COM automation to drive real Word, a VBA test harness for ground-truth testing, and generates repair prompts for Claude Code CLI.

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────┐
│                    Controller (Python)                       │
│                                                             │
│  1. Import VBA source (src/) into Word document via COM     │
│  2. Import test harness (harness/) into same document       │
│  3. Run TestHarness.RunAllTests via Application.Run         │
│  4. Harness writes JSON results to results/ directory       │
│  5. Controller reads results, generates summary + report    │
│  6. If tests fail: generate repair_prompt.md                │
│  7. Human runs Claude Code CLI against the prompt           │
│  8. Claude patches only src/ files                          │
│  9. Controller reruns (back to step 1)                      │
│ 10. Stop on: all pass OR max iterations reached             │
└─────────────────────────────────────────────────────────────┘

Key constraints:
- Controller decides when to stop, not Claude
- Tests/fixtures/harness are immutable by default
- Claude may only modify files under src/
- Word COM timeout = 60s default (configurable)
```

## Folder Structure

```
word-code-patch-automator/
├── CLAUDE.md                   # Instructions for Claude (immutable paths, rules)
├── project.json                # Project config (doc path, timeouts, paths)
├── locked_paths.json           # Paths Claude must not modify
├── .gitignore
│
├── src/                        # VBA source files (.bas, .cls, .frm)
│   ├── MyModule.bas            #   ← your project's VBA modules go here
│   └── MyClass.cls
│
├── harness/                    # VBA test harness (imported into Word at test time)
│   ├── TestHarness.bas         #   ← test runner framework
│   └── Test_Smoke.bas          #   ← example smoke tests
│
├── tests/
│   ├── fixtures/               # Test documents (.docm) — immutable
│   │   └── TestDocument.docm
│   └── expected/               # Expected outputs for regression tests
│
├── controller/
│   ├── controller.py           # Main controller script
│   ├── vba_io.py               # Standalone import/export tool
│   └── templates/
│       └── repair_prompt.md    # Repair prompt template
│
└── results/                    # Generated each run (gitignored)
    ├── results_001.json        #   raw harness output
    ├── report_001.json         #   detailed machine-readable report
    ├── summary_001.txt         #   concise human-readable summary
    ├── repair_prompt_001.md    #   repair prompt for Claude
    └── run_log.json            #   combined log of all iterations
```

## Test Layers

| Layer | Purpose | Speed | Where it runs |
|-------|---------|-------|---------------|
| **Smoke** | Harness + environment sanity | <1s | Word COM |
| **Fast logic** | Pure VBA logic, no doc I/O | Fast | Word COM |
| **Regression** | Known bugs on fixture documents | Medium | Word COM on fixture .docm |
| **Integration** | Full document operations | Slower | Word COM on fixture .docm |

All tests run inside real Word via COM — there is no mock layer in the MVP.

## Setup (Windows VM)

### Prerequisites
- Python 3.8+ on Windows
- `pip install pywin32`
- Microsoft Word installed
- Word macro security: **Trust access to the VBA project object model**
  (File → Options → Trust Center → Trust Center Settings → Macro Settings)
- Git

### Quick Start

1. Clone this repo inside the Windows VM
2. Place your VBA source files in `src/`
3. Place a macro-enabled test document in `tests/fixtures/` and update `project.json`
4. Add test modules (named `Test_*.bas`) to `harness/`
5. Run the controller:

```powershell
cd controller
python controller.py
```

6. If tests fail, run Claude Code against the generated prompt:

```powershell
claude -p results\repair_prompt_001.md
```

7. Re-run the controller to test the fix:

```powershell
python controller.py
```

## Import / Export

Standalone utility for importing/exporting VBA modules:

```powershell
# Export from Word doc to src/
python controller\vba_io.py export

# Import from src/ into Word doc
python controller\vba_io.py import

# Export with Word visible (for debugging)
python controller\vba_io.py export --visible

# Full roundtrip (export then re-import)
python controller\vba_io.py roundtrip
```

## Configuration (project.json)

| Key | Description | Default |
|-----|-------------|---------|
| `project_name` | Display name | `"MyVBAProject"` |
| `word_doc` | Path to the test document (relative to repo root) | — |
| `src_dir` | VBA source directory | `"src"` |
| `controller.max_iterations` | Max repair cycles before giving up | `5` |
| `controller.timeout_seconds` | Per-test-run timeout in seconds | `60` |
| `controller.word_visible` | Show Word window during testing | `false` |
| `controller.kill_on_timeout` | Force-kill Word on timeout | `true` |
| `mutable_paths` | Paths Claude is allowed to edit | `["src/"]` |
| `immutable_paths` | Paths Claude must not touch | tests, harness, controller, config |

## Reporting

Each run produces:

- **`report_NNN.json`** — Full machine-readable report with per-test timing, pass/fail, messages
- **`summary_NNN.txt`** — Concise human summary:
  ```
  === Test Run Summary (iteration 1) ===
  Time: 2026-03-13 14:30:00
  Elapsed: 2.34s
  Tests: 5 total, 3 passed, 2 failed

  Failed tests:
    FAIL: Test_Logic.Test_ParseDate (12.3ms) — Expected [2026] but got [2025]
    FAIL: Test_Regression.Test_Bug42 (45.1ms) — Missing paragraph

  Result: FAIL
  ```
- **`repair_prompt_NNN.md`** — Ready-to-use prompt for Claude Code CLI
- **`run_log.json`** — Combined log of all iterations in this run

## Immutable Path Enforcement

Three layers of protection:

1. **CLAUDE.md** — Claude Code reads this and is instructed not to touch locked paths
2. **locked_paths.json** — Machine-readable list of locked paths
3. **Controller check** — `check_locks_before_run()` uses `git diff` to detect violations before each run and aborts if locked files were changed

## Writing Tests

Create a module named `Test_<Category>.bas` in `harness/`:

```vba
Attribute VB_Name = "Test_MyCategory"
Option Explicit

Public Sub Test_MyCategory_SomeBehavior()
    ' Arrange
    Dim result As String
    result = MyModule.DoSomething("input")

    ' Assert
    AssertEqual "expected output", result, "DoSomething should return expected output"
End Sub

Public Sub Test_MyCategory_EdgeCase()
    AssertTrue IsNumeric("123"), "Should recognise numeric string"
    AssertFalse IsNumeric("abc"), "Should reject non-numeric string"
End Sub
```

Available assertions: `AssertTrue`, `AssertFalse`, `AssertEqual`, `AssertNotEqual`, `AssertContains`, `Fail`.

## MVP vs Future

### MVP (this version)
- Semi-automatic: controller stops after generating repair prompt, human runs Claude
- One Word document at a time
- All tests run in real Word via COM
- JSON + text reporting
- Git-based locked path checking
- Standard modules (.bas) and class modules (.cls) import/export

### Future Upgrades
- Fully automatic: controller invokes Claude Code CLI directly
- Multiple fixture documents per test suite
- UserForm (.frm/.frx) test support beyond import/export
- Per-rule timing for document checker macros
- Parallel test execution
- Test tagging and selective test runs
- CI integration (optional)
- Diff-based repair prompts (show Claude only what changed)
