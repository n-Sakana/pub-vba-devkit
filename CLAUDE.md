# CLAUDE.md - pub/vba-devkit

## Project Overview

A command-line development kit for Excel / VBA assets.
Provides a PowerShell-driven workflow for extracting VBA source from `.xlsm`,
diffing, environment-info testing, sanitizing, and unlocking protection.
The VBA work platform for the teacher (dummy-org).

## Tech Stack

- PowerShell 5.1 (the version bundled with Windows) + `.ps1` / `.bat` wrappers
- Excel automation: COM (Excel.Application)
- Distribution: a directory copy is enough; works via `git clone`

## Directory Layout

```
vba-devkit/
├── *.bat                # Entry points (Analyze / Diff / EnvTest / Extract / Sanitize / Unlock)
├── lib/
│   ├── VBAToolkit.psm1  # Shared module
│   ├── {Analyze|Diff|EnvTest|Extract|Sanitize|Unlock}.ps1
│   └── internal/
├── config/              # Per-command config templates
├── demos/               # Samples for sanity checks
├── samples/             # Typical input examples
├── test/                # Automated tests
├── docs/                # Detailed usage
└── env-test-results/    # Output of EnvTest.bat (**.gitignore target**)
```

## Commands

```cmd
Extract.bat <path.xlsm>      # Extract VBA source → .bas / .cls / .frm
Diff.bat <old.xlsm> <new>    # VBA diff between two workbooks
Analyze.bat <path.xlsm>      # Analyze VBA structure, imports, call graph
EnvTest.bat                  # Collect PC-specific info (tenant name, paths, registry)
Sanitize.bat <path>          # Sanitize (replace tenant names with ***)
Unlock.bat <path.xlsm>       # Unlock VBA project protection
```

## Design Principles

- **Sanitization handling**: output of `EnvTest.bat` and the `env-test-results/` directory may contain environment-dependent strings such as tenant names
  - Reliably excluded via `.gitignore`
  - Run `Sanitize.bat` before commit, or manually verify the `Mask-Path` logic
  - Fixed an `OneDrive - {tenant}` style leak on 2026-04-18 (commit `e4c17ee`)
- **Windows only**: assumes a PC with bundled PowerShell 5.1 + Excel installed
- **Avoid COM leaks**: always Quit Excel.Application after use and follow up with `[GC]::Collect()`-equivalent cleanup
- **No build step**: `.psm1` / `.ps1` are executed directly

## Tests

```cmd
test\Run-All.bat           # All tests
test\Test-Extract.ps1      # Individual
```

## Notes

- VBA assets live entirely inside `.xlsm`; the workflow extracts them into git-friendly forms (.bas/.cls) for management
- To prevent forgetting sanitization: a git pre-commit hook that blocks staging of `env-test-results/` is recommended
- When letting an AI operate, `Unlock.bat` is powerful — wait for explicit instruction from the teacher

## Related

- pub/casedesk — the main consumer of this devkit (a VBA add-in for case management)
- pub/watchbox — not VBA but PS+C#; same distribution vibe (double-click launch via launch.bat)
