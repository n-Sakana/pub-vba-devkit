# AGENTS.md - pub/vba-devkit

Entry-point notes for Codex. For repo-specific details see [README.md](README.md) and [CLAUDE.md](CLAUDE.md); for the cross-repo topology see [fin/hub/ARCHITECTURE.md](../../fin/hub/ARCHITECTURE.md).

## Role

- A CLI development kit for Excel / VBA assets
- Bundles extraction, diff, analysis, environment testing, sanitization, and unlocking
- A local toolset that assumes Windows + Excel

## runtime / connectivity

- runtime: local Windows only
- shell: PowerShell 5.1 + `*.bat` wrapper
- automation: Excel COM

## Where to look first

- `lib/VBAToolkit.psm1` - shared functions
- `lib/Extract.ps1`, `Diff.ps1`, `Analyze.ps1`, `Sanitize.ps1`, `Unlock.ps1`
- `config/` - config templates
- `test/` - automated tests

## Commands

```cmd
Extract.bat <path.xlsm>
Diff.bat <old.xlsm> <new.xlsm>
Analyze.bat <path.xlsm>
EnvTest.bat
Sanitize.bat <path>
Unlock.bat <path.xlsm>
test\Run-All.bat
```

## Guardrails

- `env-test-results/` may contain confidential information; do not put it in git
- Always release COM objects
- `Unlock.bat` is powerful — do not run it without explicit instruction
- Do not break the Windows-only assumption
