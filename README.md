# vba-devkit

Excel VBA migration toolkit. Binary-level analysis without opening Excel.

## Tools

### Main

| BAT | Description |
|-----|-------------|
| `EnvTest.bat` | Unified environment test launcher (Survey / Probe / Full) |
| `Extract.bat` | Extract VBA source code (individual modules + combined.txt) |
| `Analyze.bat` | Analysis + sanitization + migration guide + CSV |
| `Diff.bat` | Side-by-side VBA code comparison |

### Environment Test

| Mode | Description |
|------|-------------|
| `EnvTest.bat` -> `S` | Survey only |
| `EnvTest.bat` -> `B` | Probe Basic |
| `EnvTest.bat` -> `E` | Probe Basic + Extended |
| `EnvTest.bat` -> `F` | Survey + Probe Basic |
| `EnvTest.bat` -> `X` | Survey + Probe Basic + Extended |
| `EnvTest.bat` -> `G` | Generate `probe_storage.xlsm` for SharePoint/OneDrive comparison |

`Survey` prints live section/item results while running. `EnvTest.bat` is the single entry point when you want to choose between Survey, Probe, or a combined run.

For SharePoint/OneDrive migration, the key decision is now explicit:

- If `OneDrive Environment` and `Local Sync Root*` tests pass, shared VBA code can likely use `Environ$("OneDriveCommercial") + relative path`.
- If those VBA-side tests fail but host-side `Get-ChildItem` / `.NET Directory.GetFiles` pass, use a casedesk-style host/service for storage enumeration.
- If Win32 P/Invoke fails, check `UIAutomation` before falling back to `WScript.Shell` / `SendKeys`.

See also: `docs/storage-path-strategy.md`

### Auxiliary

| BAT | Description |
|-----|-------------|
| `Unlock.bat` | Remove VBA project password protection (non-destructive) |

## Analyze

Analyze is the core tool. It detects risks across 4 categories:

| Category | What it finds | Highlight |
|----------|---------------|-----------|
| **EDR** | Win32 API, Shell, COM, WMI, DLL loading | Blue |
| **Compatibility** | PtrSafe, DAO, legacy controls, DDE | Purple |
| **Environment** (Risk/Review/Info) | Path resolution, SaveAs, Dir(), external links, AutoSave events | Green |
| **Business** | Outlook, Word, Access, Print, PDF, external EXE | Orange |

Environment patterns use 3-tier severity:
- **Risk**: Code that breaks on cloud (e.g. `Dir(ThisWorkbook.Path)`, adjacent file creation)
- **Review**: Context-dependent (e.g. `CurDir`, `BeforeSave` events)
- **Info**: Safe alone but dangerous in combination (e.g. `ThisWorkbook.Path`)

3 modes:
1. **Settings GUI** (no args): Configure detect/sanitize per pattern
2. **File analysis**: Drop file 竊・HTML viewer + text report + sanitized copy + CSV
3. **Folder analysis**: Drop folder 竊・analyze all xlsm/xlam/xls recursively

## Output

```
output/
笏懌楳笏 20260328_120000_extract/
笏・  笏懌楳笏 modules/<baseName>/   .bas / .cls / .frm (per-file subfolder)
笏・  笏披楳笏 <baseName>_combined.txt
笏懌楳笏 20260328_120500_analyze/
笏・  笏懌楳笏 analyze.csv           CSV with all files (EDR/Compat/Env/Biz/judgment columns)
笏・  笏懌楳笏 <name>_analyze.txt    Text report per file
笏・  笏懌楳笏 <name>_analyze.html   HTML viewer (sidebar + code + outline + tooltips)
笏・  笏披楳笏 <name>.xlsm           Sanitized copy (if applicable)
笏懌楳笏 20260328_121000_diff/
笏・  笏懌楳笏 diff.txt
笏・  笏披楳笏 diff.html
笏懌楳笏 20260328_121300_survey/
笏・  笏懌楳笏 survey.txt
笏・  笏披楳笏 survey.json
笏懌楳笏 20260328_121400_envtest/
笏・  笏懌楳笏 envtest.txt
笏・  笏懌楳笏 survey.txt
笏・  笏披楳笏 survey.json
笏・  笏懌楳笏 probe.txt
笏・  笏披楳笏 probe_storage.xlsm
笏披楳笏 20260328_121500_unlock/
    笏披楳笏 <name>.xlsm
```

## Structure

```
vba-devkit/
笏懌楳笏 EnvTest.bat / Extract.bat / Analyze.bat / Diff.bat / Unlock.bat
笏懌楳笏 config/
笏・  笏披楳笏 analyze.json         Detect/sanitize settings per pattern
笏懌楳笏 lib/
笏・  笏懌楳笏 VBAToolkit.psm1      Core: OLE2, VBA compress/decompress, C# (Add-Type),
笏・  笏・                       analysis engine, API replacement DB (60+ entries),
笏・  笏・                       HTML templates
笏・  笏懌楳笏 EnvTest.ps1
笏・  笏懌楳笏 Extract.ps1
笏・  笏懌楳笏 Analyze.ps1
笏・  笏懌楳笏 Diff.ps1
笏・  笏懌楳笏 Unlock.ps1
笏・  笏懌楳笏 internal/Survey.ps1
笏・  笏披楳笏 internal/Probe.ps1
笏懌楳笏 test/                    Test fixtures (.xlsm)
笏披楳笏 docs/                    Specs and investigation results
```

## How it works

OLE2 Compound Document + MS-OVBA decompression via PowerShell + C# (Add-Type). No Excel process except Unlock (.xls conversion) and EnvTest/Probe (test injection).

Survey is separate from Probe internally, but `EnvTest.bat` is the single launcher. Survey inventories what is installed and registered on the machine. Probe performs active VBA / PowerShell / host-side compatibility checks.
