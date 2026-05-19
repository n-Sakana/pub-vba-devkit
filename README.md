# vba-devkit

Excel VBA migration toolkit. Binary-level analysis without opening Excel.

## Tools

### Main

| BAT | Description |
|-----|-------------|
| `EnvTest.bat` | Unified environment test launcher (Survey / Probe / Full) |
| `Extract.bat` | Extract VBA source code (individual modules + combined.txt) |
| `Analyze.bat` | Analysis + sanitization + migration guide + CSV |
| `RescueSheets.bat` | Create a sheet/data rescue copy by removing or wiping VBA code |
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

### Sheet rescue / macro removal

`RescueSheets.bat <path>` is the safest route when an EDR-blocked VBA project makes the entire workbook unusable. It does **not** try to preserve executable macro behavior.

- `.xlsm` / `.xltm`: removes the whole VBA package part and writes a macro-free copy (`*_macrofree.xlsx` / `*_macrofree.xltx`).
- `.xls`: keeps the workbook container but clears every VBA module source stream and zero-fills stale p-code before saving `*_code_removed.xls`.
- `.xlam`: there is no macro-free sheet workbook format for add-ins, so it uses the same code-clearing path as `.xls`.

Use `Sanitize.bat` when you want the older targeted line-level masking. Use `RescueSheets.bat` when sheet structure and cell data matter more than keeping existing code.

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
2. **File analysis**: Drop file → HTML viewer + text report + sanitized copy + CSV
3. **Folder analysis**: Drop folder → analyze all xlsm/xlam/xls recursively

## Output

```
output/
├── 20260328_120000_extract/
│   ├── modules/<baseName>/   .bas / .cls / .frm (per-file subfolder)
│   └── <baseName>_combined.txt
├── 20260328_120500_analyze/
│   ├── analyze.csv           CSV with all files (EDR/Compat/Env/Biz/judgment columns)
│   ├── <name>_analyze.txt    Text report per file
│   ├── <name>_analyze.html   HTML viewer (sidebar + code + outline + tooltips)
│   └── <name>.xlsm           Sanitized copy (if applicable)
├── 20260328_121000_diff/
│   ├── diff.txt
│   └── diff.html
├── 20260328_121300_survey/
│   ├── survey.txt
│   └── survey.json
├── 20260328_121400_envtest/
│   ├── envtest.txt
│   ├── survey.txt
│   ├── survey.json
│   ├── probe.txt
│   └── probe_storage.xlsm
├── 20260328_121450_rescue/
│   └── <name>_macrofree.xlsx / <name>_code_removed.xls
└── 20260328_121500_unlock/
    └── <name>.xlsm
```

## Structure

```
vba-devkit/
├── EnvTest.bat / Extract.bat / Analyze.bat / RescueSheets.bat / Diff.bat / Unlock.bat
├── config/
│   └── analyze.json         Detect/sanitize settings per pattern
├── lib/
│   ├── VBAToolkit.psm1      Core: OLE2, VBA compress/decompress, C# (Add-Type),
│   │                       analysis engine, API replacement DB (60+ entries),
│   │                       HTML templates
│   ├── EnvTest.ps1
│   ├── Extract.ps1
│   ├── Analyze.ps1
│   ├── RescueSheets.ps1
│   ├── Diff.ps1
│   ├── Unlock.ps1
│   ├── internal/Survey.ps1
│   └── internal/Probe.ps1
├── test/                    Test fixtures (.xlsm)
└── docs/                    Specs and investigation results
```

## How it works

OLE2 Compound Document + MS-OVBA decompression via PowerShell + C# (Add-Type). No Excel process except Unlock (.xls conversion) and EnvTest/Probe (test injection).

Survey is separate from Probe internally, but `EnvTest.bat` is the single launcher. Survey inventories what is installed and registered on the machine. Probe performs active VBA / PowerShell / host-side compatibility checks.
