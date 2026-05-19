# vba-devkit

Excel VBA migration toolkit. Binary-level analysis without opening Excel.

## Tools

### Main

| BAT | Description |
|-----|-------------|
| `EnvTest.bat` | Unified environment test launcher (Survey / Probe / Full) |
| `Extract.bat` | Extract VBA source code (individual modules + combined.txt) |
| `Analyze.bat` | Analysis + migration guide + CSV |
| `Sanitize.bat` | Break EDR-triggering VBA while keeping the original workbook format |
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

## Analyze

Analyze is the core screening tool. It detects risks across 4 categories:

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
1. **Settings GUI** (no args): Configure detection per pattern
2. **File analysis**: Drop file -> HTML viewer + text report + CSV
3. **Folder analysis**: Drop folder -> analyze all xlsm/xlam/xls recursively

## Sheet Rescue

`RescueSheets.bat <path>` is the safest route when an EDR-blocked VBA project makes the entire workbook unusable. It does not try to preserve executable macro behavior.

- `.xlsm` / `.xltm`: removes the whole VBA package part and writes a macro-free copy (`*_macrofree.xlsx` / `*_macrofree.xltx`).
- `.xls`: keeps the workbook container but clears every VBA module source stream and zero-fills stale p-code before saving `*_code_removed.xls`.
- `.xlam`: there is no macro-free sheet workbook format for add-ins, so it uses the same code-clearing path as `.xls`.

## Sanitize

`Sanitize.bat [mode 1-10] <path>` accepts `.xls`, `.xlsm`, and `.xlam` files. If a folder is passed, files are processed recursively. PowerShell callers can use `lib\Sanitize.ps1 -Mode 6 -Path <path>`.

Sanitize is for sheet rescue, not EDR bypass. It irreversibly breaks VBA statements that match the same EDR/NG criteria used by Analyze: `Win32 API (Declare)`, `Shell / process`, and `PowerShell / WScript`. It also extracts callable names from `Declare` statements and breaks their call sites. Statements with VBA line continuation (`_`) are replaced as a whole.

Output is written under `output/<timestamp>_sanitize/` as `<name>_sanitized.<ext>` plus `sanitize.csv`. The original workbook is never overwritten. Replaced VBA lines become harmless comments containing `***`; original dangerous words such as API names or process-launch terms are not kept in those comments.

The sanitized workbook is expected to have broken macros. The goal is to make Excel workbook structure, worksheets, and cell data recoverable while retaining the original workbook format. The sanitizer rewrites compressed VBA source while preserving the existing p-code/performance-cache prefix, because previous zero-fill attempts corrupted workbooks.

Modes:

| Mode | Style | Notes |
|------|-------|-------|
| 1 | Safe readable metadata | No original danger words; records role, library family, argument count, return type. |
| 2 | VBA `'` comment with original text | Preserves original dangerous words in a VBA comment. May still trip EDR. |
| 3 | `Rem` comment with original text | Preserves original dangerous words; marked experimental because scanners/analyzers may still match it. |
| 4 | `//` line with original text | Non-VBA comment style; intentionally breaks code and preserves original text. |
| 5 | `/* ... */` line with original text | C-style marker; intentionally breaks code and preserves original text. |
| 6 | Light partial mask | Masks danger tokens lightly, e.g. recognizable first/last chunks. |
| 7 | Medium partial mask | More masking than mode 6. |
| 8 | Strong partial mask | Mostly first/last character only. |
| 9 | Initial-only mask | Keeps only initial character for danger tokens. |
| 10 | Skeleton mask | Masks all alphanumerics in the unsafe statement while keeping punctuation shape. |

Use `RescueSheets.bat` when sheet structure and cell data matter more than keeping the original macro-enabled file format. Use `Sanitize.bat` when you need targeted line-level destruction inside the original workbook format.

## Output

```text
output/
├── 20260328_120000_extract/
│   ├── modules/<baseName>/   .bas / .cls / .frm (per-file subfolder)
│   └── <baseName>_combined.txt
├── 20260328_120500_analyze/
│   ├── analyze.csv           CSV with all files (EDR/Compat/Env/Biz/judgment columns)
│   ├── <name>_analyze.txt    Text report per file
│   └── <name>_analyze.html   HTML viewer (sidebar + code + outline + tooltips)
├── 20260328_120700_sanitize/
│   ├── sanitize.csv
│   └── <name>_sanitized.xlsm
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

```text
vba-devkit/
├── EnvTest.bat / Extract.bat / Analyze.bat / Sanitize.bat / RescueSheets.bat / Diff.bat / Unlock.bat
├── config/
│   └── analyze.json         Detection settings per pattern
├── lib/
│   ├── VBAToolkit.psm1      Core: OLE2, VBA compress/decompress, C# (Add-Type),
│   │                       analysis engine, API replacement DB (60+ entries),
│   │                       HTML templates
│   ├── EnvTest.ps1
│   ├── Extract.ps1
│   ├── Analyze.ps1
│   ├── Sanitize.ps1
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
