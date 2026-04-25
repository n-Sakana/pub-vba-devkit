# Storage Path Strategy

## Core idea

For SharePoint/OneDrive migration, the first question is not "How do we keep using URL paths?".
It is:

1. Does `Environ$("OneDriveCommercial")` return a local sync root on the target PC?
2. Does `Dir(syncRoot & "\relative\path\*.xlsx")` work from VBA?

If both answers are yes, storage path migration can likely stay inside VBA.

```vb
Dim syncRoot As String
Dim path As String

syncRoot = Environ$("OneDriveCommercial")
path = syncRoot & "\site-name\library-name\folder\"

Debug.Print Dir(path & "*.xlsx")
```

This avoids both failure modes:

- Bad: hard-coded full local path
  - breaks per user because `C:\Users\<name>\...` differs
- Bad: URL path
  - `Dir()` and adjacent file operations break
- Good: `OneDriveCommercial + relative path`
  - user-specific root is resolved dynamically
  - the library-relative path stays shared across users

## What EnvTest should prove

`EnvTest -> Probe` now checks these points directly:

- `OneDrive Environment`
- `Local Sync Root`
- `Local Sync Root Enumeration`
  - `Dir(local sync root)`
  - `Dir(local sync root, *.xls*)`
  - `FSO.GetFolder(local sync root)`
- host-side fallback
  - `Get-ChildItem`
  - `.NET Directory.GetFiles`

## Decision rule

- If VBA-side local sync root tests are `OK`
  - use a VBA common module:
    - resolve sync root from environment variable
    - append shared relative path
    - keep `Dir()` / `FSO` logic in VBA
- If VBA-side tests fail but host-side path enumeration is `OK`
  - use a casedesk-style host/service only for storage enumeration
- If both fail
  - path migration needs a different workflow or product decision

## Window automation rule

Storage and window automation are separate decisions.

- If `P/Invoke` tests are `OK`
  - Win32-based automation via PS/C# is viable
- If `P/Invoke` fails but `UIAutomation` is `OK`
  - try UIAutomation next
- If both fail
  - `WScript.Shell AppActivate / SendKeys` is the last resort
