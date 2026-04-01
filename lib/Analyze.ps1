param([string[]]$Paths)
$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot\VBAToolkit.psm1" -Force -DisableNameChecking

$configDir = Join-Path (Split-Path "$PSScriptRoot" -Parent) 'config'
$configPath = Join-Path $configDir 'analyze.json'

# ============================================================================
# Evidence basis mapping (probe results 2026-03-30)
# ============================================================================

$script:EvidenceMap = @{
    # EDR - observed BLOCKED
    'Win32 API (Declare)' = @{ Basis = 'observed'; Note = 'Declare is BLOCKED in probe (file corruption on save/reopen)' }
    'Shell / process'     = @{ Basis = 'observed'; Note = 'WScript.Shell.Run FAIL (write error on Run)' }
    'PowerShell / WScript' = @{ Basis = 'observed'; Note = 'powershell via VBA FAIL (write error on Run)' }

    # Compat - observed FAIL
    'Deprecated: DAO'       = @{ Basis = 'observed'; Note = 'DAO.DBEngine.36 FAIL (class not registered)' }
    'Deprecated: Legacy Controls' = @{ Basis = 'observed'; Note = 'MSComDlg.CommonDialog / MSCAL.Calendar FAIL' }
    'Deprecated: DDE'       = @{ Basis = 'observed'; Note = 'DDEInitiate FAIL' }
    'SendKeys'              = @{ Basis = 'observed'; Note = 'SendKeys OK in probe but unstable across environments' }
    'AppActivate'           = @{ Basis = 'observed'; Note = 'AppActivate FAIL in probe' }
    'keybd_event'           = @{ Basis = 'inference'; Note = 'GUI API -- unstable without foreground window' }
    'Sleep'                 = @{ Basis = 'inference'; Note = 'kernel32 Sleep -- Declare required, will be BLOCKED' }
    'Deprecated: IE Automation' = @{ Basis = 'observed'; Note = 'IE removed from OS' }
    'DLL loading'           = @{ Basis = 'inference'; Note = 'LoadLibrary/GetProcAddress -- Declare required, likely BLOCKED' }
    '64-bit: Missing PtrSafe' = @{ Basis = 'inference'; Note = 'Legacy Declare without PtrSafe -- will not compile on 64-bit' }
    '64-bit: Long for handles' = @{ Basis = 'inference'; Note = 'Handle as Long in PtrSafe -- may truncate on 64-bit' }
    '64-bit: VarPtr/ObjPtr/StrPtr' = @{ Basis = 'inference'; Note = 'Returns LongPtr on 64-bit' }

    # Compat - observed OK but kept
    'COM / CreateObject' = @{ Basis = 'observed'; Note = 'Standard COM CreateObject OK in probe' }
    'COM / GetObject'    = @{ Basis = 'observed'; Note = 'GetObject OK in probe' }

    # Path - inference (design issue, not security block)
    'Fixed drive letter'     = @{ Basis = 'inference'; Note = 'Hardcoded drive letter -- breaks on OneDrive/SharePoint' }
    'UNC path'               = @{ Basis = 'inference'; Note = 'Hardcoded UNC -- may change in new environment' }
    'User folder'            = @{ Basis = 'inference'; Note = 'Hardcoded C:\Users path -- breaks on profile change' }
    'Desktop / Documents'    = @{ Basis = 'inference'; Note = 'Hardcoded known folder -- may redirect to OneDrive' }
    'AppData'                = @{ Basis = 'inference'; Note = 'Hardcoded AppData path' }
    'Program Files'          = @{ Basis = 'inference'; Note = 'Hardcoded Program Files path' }
    'ThisWorkbook.Path'      = @{ Basis = 'inference'; Note = 'ThisWorkbook.Path as root -- URL on OneDrive/SharePoint' }
    'Dir() path check'       = @{ Basis = 'inference'; Note = 'Dir() for path existence -- may fail with URL paths' }
    'Path concatenation'     = @{ Basis = 'inference'; Note = 'String path concatenation -- fragile with URL paths' }
    'SaveAs call'            = @{ Basis = 'inference'; Note = 'SaveAs with path -- review target location' }
    'CurDir'                 = @{ Basis = 'inference'; Note = 'CurDir -- undefined on OneDrive sync folders' }
    'ChDir'                  = @{ Basis = 'inference'; Note = 'ChDir -- may fail on URL paths' }
    'External workbook open (literal)' = @{ Basis = 'inference'; Note = 'Workbooks.Open with literal path' }
    'External workbook ref'  = @{ Basis = 'inference'; Note = 'External workbook reference -- path may change' }
    'BeforeSave event'       = @{ Basis = 'inference'; Note = 'BeforeSave handler -- review save flow for new storage' }
    'AfterSave event'        = @{ Basis = 'inference'; Note = 'AfterSave handler -- review save flow for new storage' }
    'LinkSources / UpdateLink' = @{ Basis = 'inference'; Note = 'LinkSources -- external links may break on path change' }
    'Workbooks.Open (variable)' = @{ Basis = 'inference'; Note = 'Workbooks.Open with variable path -- review source' }
    'Connection string'      = @{ Basis = 'inference'; Note = 'Connection string -- server/path may change' }
    'Fixed printer name'     = @{ Basis = 'inference'; Note = 'Hardcoded printer name -- may not exist in new env' }
    'Fixed IP address'       = @{ Basis = 'inference'; Note = 'Hardcoded IP -- may change in new network' }
    'Fixed connection host'  = @{ Basis = 'inference'; Note = 'Hardcoded hostname -- may change' }
    'localhost'              = @{ Basis = 'inference'; Note = 'localhost reference' }
}

# ============================================================================
# 3-axis remapping: old 4-axis (edr/compat/env/biz) -> new 3-axis (edr/compat/path)
# ============================================================================

# Which old-axis patterns belong to new EDR axis (only probe-BLOCKED patterns)
$script:EdrPatterns = @(
    'Win32 API (Declare)'
    'Shell / process'
    'PowerShell / WScript'
)

# Which old-axis patterns belong to new Compat axis
$script:CompatPatterns = @(
    # From old compat axis
    '64-bit: Missing PtrSafe'
    '64-bit: Long for handles'
    '64-bit: VarPtr/ObjPtr/StrPtr'
    'Deprecated: DDE'
    'Deprecated: IE Automation'
    'Deprecated: Legacy Controls'
    'Deprecated: DAO'
    'Legacy: DefType'
    'Legacy: GoSub'
    'Legacy: While/Wend'
    # From old edr axis (not actually blocked by EDR)
    'DLL loading'
    'SendKeys'
    'AppActivate'
)

# Which old-axis patterns belong to new Path axis
$script:PathPatterns = @(
    'Fixed drive letter'
    'UNC path'
    'User folder'
    'Desktop / Documents'
    'AppData'
    'Program Files'
    'Fixed printer name'
    'Fixed IP address'
    'Fixed connection host'
    'localhost'
    'Connection string'
    'External workbook open (literal)'
    'Dir() path check'
    'Path concatenation'
    'SaveAs call'
    'External workbook ref'
    'BeforeSave event'
    'AfterSave event'
    'LinkSources / UpdateLink'
    'Workbooks.Open (variable)'
    'CurDir'
    'ChDir'
    'ThisWorkbook.Path'
)

function Get-NewAxis {
    param([string]$PatternName)
    if ($script:EdrPatterns -contains $PatternName) { return 'edr' }
    if ($script:CompatPatterns -contains $PatternName) { return 'compat' }
    if ($script:PathPatterns -contains $PatternName) { return 'path' }
    return $null  # not mapped (e.g. info-only old patterns we skip)
}

function Get-EvidenceBasis {
    param([string]$PatternName)
    $ev = $script:EvidenceMap[$PatternName]
    if ($ev) { return $ev.Basis }
    return 'inference'
}

function Get-EvidenceNote {
    param([string]$PatternName)
    $ev = $script:EvidenceMap[$PatternName]
    if ($ev) { return $ev.Note }
    return ''
}

# ============================================================================
# Config helpers (3-axis)
# ============================================================================

function Get-DefaultConfig {
    return @{
        edr = [ordered]@{
            'Win32 API (Declare)' = @{ detect = $true }
            'Shell / process' = @{ detect = $true }
            'PowerShell / WScript' = @{ detect = $true }
        }
        compat = [ordered]@{
            'DLL loading' = @{ detect = $true }
            'SendKeys' = @{ detect = $true }
            'AppActivate' = @{ detect = $true }
            '64-bit: Missing PtrSafe' = @{ detect = $false }
            '64-bit: Long for handles' = @{ detect = $true }
            '64-bit: VarPtr/ObjPtr/StrPtr' = @{ detect = $true }
            'Deprecated: DDE' = @{ detect = $true }
            'Deprecated: IE Automation' = @{ detect = $true }
            'Deprecated: Legacy Controls' = @{ detect = $true }
            'Deprecated: DAO' = @{ detect = $true }
            'Legacy: DefType' = @{ detect = $false }
            'Legacy: GoSub' = @{ detect = $false }
            'Legacy: While/Wend' = @{ detect = $false }
        }
        path = [ordered]@{
            'Fixed drive letter' = @{ detect = $true }
            'UNC path' = @{ detect = $true }
            'User folder' = @{ detect = $true }
            'Desktop / Documents' = @{ detect = $true }
            'AppData' = @{ detect = $true }
            'Program Files' = @{ detect = $true }
            'Fixed printer name' = @{ detect = $true }
            'Fixed IP address' = @{ detect = $true }
            'Fixed connection host' = @{ detect = $true }
            'localhost' = @{ detect = $false }
            'Connection string' = @{ detect = $true }
            'External workbook open (literal)' = @{ detect = $true }
            'Dir() path check' = @{ detect = $true }
            'Path concatenation' = @{ detect = $true }
            'SaveAs call' = @{ detect = $true }
            'External workbook ref' = @{ detect = $true }
            'BeforeSave event' = @{ detect = $true }
            'AfterSave event' = @{ detect = $true }
            'LinkSources / UpdateLink' = @{ detect = $true }
            'Workbooks.Open (variable)' = @{ detect = $true }
            'CurDir' = @{ detect = $true }
            'ChDir' = @{ detect = $true }
            'ThisWorkbook.Path' = @{ detect = $true }
        }
    }
}

function Load-AnalyzeConfig {
    param([string]$Path)
    if (Test-Path "$Path") {
        $json = Get-Content "$Path" -Raw -Encoding UTF8 | ConvertFrom-Json
        $cfg = Get-DefaultConfig
        foreach ($section in @('edr','compat','path')) {
            if ($json.$section) {
                $json.$section.PSObject.Properties | ForEach-Object {
                    if ($cfg.$section.Contains($_.Name)) {
                        $d = $true
                        if ($null -ne $_.Value.detect) { $d = [bool]$_.Value.detect }
                        $cfg.$section[$_.Name] = @{ detect = $d }
                    }
                }
            }
        }
        return $cfg
    }
    return Get-DefaultConfig
}

function Save-AnalyzeConfig {
    param([hashtable]$Config, [string]$Path)
    $dir = Split-Path "$Path" -Parent
    if (-not (Test-Path "$dir")) { [void][IO.Directory]::CreateDirectory($dir) }
    $obj = [ordered]@{ edr = [ordered]@{}; compat = [ordered]@{}; path = [ordered]@{} }
    foreach ($section in @('edr','compat','path')) {
        if ($Config.$section) {
            foreach ($k in $Config.$section.Keys) {
                $obj.$section[$k] = [ordered]@{ detect = $Config.$section[$k].detect }
            }
        }
    }
    $obj | ConvertTo-Json -Depth 3 | Set-Content "$Path" -Encoding UTF8
}

# ============================================================================
# Remap old 4-axis analysis results to new 3-axis structure
# ============================================================================

function Remap-AnalysisTo3Axis {
    param([hashtable]$Analysis)

    $edrHits = 0; $compatHits = 0; $pathHits = 0
    $edrFindings = [ordered]@{}
    $compatFindings = [ordered]@{}
    $pathFindings = [ordered]@{}

    # Remap old EDR findings (Findings)
    foreach ($cat in $Analysis.Findings.Keys) {
        $axis = Get-NewAxis $cat
        $f = $Analysis.Findings[$cat]
        $count = $f.Findings.Count
        if ($count -eq 0) { continue }
        switch ($axis) {
            'edr'    { $edrFindings[$cat] = $f; $edrHits += $count }
            'compat' { $compatFindings[$cat] = $f; $compatHits += $count }
            'path'   { $pathFindings[$cat] = $f; $pathHits += $count }
        }
    }

    # Remap old compat findings (CompatFindings)
    foreach ($cat in $Analysis.CompatFindings.Keys) {
        $axis = Get-NewAxis $cat
        $f = $Analysis.CompatFindings[$cat]
        $count = $f.Findings.Count
        if ($count -eq 0) { continue }
        switch ($axis) {
            'edr'    { $edrFindings[$cat] = $f; $edrHits += $count }
            'compat' { $compatFindings[$cat] = $f; $compatHits += $count }
            'path'   { $pathFindings[$cat] = $f; $pathHits += $count }
        }
    }

    # Remap old env findings (EnvFindings)
    foreach ($cat in $Analysis.EnvFindings.Keys) {
        $axis = Get-NewAxis $cat
        $f = $Analysis.EnvFindings[$cat]
        $count = $f.Findings.Count
        if ($count -eq 0) { continue }
        switch ($axis) {
            'edr'    { $edrFindings[$cat] = $f; $edrHits += $count }
            'compat' { $compatFindings[$cat] = $f; $compatHits += $count }
            'path'   { $pathFindings[$cat] = $f; $pathHits += $count }
        }
    }

    # Remap old env info findings (InfoFindings) -> path axis
    foreach ($cat in $Analysis.InfoFindings.Keys) {
        $axis = Get-NewAxis $cat
        $f = $Analysis.InfoFindings[$cat]
        $count = $f.Findings.Count
        if ($count -eq 0) { continue }
        if ($axis -eq 'path') {
            $pathFindings[$cat] = $f; $pathHits += $count
        }
    }

    # Remap old biz findings: not in 3-axis, but we check for compat-like patterns
    # Biz findings that have a new-axis mapping go there; others are dropped from main counts
    foreach ($cat in $Analysis.BizFindings.Keys) {
        $axis = Get-NewAxis $cat
        $f = $Analysis.BizFindings[$cat]
        $count = $f.Findings.Count
        if ($count -eq 0) { continue }
        if ($axis) {
            switch ($axis) {
                'compat' { $compatFindings[$cat] = $f; $compatHits += $count }
                'path'   { $pathFindings[$cat] = $f; $pathHits += $count }
            }
        }
    }

    # Count API call-site lines (these are highlighted in HTML as hl-edr)
    $apiCallLineCount = 0
    if ($Analysis.ApiCallNames.Count -gt 0 -and $Analysis.AllCode) {
        foreach ($fileName in $Analysis.AllCode.Keys) {
            foreach ($line in $Analysis.AllCode[$fileName]) {
                if ($line -match '^\s*''') { continue }
                foreach ($apiName in $Analysis.ApiCallNames) {
                    if ($line -match "\b$([regex]::Escape($apiName))\b") {
                        $apiCallLineCount++
                        break
                    }
                }
            }
        }
    }

    return @{
        EdrHits = $edrHits + $apiCallLineCount
        CompatHits = $compatHits
        PathHits = $pathHits
        EdrFindings = $edrFindings
        CompatFindings = $compatFindings
        PathFindings = $pathFindings
    }
}

# ============================================================================
# Build hits list (line-level findings for hits.csv and _hits.json)
# ============================================================================

function Build-HitsList {
    param(
        [hashtable]$Project,
        [hashtable]$Remapped,
        [string]$FileName,
        [hashtable]$DetectRulesByName
    )

    $hits = [System.Collections.ArrayList]::new()

    foreach ($modName in $Project.Modules.Keys) {
        $mod = $Project.Modules[$modName]
        $lines = if ($mod.Lines -is [array]) { $mod.Lines } else { @($mod.Lines) }

        for ($i = 0; $i -lt $lines.Count; $i++) {
            $line = $lines[$i]
            if ($line -match '^\s*''') { continue }

            foreach ($ruleName in $DetectRulesByName.Keys) {
                $rule = $DetectRulesByName[$ruleName]
                if ($line -match $rule.Pattern) {
                    $axis = Get-NewAxis $ruleName
                    if (-not $axis) { continue }
                    $snippet = $line.Trim()
                    if ($snippet.Length -gt 120) { $snippet = $snippet.Substring(0, 117) + '...' }
                    [void]$hits.Add([ordered]@{
                        MacroName     = $FileName
                        Category      = $axis
                        PatternName   = $ruleName
                        Module        = "$modName.$($mod.Ext)"
                        LineNum       = ($i + 1)
                        CodeSnippet   = $snippet
                        EvidenceBasis = Get-EvidenceBasis $ruleName
                    })
                }
            }
        }
    }

    return $hits
}

# ============================================================================
# Build CSV row (3-axis)
# ============================================================================

function Build-AnalyzeCsvRow {
    param(
        [hashtable]$Analysis,
        [hashtable]$Remapped,
        [hashtable]$Project,
        [string]$FileName,
        [string]$RelPath,
        [System.Collections.ArrayList]$Hits
    )

    # Module counts
    $bas = 0; $cls = 0; $frm = 0; $totalModules = 0; $codeLines = 0
    foreach ($modName in $Project.Modules.Keys) {
        $mod = $Project.Modules[$modName]
        $totalModules++
        switch ($mod.Ext) { 'bas' { $bas++ } 'cls' { $cls++ } 'frm' { $frm++ } }
        $codeLines += $mod.Lines.Count
    }

    # RiskLevel based on 3-axis
    $riskLevel = 'Low'
    if ($Remapped.EdrHits -gt 0) { $riskLevel = 'High' }
    elseif ($Remapped.CompatHits -gt 0) { $riskLevel = 'Medium' }
    elseif ($Remapped.PathHits -ge 3) { $riskLevel = 'Medium' }

    # MigrationClass
    $totalHits = $Remapped.EdrHits + $Remapped.CompatHits + $Remapped.PathHits
    $migClasses = [System.Collections.ArrayList]::new()
    if ($totalHits -eq 0) {
        [void]$migClasses.Add('NoChange')
    } else {
        if ($Remapped.EdrHits -gt 0) {
            [void]$migClasses.Add('Rebuild')
        }
        if ($Remapped.CompatFindings.Contains('Deprecated: DAO') -or
            $Remapped.CompatFindings.Contains('DLL loading') -or
            $Remapped.CompatFindings.Contains('Deprecated: IE Automation')) {
            [void]$migClasses.Add('NeedsReplacement')
        }
        if ($Remapped.PathHits -gt 0) {
            [void]$migClasses.Add('StorageReview')
        }
        if ($Remapped.CompatHits -gt 0 -and $migClasses.Count -eq 0) {
            [void]$migClasses.Add('MinorFix')
        }
        if ($migClasses.Count -eq 0) {
            [void]$migClasses.Add('MinorFix')
        }
    }

    # Evidence summary: top observed findings
    $evidenceParts = [System.Collections.ArrayList]::new()
    foreach ($cat in $Remapped.EdrFindings.Keys) {
        $ev = Get-EvidenceBasis $cat
        $n = $Remapped.EdrFindings[$cat].Findings.Count
        [void]$evidenceParts.Add("[$ev] $cat($n)")
        if ($evidenceParts.Count -ge 3) { break }
    }
    foreach ($cat in $Remapped.CompatFindings.Keys) {
        if ($evidenceParts.Count -ge 5) { break }
        $ev = Get-EvidenceBasis $cat
        $n = $Remapped.CompatFindings[$cat].Findings.Count
        [void]$evidenceParts.Add("[$ev] $cat($n)")
    }
    foreach ($cat in $Remapped.PathFindings.Keys) {
        if ($evidenceParts.Count -ge 6) { break }
        $ev = Get-EvidenceBasis $cat
        $n = $Remapped.PathFindings[$cat].Findings.Count
        [void]$evidenceParts.Add("[$ev] $cat($n)")
    }

    # ReviewNote
    $reviewNote = ''
    if ($Remapped.EdrHits -gt 0) { $reviewNote = 'EDR blocking detected -- sanitize or rebuild required' }
    elseif ($Remapped.CompatFindings.Contains('Deprecated: DAO')) { $reviewNote = 'DAO dependency -- needs ADO replacement' }
    elseif ($Remapped.PathHits -ge 3) { $reviewNote = 'Multiple hardcoded paths -- storage migration review needed' }

    $row = [ordered]@{
        Timestamp      = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        RelativePath   = $(if ([IO.Path]::GetDirectoryName($RelPath)) { [IO.Path]::GetDirectoryName($RelPath) } else { '.' })
        FileName       = $FileName
        Bas            = $bas
        Cls            = $cls
        Frm            = $frm
        TotalModules   = $totalModules
        CodeLines      = $codeLines
        EdrHits        = $Remapped.EdrHits
        CompatHits     = $Remapped.CompatHits
        PathHits       = $Remapped.PathHits
        RiskLevel      = $riskLevel
        MigrationClass = $migClasses -join '; '
        Evidence       = $evidenceParts -join '; '
        ReviewNote     = $reviewNote
    }
    return $row
}

# ============================================================================
# Text report (3-axis)
# ============================================================================

function Build-AnalyzeTextReport {
    param(
        [hashtable]$Analysis,
        [hashtable]$Remapped,
        $AllModLines,
        [System.Collections.Specialized.OrderedDictionary]$CsvRow,
        [string]$FileName,
        [hashtable]$Replacements
    )
    $txtSb = [System.Text.StringBuilder]::new()
    [void]$txtSb.AppendLine("# VBA Screening Report")
    [void]$txtSb.AppendLine("# Source: $FileName")
    [void]$txtSb.AppendLine("# Date: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$txtSb.AppendLine("")
    [void]$txtSb.AppendLine("## Modules ($($CsvRow.TotalModules))")
    foreach ($modName in $AllModLines.Keys) {
        $ml = $AllModLines[$modName]
        [void]$txtSb.AppendLine("  $modName.$($ml.Ext) ($($ml.Lines.Count) lines)")
    }
    [void]$txtSb.AppendLine("  Total: $($CsvRow.CodeLines) lines")
    [void]$txtSb.AppendLine("")

    # Section 1: EDR Risks
    if ($Remapped.EdrFindings.Count -gt 0) {
        [void]$txtSb.AppendLine("## EDR Risks ($($Remapped.EdrHits)) -- probe-BLOCKED patterns")
        foreach ($cat in $Remapped.EdrFindings.Keys) {
            $f = $Remapped.EdrFindings[$cat]
            $basis = Get-EvidenceBasis $cat
            $note = Get-EvidenceNote $cat
            [void]$txtSb.AppendLine("  $cat ($($f.Findings.Count)) [$basis]")
            [void]$txtSb.AppendLine("    Evidence: $note")
            if ($f.Aggregate) {
                [void]$txtSb.AppendLine("    (aggregated: $($f.Findings.Count) occurrences)")
            } else {
                foreach ($finding in $f.Findings) { [void]$txtSb.AppendLine("    $finding") }
            }
        }
        [void]$txtSb.AppendLine("")
    }

    # Section 2: Compatibility Risks
    if ($Remapped.CompatFindings.Count -gt 0) {
        [void]$txtSb.AppendLine("## Compatibility Risks ($($Remapped.CompatHits))")
        foreach ($cat in $Remapped.CompatFindings.Keys) {
            $f = $Remapped.CompatFindings[$cat]
            $basis = Get-EvidenceBasis $cat
            $note = Get-EvidenceNote $cat
            [void]$txtSb.AppendLine("  $cat ($($f.Findings.Count)) [$basis]")
            if ($note) { [void]$txtSb.AppendLine("    Evidence: $note") }
            foreach ($finding in $f.Findings) { [void]$txtSb.AppendLine("    $finding") }
        }
        [void]$txtSb.AppendLine("")
    }

    # Section 3: Hardcoded Path Risks
    if ($Remapped.PathFindings.Count -gt 0) {
        [void]$txtSb.AppendLine("## Hardcoded Path Risks ($($Remapped.PathHits))")
        foreach ($cat in $Remapped.PathFindings.Keys) {
            $f = $Remapped.PathFindings[$cat]
            $basis = Get-EvidenceBasis $cat
            $note = Get-EvidenceNote $cat
            [void]$txtSb.AppendLine("  $cat ($($f.Findings.Count)) [$basis]")
            if ($note) { [void]$txtSb.AppendLine("    Evidence: $note") }
            foreach ($finding in $f.Findings) { [void]$txtSb.AppendLine("    $finding") }
        }
        [void]$txtSb.AppendLine("")
    }

    # Win32 API details
    if ($Analysis.ApiDecls.Count -gt 0) {
        [void]$txtSb.AppendLine("## Win32 API Usage Details")
        foreach ($decl in $Analysis.ApiDecls) {
            [void]$txtSb.AppendLine("  $($decl.Name)")
            [void]$txtSb.AppendLine("    $($decl.File) L$($decl.Line): $($decl.Sig)")
            $info = $Replacements[$decl.Name]
            if ($info) { [void]$txtSb.AppendLine("    Alternative: $($info.Alt)") }
        }
        [void]$txtSb.AppendLine("")
    }

    # External References
    if ($Analysis.ExternalRefs.Count -gt 0) {
        [void]$txtSb.AppendLine("## External References ($($Analysis.ExternalRefs.Count))")
        foreach ($ref in $Analysis.ExternalRefs) { [void]$txtSb.AppendLine("  $ref") }
        [void]$txtSb.AppendLine("")
    }

    [void]$txtSb.AppendLine("## Summary")
    [void]$txtSb.AppendLine("  $($CsvRow.EdrHits) EDR, $($CsvRow.CompatHits) compat, $($CsvRow.PathHits) path")
    [void]$txtSb.AppendLine("  RiskLevel: $($CsvRow.RiskLevel) | MigrationClass: $($CsvRow.MigrationClass)")

    return $txtSb.ToString()
}

# ============================================================================
# HTML viewer (3-axis colors: edr=blue, compat=purple, path=green)
# ============================================================================

function Build-AnalyzeHtml {
    param(
        $AllModLines,
        [hashtable]$ModHighlights,
        [System.Collections.Specialized.OrderedDictionary]$TooltipEntries,
        [string]$OutPrefix,
        [string]$FileName,
        [System.Collections.Specialized.OrderedDictionary]$CsvRow,
        [string]$OutDir,
        [hashtable]$PatternDefs
    )

    $he = { param($s) [System.Net.WebUtility]::HtmlEncode($s) }

    # Build sidebar
    $sidebarSb = [System.Text.StringBuilder]::new()
    $modIdx = 0; $firstHlIdx = -1
    foreach ($modName in $AllModLines.Keys) {
        $ml = $AllModLines[$modName]
        $hlCount = 0
        if ($ModHighlights[$modName]) { $hlCount = $ModHighlights[$modName].Count }
        $cls = if ($hlCount -gt 0) { 'has-hl' } else { 'no-hl' }
        if ($firstHlIdx -eq -1 -and $hlCount -gt 0) { $firstHlIdx = $modIdx }
        $label = "$modName.$($ml.Ext)"
        if ($hlCount -gt 0) { $label += " ($hlCount)" }
        [void]$sidebarSb.Append("<div class=`"item $cls`" onclick=`"showTab($modIdx)`" id=`"tab$modIdx`">$(& $he $label)</div>")
        $modIdx++
    }
    if ($firstHlIdx -eq -1) { $firstHlIdx = 0 }

    # Build content
    $contentSb = [System.Text.StringBuilder]::new()
    $modIdx = 0
    foreach ($modName in $AllModLines.Keys) {
        $ml = $AllModLines[$modName]
        $hlMap = $ModHighlights[$modName]
        [void]$contentSb.Append("<div class=`"module`" id=`"mod$modIdx`"><table class=`"code-table`">")
        for ($i = 0; $i -lt $ml.Lines.Count; $i++) {
            $trClass = ''
            $dataApi = ''
            if ($hlMap -and $hlMap.ContainsKey($i)) {
                $hl = $hlMap[$i]
                $trClass = $hl.Color
                $dataApi = $hl.PatternName
            }
            $ln = $i + 1
            $dataAttr = if ($dataApi) { " data-api=`"$(& $he $dataApi)`"" } else { '' }
            [void]$contentSb.Append("<tr class=`"$trClass`"$dataAttr><td class=`"ln`">$ln</td><td class=`"code`">$(& $he $ml.Lines[$i])</td></tr>")
        }
        [void]$contentSb.Append("</table></div>")
        $modIdx++
    }

    # Build outline items
    $outlineItems = [System.Collections.ArrayList]::new()
    foreach ($modName in $AllModLines.Keys) {
        $hlMap = $ModHighlights[$modName]
        if (-not $hlMap) { continue }
        $ext = $AllModLines[$modName].Ext
        foreach ($lineIdx in ($hlMap.Keys | Sort-Object { [int]$_ })) {
            $hl = $hlMap[$lineIdx]
            $ln = [int]$lineIdx + 1
            $label = "L$ln $($hl.PatternName)"
            if ($label.Length -gt 50) { $label = $label.Substring(0, 47) + '...' }
            [void]$outlineItems.Add(@{ ModName = "$modName.$ext"; LineNum = $ln; Label = $label; Color = $hl.Color })
        }
    }

    # Build tooltip JS data
    $tooltipJsSb = [System.Text.StringBuilder]::new()
    [void]$tooltipJsSb.Append('{')
    $first = $true
    foreach ($key in $TooltipEntries.Keys) {
        $info = $TooltipEntries[$key]
        $altJs = ($info.Alt -replace '\\','\\\\' -replace "'","\'")
        $noteJs = ($info.Note -replace '\\','\\\\' -replace "'","\'")
        $exJs = ((& $he $info.Example) -replace '\\','\\\\' -replace "'","\'" -replace "`r`n",'\n' -replace "`n",'\n')
        $comma = if ($first) { '' } else { ',' }
        $first = $false
        [void]$tooltipJsSb.Append("$comma'$(& $he $key)':{alt:'$altJs',note:'$noteJs',ex:'$exJs'}")
    }
    [void]$tooltipJsSb.Append('}')

    # CSS (3 axes: edr=blue, compat=purple, path=green)
    $analyzeCss = @"
.sidebar .item.has-hl { color: #e8ab53; }
.sidebar .item.no-hl { color: #606060; }
.code-table { width: 100%; border-collapse: collapse; }
.code-table td { padding: 0 8px; line-height: 20px; vertical-align: top; white-space: pre; overflow: hidden; text-overflow: ellipsis; }
.code-table .ln { width: 50px; min-width: 50px; text-align: right; color: #606060; padding-right: 12px; user-select: none; border-right: 1px solid #3c3c3c; }
.code-table .code { color: #d4d4d4; }
tr.hl-edr td.code { background: #1b2e4a; color: #a0c4f0; cursor: pointer; }
tr.hl-edr td.ln { color: #cccccc; }
tr.hl-compat td.code { background: #3a1b4a; color: #c4a0f0; cursor: pointer; }
tr.hl-compat td.ln { color: #cccccc; }
tr.hl-path td.code { background: #1b3a2a; color: #a0f0c4; cursor: pointer; }
tr.hl-path td.ln { color: #cccccc; }
.minimap { right: 250px; }
.minimap .mark.m-hl-edr { background: #4fc1ff; }
.minimap .mark.m-hl-compat { background: #9a5eff; }
.minimap .mark.m-hl-path { background: #50d090; }
.outline { width: 250px; min-width: 250px; background: #252526; border-left: 1px solid #3c3c3c; overflow-y: auto; padding: 8px 0; }
.outline .ol-header { padding: 6px 12px; font-size: 11px; color: #888; text-transform: uppercase; }
.outline .ol-item { padding: 3px 12px; font-size: 12px; cursor: pointer; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.outline .ol-item:hover { background: #2a2d2e; }
.outline .ol-item.c-edr { color: #4fc1ff; }
.outline .ol-item.c-compat { color: #9a5eff; }
.outline .ol-item.c-path { color: #50d090; }
.hover-hint { position: fixed; background: #444; color: #ccc; padding: 2px 8px; border-radius: 3px; font-size: 11px; pointer-events: none; z-index: 50; display: none; }
.tooltip { position: fixed; background: #2d2d2d; border: 1px solid #555; border-radius: 4px; padding: 10px 14px; max-width: 500px; z-index: 100; display: none; font-size: 12px; line-height: 1.5; box-shadow: 0 4px 12px rgba(0,0,0,0.5); user-select: text; }
.tooltip .tt-api { color: #4fc1ff; font-weight: bold; font-size: 14px; }
.tooltip .tt-alt { color: #6a9955; margin-top: 4px; }
.tooltip .tt-note { color: #b0b0b0; font-style: italic; margin-top: 4px; }
.tooltip .tt-evidence { color: #e8ab53; margin-top: 4px; font-size: 11px; }
.tooltip pre { background: #1e1e1e; border: 1px solid #3c3c3c; border-radius: 3px; padding: 8px; margin-top: 6px; font-size: 11px; line-height: 1.4; max-height: 200px; overflow-y: auto; position: relative; }
.tooltip .tt-copy { position: absolute; top: 6px; right: 6px; background: none; border: none; cursor: pointer; opacity: 0.5; padding: 2px; }
.tooltip .tt-copy:hover { opacity: 1; }
.tooltip .tt-copy svg { width: 14px; height: 14px; fill: #ccc; }
"@

    # Extra HTML
    $extraHtml = @"
<div class="outline" id="outline"></div>
<div class="tooltip" id="tooltip"></div>
<div class="hover-hint" id="hoverHint">Click for details</div>
"@

    # Build outline data as JS array
    $olJsSb = [System.Text.StringBuilder]::new()
    [void]$olJsSb.Append('[')
    $olFirst = $true
    foreach ($item in $outlineItems) {
        $comma = if ($olFirst) { '' } else { ',' }
        $olFirst = $false
        $colorCls = switch ($item.Color) {
            'hl-edr' { 'c-edr' }
            'hl-compat' { 'c-compat' }
            'hl-path' { 'c-path' }
            default { '' }
        }
        $modLabel = & $he $item.ModName
        $olLabel = & $he $item.Label
        [void]$olJsSb.Append("${comma}{mod:'$modLabel',ln:$($item.LineNum),label:'$olLabel',cls:'$colorCls'}")
    }
    [void]$olJsSb.Append(']')

    # Build evidence JS data for tooltips
    $evidenceJsSb = [System.Text.StringBuilder]::new()
    [void]$evidenceJsSb.Append('{')
    $evFirst = $true
    foreach ($patName in $script:EvidenceMap.Keys) {
        $ev = $script:EvidenceMap[$patName]
        $comma = if ($evFirst) { '' } else { ',' }
        $evFirst = $false
        $noteJs = ($ev.Note -replace '\\','\\\\' -replace "'","\'")
        [void]$evidenceJsSb.Append("$comma'$(& $he $patName)':{basis:'$($ev.Basis)',note:'$noteJs'}")
    }
    [void]$evidenceJsSb.Append('}')

    # JS
    $analyzeJs = @"
const outline = document.getElementById('outline');
const tooltip = document.getElementById('tooltip');
const hoverHint = document.getElementById('hoverHint');
const apiInfo = $($tooltipJsSb.ToString());
const evidenceInfo = $($evidenceJsSb.ToString());
const outlineData = $($olJsSb.ToString());

var _baseShowTab = showTab;
showTab = function(idx) {
  _baseShowTab(idx);
  updateOutline();
};

function scrollToRow(r) {
  const rRect = r.getBoundingClientRect();
  const cRect = content.getBoundingClientRect();
  const offset = rRect.top - cRect.top + content.scrollTop;
  content.scrollTo({ top: offset - content.clientHeight / 3, behavior: 'smooth' });
}

function updateOutline() {
  outline.innerHTML = '';
  const hdr = document.createElement('div');
  hdr.className = 'ol-header'; hdr.textContent = 'Detected Lines';
  outline.appendChild(hdr);
  const mod = document.querySelector('.module.active');
  if (!mod) return;
  const modIdx = parseInt(mod.id.replace('mod', ''));
  const tabEl = document.getElementById('tab' + modIdx);
  const modName = tabEl ? tabEl.textContent.replace(/ \(\d+\)$/, '') : '';
  const rows = mod.querySelectorAll('tr');
  rows.forEach(r => {
    const cls = r.className;
    if (!cls || (!cls.includes('hl-edr') && !cls.includes('hl-compat') && !cls.includes('hl-path'))) return;
    const ln = r.querySelector('.ln');
    if (!ln) return;
    const lineNum = ln.textContent;
    const api = r.dataset.api || '';
    const label = 'L' + lineNum + ' ' + api;
    const colorCls = cls.includes('hl-edr') ? 'c-edr' : cls.includes('hl-compat') ? 'c-compat' : 'c-path';
    const item = document.createElement('div');
    item.className = 'ol-item ' + colorCls;
    item.textContent = label.substring(0, 50);
    item.addEventListener('click', () => scrollToRow(r));
    outline.appendChild(item);
  });
}

let pinnedTooltip = null;
function showTooltipAt(tr) {
  const api = tr.dataset.api;
  if (!api) return;
  const info = apiInfo[api];
  const ev = evidenceInfo[api];
  let html = '<div class="tt-api">' + api + '</div>';
  if (ev) {
    html += '<div class="tt-evidence">[' + ev.basis + '] ' + ev.note + '</div>';
  }
  if (info) {
    html += '<div class="tt-alt">Alternative: ' + info.alt + '</div>';
    if (info.note) html += '<div class="tt-note">' + info.note + '</div>';
    if (info.ex) html += '<pre><button class="tt-copy" onclick="copyPre(this)" title="Copy"><svg viewBox="0 0 24 24"><path d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/></svg></button>' + info.ex.replace(/\\n/g, '\n') + '</pre>';
  } else if (!ev) {
    return;
  }
  tooltip.innerHTML = html;
  tooltip.style.display = 'block';
  const rect = tr.getBoundingClientRect();
  let top = rect.bottom + 4;
  let left = rect.left + 60;
  if (top + tooltip.offsetHeight > window.innerHeight) top = rect.top - tooltip.offsetHeight - 4;
  if (left + tooltip.offsetWidth > window.innerWidth - 270) left = window.innerWidth - 270 - tooltip.offsetWidth - 10;
  tooltip.style.top = top + 'px';
  tooltip.style.left = left + 'px';
  pinnedTooltip = tr;
}
function copyPre(btn) {
  const pre = btn.closest('pre');
  const text = pre.textContent.trim();
  navigator.clipboard.writeText(text).then(() => {
    btn.innerHTML = '<svg viewBox="0 0 24 24"><path d="M9 16.17L4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z" fill="#6a9955"/></svg>';
    setTimeout(() => { btn.innerHTML = '<svg viewBox="0 0 24 24"><path d="M16 1H4c-1.1 0-2 .9-2 2v14h2V3h12V1zm3 4H8c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h11c1.1 0 2-.9 2-2V7c0-1.1-.9-2-2-2zm0 16H8V7h11v14z"/></svg>'; }, 1500);
  });
}

content.addEventListener('mousemove', (e) => {
  const tr = e.target.closest('tr.hl-edr, tr.hl-compat, tr.hl-path');
  if (tr && tr.dataset.api && !pinnedTooltip) {
    hoverHint.style.display = 'block';
    hoverHint.style.left = (e.clientX + 12) + 'px';
    hoverHint.style.top = (e.clientY - 8) + 'px';
  } else {
    hoverHint.style.display = 'none';
  }
});
content.addEventListener('mouseleave', () => { hoverHint.style.display = 'none'; });
content.addEventListener('click', (e) => {
  hoverHint.style.display = 'none';
  const tr = e.target.closest('tr.hl-edr, tr.hl-compat, tr.hl-path');
  if (!tr) { tooltip.style.display = 'none'; pinnedTooltip = null; return; }
  if (pinnedTooltip === tr) {
    tooltip.style.display = 'none'; pinnedTooltip = null;
  } else {
    showTooltipAt(tr);
  }
});
"@

    $htmlSubtitle = "$FileName -- $($CsvRow.EdrHits) EDR, $($CsvRow.CompatHits) compat, $($CsvRow.PathHits) path"
    $htmlPath = Join-Path $OutDir "${OutPrefix}_analyze.html"

    New-HtmlBase -Title "VBA Screening: $FileName" -Subtitle $htmlSubtitle `
        -ExtraCss $analyzeCss -SidebarHtml $sidebarSb.ToString() -ContentHtml $contentSb.ToString() `
        -ExtraHtml $extraHtml -ExtraJs $analyzeJs `
        -HighlightSelector 'tr.hl-edr, tr.hl-compat, tr.hl-path' `
        -FirstTabIndex $firstHlIdx -OutputPath $htmlPath

    return $htmlPath
}

# ============================================================================
# Mode 1: No args -> Settings GUI (3-axis)
# ============================================================================

if (-not $Paths -or $Paths.Count -eq 0) {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $cfg = Load-AnalyzeConfig $configPath

    $form = New-Object System.Windows.Forms.Form
    $form.Text = 'Analyze Settings'
    $form.Size = New-Object System.Drawing.Size(600, 700)
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MaximizeBox = $false
    $form.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#1e1e1e')
    $form.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#d4d4d4')
    $form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

    $lblTitle = New-Object System.Windows.Forms.Label
    $lblTitle.Text = 'Select patterns to detect'
    $lblTitle.Location = New-Object System.Drawing.Point(20, 12)
    $lblTitle.Size = New-Object System.Drawing.Size(400, 22)
    $lblTitle.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#cccccc')
    $lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 10)
    $form.Controls.Add($lblTitle)

    $scrollPanel = New-Object System.Windows.Forms.Panel
    $scrollPanel.Location = New-Object System.Drawing.Point(0, 40)
    $scrollPanel.Size = New-Object System.Drawing.Size(584, 570)
    $scrollPanel.AutoScroll = $true
    $scrollPanel.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#1e1e1e')
    $form.Controls.Add($scrollPanel)

    $lblDetectHdr = New-Object System.Windows.Forms.Label
    $lblDetectHdr.Text = 'Detect'
    $lblDetectHdr.Location = New-Object System.Drawing.Point(370, 4)
    $lblDetectHdr.Size = New-Object System.Drawing.Size(50, 16)
    $lblDetectHdr.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#888888')
    $lblDetectHdr.Font = New-Object System.Drawing.Font('Segoe UI', 8)
    $scrollPanel.Controls.Add($lblDetectHdr)

    # 3 groups: EDR, Compat, Path
    $groupDefs = @(
        @{ Key = 'edr';    Title = 'EDR Risks (probe-BLOCKED)';    Accent = '#4a9eff' }
        @{ Key = 'compat'; Title = 'Compatibility Risks';           Accent = '#9a5eff' }
        @{ Key = 'path';   Title = 'Hardcoded Path Risks';          Accent = '#4aff9e' }
    )

    $allControls = @{}
    $curY = 24

    foreach ($grp in $groupDefs) {
        $section = $grp.Key
        $allControls[$section] = [ordered]@{}

        if ($curY -gt 24) {
            $sep = New-Object System.Windows.Forms.Label
            $sep.Location = New-Object System.Drawing.Point(20, $curY)
            $sep.Size = New-Object System.Drawing.Size(520, 1)
            $sep.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#3c3c3c')
            $scrollPanel.Controls.Add($sep)
            $curY += 10
        }

        $hdr = New-Object System.Windows.Forms.Label
        $hdr.Text = $grp.Title
        $hdr.Location = New-Object System.Drawing.Point(20, $curY)
        $hdr.Size = New-Object System.Drawing.Size(300, 22)
        $hdr.ForeColor = [System.Drawing.ColorTranslator]::FromHtml($grp.Accent)
        $hdr.Font = New-Object System.Drawing.Font('Segoe UI', 9.5, [System.Drawing.FontStyle]::Bold)
        $scrollPanel.Controls.Add($hdr)
        $curY += 28

        foreach ($name in $cfg.$section.Keys) {
            $lbl = New-Object System.Windows.Forms.Label
            $lbl.Text = $name
            $lbl.Location = New-Object System.Drawing.Point(32, ($curY + 2))
            $lbl.Size = New-Object System.Drawing.Size(320, 20)
            $lbl.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#d4d4d4')
            $scrollPanel.Controls.Add($lbl)

            $cbD = New-Object System.Windows.Forms.CheckBox
            $cbD.Location = New-Object System.Drawing.Point(380, $curY)
            $cbD.Size = New-Object System.Drawing.Size(20, 20)
            $cbD.Checked = $cfg.$section[$name].detect
            $cbD.FlatStyle = 'Flat'
            $scrollPanel.Controls.Add($cbD)

            $allControls[$section][$name] = @{ Detect = $cbD }
            $curY += 26
        }
        $curY += 8
    }

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = 'OK'
    $btnOk.Size = New-Object System.Drawing.Size(100, 34)
    $btnOk.Location = New-Object System.Drawing.Point(370, 620)
    $btnOk.FlatStyle = 'Flat'
    $btnOk.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#0e639c')
    $btnOk.ForeColor = [System.Drawing.Color]::White
    $btnOk.Font = New-Object System.Drawing.Font('Segoe UI', 9.5)
    $btnOk.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($btnOk)
    $form.AcceptButton = $btnOk

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = 'Cancel'
    $btnCancel.Size = New-Object System.Drawing.Size(100, 34)
    $btnCancel.Location = New-Object System.Drawing.Point(480, 620)
    $btnCancel.FlatStyle = 'Flat'
    $btnCancel.BackColor = [System.Drawing.ColorTranslator]::FromHtml('#3c3c3c')
    $btnCancel.ForeColor = [System.Drawing.ColorTranslator]::FromHtml('#d4d4d4')
    $btnCancel.Font = New-Object System.Drawing.Font('Segoe UI', 9.5)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)
    $form.CancelButton = $btnCancel

    $result = $form.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $newCfg = @{ edr = [ordered]@{}; compat = [ordered]@{}; path = [ordered]@{} }
        foreach ($section in @('edr','compat','path')) {
            foreach ($name in $allControls[$section].Keys) {
                $ctrl = $allControls[$section][$name]
                $newCfg.$section[$name] = @{ detect = $ctrl.Detect.Checked }
            }
        }
        Save-AnalyzeConfig $newCfg $configPath
        Write-Host "Settings saved to $configPath" -ForegroundColor Green
    } else {
        Write-Host "Cancelled." -ForegroundColor Yellow
    }
    $form.Dispose()
    exit 0
}

# ============================================================================
# Mode 2/3: File/Folder analysis
# ============================================================================

$sw = [System.Diagnostics.Stopwatch]::StartNew()

# Collect all xlsm/xlam/xls files
$files = [System.Collections.ArrayList]::new()
$baseDir = $null

foreach ($p in $Paths) {
    $p = $p.Trim().Trim('"')
    $resolved = (Resolve-Path -LiteralPath $p -ErrorAction SilentlyContinue).Path
    if (-not $resolved) { Write-VbaError 'Analyze' $p 'Path not found'; continue }

    if (Test-Path -LiteralPath $resolved -PathType Container) {
        if (-not $baseDir) { $baseDir = $resolved }
        Get-ChildItem ([WildcardPattern]::Escape($resolved)) -Recurse -File -Include '*.xlsm','*.xlam','*.xls' | Where-Object {
            $_.FullName -notmatch '[\\/](output|debug_output)[\\/]'
        } | ForEach-Object {
            [void]$files.Add($_.FullName)
        }
    } else {
        $ext = [IO.Path]::GetExtension($resolved).ToLower()
        if ($ext -in '.xls','.xlsm','.xlam') {
            if (-not $baseDir) { $baseDir = [IO.Path]::GetDirectoryName($resolved) }
            [void]$files.Add($resolved)
        }
    }
}

if ($files.Count -eq 0) {
    Write-Host "No Excel files found." -ForegroundColor Yellow
    exit 0
}

if (-not $baseDir) { $baseDir = [IO.Path]::GetDirectoryName($files[0]) }

Write-VbaLog 'Analyze' $baseDir "=== Analyze session started: $($files.Count) files from $baseDir ==="

# Load config
Write-VbaLog 'Analyze' $configPath "Loading config"
$cfg = Load-AnalyzeConfig $configPath

# Build active rules from VBAToolkit analysis patterns, mapped to 3 axes
$patternDefs = Get-VbaAnalysis -Project @{ Modules = [ordered]@{}; Ole2 = $null }
$detectRules = [System.Collections.ArrayList]::new()
$detectRulesByName = @{}

# EDR patterns from VBAToolkit (old Patterns = edr axis in old code)
foreach ($name in $patternDefs.Patterns.Keys) {
    $axis = Get-NewAxis $name
    if (-not $axis) { continue }
    $cfgSection = $cfg[$axis]
    if (-not $cfgSection) { continue }
    $patCfg = $cfgSection[$name]
    if ($patCfg -and $patCfg.detect) {
        $rule = @{ Name = $name; Pattern = $patternDefs.Patterns[$name].Pattern; Category = $axis }
        [void]$detectRules.Add($rule)
        $detectRulesByName[$name] = $rule
    }
}

# Compat patterns from VBAToolkit
foreach ($name in $patternDefs.CompatPatterns.Keys) {
    $axis = Get-NewAxis $name
    if (-not $axis) { continue }
    $cfgSection = $cfg[$axis]
    if (-not $cfgSection) { continue }
    $patCfg = $cfgSection[$name]
    if ($patCfg -and $patCfg.detect) {
        $rule = @{ Name = $name; Pattern = $patternDefs.CompatPatterns[$name].Pattern; Category = $axis }
        [void]$detectRules.Add($rule)
        $detectRulesByName[$name] = $rule
    }
}

# Env patterns from VBAToolkit -> path axis
foreach ($name in $patternDefs.EnvPatterns.Keys) {
    $axis = Get-NewAxis $name
    if (-not $axis) { continue }
    $cfgSection = $cfg[$axis]
    if (-not $cfgSection) { continue }
    $patCfg = $cfgSection[$name]
    if ($patCfg -and $patCfg.detect) {
        $rule = @{ Name = $name; Pattern = $patternDefs.EnvPatterns[$name].Pattern; Category = $axis }
        [void]$detectRules.Add($rule)
        $detectRulesByName[$name] = $rule
    }
}

# EnvInfo patterns from VBAToolkit -> path axis
foreach ($name in $patternDefs.EnvInfoPatterns.Keys) {
    $axis = Get-NewAxis $name
    if (-not $axis) { continue }
    $cfgSection = $cfg[$axis]
    if (-not $cfgSection) { continue }
    $patCfg = $cfgSection[$name]
    if ($patCfg -and $patCfg.detect) {
        $rule = @{ Name = $name; Pattern = $patternDefs.EnvInfoPatterns[$name].Pattern; Category = $axis }
        [void]$detectRules.Add($rule)
        $detectRulesByName[$name] = $rule
    }
}

# Biz patterns from VBAToolkit (only those mapped to a new axis)
foreach ($name in $patternDefs.BizPatterns.Keys) {
    $axis = Get-NewAxis $name
    if (-not $axis) { continue }
    $cfgSection = $cfg[$axis]
    if (-not $cfgSection) { continue }
    $patCfg = $cfgSection[$name]
    if ($patCfg -and $patCfg.detect) {
        $rule = @{ Name = $name; Pattern = $patternDefs.BizPatterns[$name].Pattern; Category = $axis }
        [void]$detectRules.Add($rule)
        $detectRulesByName[$name] = $rule
    }
}

# Create output directory
$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$devkitRoot = Split-Path "$PSScriptRoot" -Parent
$outputRoot = Join-Path $devkitRoot 'output'
$outDir = Join-Path $outputRoot "${timestamp}_analyze"
[void][IO.Directory]::CreateDirectory($outDir)

# Detect filename collisions
$fileNameCounts = @{}
foreach ($f in $files) {
    $fn = [IO.Path]::GetFileNameWithoutExtension($f)
    if ($fileNameCounts.ContainsKey($fn)) { $fileNameCounts[$fn]++ } else { $fileNameCounts[$fn] = 1 }
}

Write-VbaLog 'Analyze' $baseDir "Rules loaded: detect=$($detectRules.Count)"
Write-VbaLog 'Analyze' $baseDir "Detect rules: $($detectRulesByName.Keys -join '; ')"
Write-VbaLog 'Analyze' $baseDir "Output dir: $outDir"

$replacements = Get-VbaApiReplacements
$csvRows = [System.Collections.ArrayList]::new()
$allHits = [System.Collections.ArrayList]::new()
$processed = 0

# CSV column definitions
$csvColumnNames = @('Timestamp','RelativePath','FileName','Bas','Cls','Frm','TotalModules','CodeLines','EdrHits','CompatHits','PathHits','RiskLevel','MigrationClass','Evidence','ReviewNote')
$hitsCsvColumns = @('MacroName','Category','PatternName','Module','LineNum','CodeSnippet','EvidenceBasis')

foreach ($filePath in $files) {
    $processed++
    $fileTimer = [System.Diagnostics.Stopwatch]::StartNew()
    $fileName = [IO.Path]::GetFileName($filePath)
    $baseName = [IO.Path]::GetFileNameWithoutExtension($filePath).Trim()
    $fileExt = [IO.Path]::GetExtension($filePath)

    # Determine output prefix — add subfolder name when file is not directly under baseDir
    $outPrefix = $baseName
    $fileDir = [IO.Path]::GetDirectoryName($filePath)
    if ($fileDir -ne $baseDir) {
        $relDir = ''
        if ($filePath.StartsWith($baseDir)) {
            $relDir = [IO.Path]::GetDirectoryName($filePath.Substring($baseDir.Length).TrimStart('\', '/'))
        }
        if ($relDir) {
            $outPrefix = ($relDir -replace '[\\/]', '_') + "_$baseName"
        } else {
            $parentDir = Split-Path (Split-Path "$filePath" -Parent) -Leaf
            $outPrefix = "${parentDir}_${baseName}"
        }
    }

    # Relative path from base
    $relPath = $filePath
    if ($filePath.StartsWith($baseDir)) {
        $relPath = $filePath.Substring($baseDir.Length).TrimStart('\', '/')
    }

    Write-VbaHeader 'Analyze' $fileName
    Write-VbaLog 'Analyze' $filePath "Processing started (file $processed of $($files.Count))"

    try {
        # Load project
        Write-VbaLog 'Analyze' $filePath "Loading project (mode=StripAttributes)"
        $project = Get-AllModuleCode $filePath -StripAttributes
        if (-not $project) {
            $errorRow = [ordered]@{}
            foreach ($col in $csvColumnNames) { $errorRow[$col] = '' }
            $errorRow.Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            $errorRow.RelativePath = [IO.Path]::GetDirectoryName($relPath)
            $errorRow.FileName = $fileName
            $errorRow.Error = 'No VBA project'
            [void]$csvRows.Add($errorRow)
            Write-VbaError 'Analyze' $fileName 'No vbaProject.bin found'
            continue
        }

        # Get old 4-axis analysis from VBAToolkit
        $analysis = Get-VbaAnalysis -Project $project

        # Remap to 3-axis
        $remapped = Remap-AnalysisTo3Axis -Analysis $analysis

        Write-VbaStatus 'Analyze' $fileName "Modules: $($project.Modules.Count)"
        Write-VbaStatus 'Analyze' $fileName "EDR hits: $($remapped.EdrHits)"
        Write-VbaStatus 'Analyze' $fileName "Compat hits: $($remapped.CompatHits)"
        Write-VbaStatus 'Analyze' $fileName "Path hits: $($remapped.PathHits)"
        Write-VbaLog 'Analyze' $filePath "Modules=$($project.Modules.Count) EDR=$($remapped.EdrHits) Compat=$($remapped.CompatHits) Path=$($remapped.PathHits)"
        Write-VbaLog 'Analyze' $filePath "Module names: $($project.Modules.Keys -join ', ')"
        if ($project.Codepage) { Write-VbaLog 'Analyze' $filePath "Codepage=$($project.Codepage)" }
        if ($project.Ole2Bytes) { Write-VbaLog 'Analyze' $filePath "Ole2Bytes size=$($project.Ole2Bytes.Length)" }

        # === Per-macro subfolder ===
        $macroSubDir = Join-Path $outDir $outPrefix
        Write-VbaLog 'Analyze' $filePath "Creating subfolder: $macroSubDir"
        [void][IO.Directory]::CreateDirectory($macroSubDir)

        # === Build hits list ===
        Write-VbaLog 'Analyze' $filePath "Building hits list (detectRules=$($detectRulesByName.Count))"
        $fileHits = Build-HitsList -Project $project -Remapped $remapped -FileName $fileName -DetectRulesByName $detectRulesByName
        Write-VbaLog 'Analyze' $filePath "Hits found: $($fileHits.Count)"
        if ($fileHits.Count -gt 0) {
            $hitSummary = ($fileHits | Group-Object { $_['Category'] } | ForEach-Object { "$($_.Name):$($_.Count)" }) -join ' '
            Write-VbaLog 'Analyze' $filePath "Hit breakdown: $hitSummary"
        }
        foreach ($hit in $fileHits) { [void]$allHits.Add($hit) }

        # === Build CSV row ===
        $csvRow = Build-AnalyzeCsvRow -Analysis $analysis -Remapped $remapped -Project $project -FileName $fileName -RelPath $relPath -Hits $fileHits

        # === Build line-level highlight data for HTML ===
        $allModLines = [ordered]@{}
        foreach ($modName in $project.Modules.Keys) {
            $mod = $project.Modules[$modName]
            $lines = if ($mod.Lines -is [array]) { $mod.Lines } else { @($mod.Lines) }
            $displayLines = @($lines | Where-Object { $_ -notmatch '^\s*Attribute\s+VB_' })
            $allModLines[$modName] = @{ Ext = $mod.Ext; Lines = $displayLines }
        }

        # Build highlight map per module (3 colors)
        $modHighlights = @{}
        foreach ($modName in $allModLines.Keys) {
            $ml = $allModLines[$modName]
            $hlMap = @{}
            for ($i = 0; $i -lt $ml.Lines.Count; $i++) {
                $line = $ml.Lines[$i]

                if ($line -match '^\s*''') { continue }

                $foundEdr = $null; $foundCompat = $null; $foundPath = $null
                foreach ($rule in $detectRules) {
                    if ($line -match $rule.Pattern) {
                        switch ($rule.Category) {
                            'edr'    { if (-not $foundEdr) { $foundEdr = $rule } }
                            'compat' { if (-not $foundCompat) { $foundCompat = $rule } }
                            'path'   { if (-not $foundPath) { $foundPath = $rule } }
                        }
                    }
                }

                # Priority: edr > compat > path (+ API call names)
                if ($foundEdr) {
                    $hlMap[$i] = @{ Color = 'hl-edr'; Category = 'edr'; PatternName = $foundEdr.Name }
                    continue
                }

                $apiMatched = $false
                foreach ($apiName in $analysis.ApiCallNames) {
                    if ($line -match "\b$([regex]::Escape($apiName))\b") {
                        $hlMap[$i] = @{ Color = 'hl-edr'; Category = 'edr'; PatternName = "API: $apiName" }
                        $apiMatched = $true
                        break
                    }
                }
                if ($apiMatched) { continue }

                if ($foundCompat) {
                    $hlMap[$i] = @{ Color = 'hl-compat'; Category = 'compat'; PatternName = $foundCompat.Name }
                    continue
                }
                if ($foundPath) {
                    $hlMap[$i] = @{ Color = 'hl-path'; Category = 'path'; PatternName = $foundPath.Name }
                    continue
                }
            }
            $modHighlights[$modName] = $hlMap
        }

        # === Generate _analyze.txt in subfolder ===
        $reportText = Build-AnalyzeTextReport -Analysis $analysis -Remapped $remapped -AllModLines $allModLines -CsvRow $csvRow -FileName $fileName -Replacements $replacements
        $txtPath = Join-Path $macroSubDir "${outPrefix}_analyze.txt"
        Write-VbaLog 'Analyze' $filePath "Writing text report: $txtPath"
        [IO.File]::WriteAllText($txtPath, $reportText, [System.Text.Encoding]::UTF8)

        # === Generate _combined.txt (integrated source with summary header) ===
        $combinedSb = [System.Text.StringBuilder]::new()
        [void]$combinedSb.AppendLine("=" * 80)
        [void]$combinedSb.AppendLine(" $fileName - VBA Source Code")
        [void]$combinedSb.AppendLine(" Analyzed: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
        [void]$combinedSb.AppendLine(" EDR: $($remapped.EdrHits)  Compat: $($remapped.CompatHits)  Path: $($remapped.PathHits)  Risk: $($csvRow.RiskLevel)")
        [void]$combinedSb.AppendLine("=" * 80)
        [void]$combinedSb.AppendLine("")

        # Module index
        [void]$combinedSb.AppendLine("MODULE INDEX")
        [void]$combinedSb.AppendLine("-" * 40)
        [void]$combinedSb.AppendLine("")
        $cStdMods = @(); $cClsMods = @(); $cFrmMods = @(); $cDocMods = @()
        $cTotalLines = 0
        foreach ($modName in $allModLines.Keys) {
            $ml = $allModLines[$modName]
            $lc = $ml.Lines.Count
            $cTotalLines += $lc
            $entry = "    $modName.$($ml.Ext) ($lc lines)"
            switch ($ml.Ext) {
                'bas' { $cStdMods += $entry }
                'cls' { $cClsMods += $entry }
                'frm' { $cFrmMods += $entry }
                default { $cDocMods += $entry }
            }
        }
        if ($cStdMods.Count -gt 0) { [void]$combinedSb.AppendLine("  Standard Modules:"); foreach ($e in $cStdMods) { [void]$combinedSb.AppendLine($e) }; [void]$combinedSb.AppendLine("") }
        if ($cClsMods.Count -gt 0) { [void]$combinedSb.AppendLine("  Class Modules:"); foreach ($e in $cClsMods) { [void]$combinedSb.AppendLine($e) }; [void]$combinedSb.AppendLine("") }
        if ($cFrmMods.Count -gt 0) { [void]$combinedSb.AppendLine("  UserForms:"); foreach ($e in $cFrmMods) { [void]$combinedSb.AppendLine($e) }; [void]$combinedSb.AppendLine("") }
        if ($cDocMods.Count -gt 0) { [void]$combinedSb.AppendLine("  Document Modules:"); foreach ($e in $cDocMods) { [void]$combinedSb.AppendLine($e) }; [void]$combinedSb.AppendLine("") }
        [void]$combinedSb.AppendLine("  Total: $cTotalLines lines across $($allModLines.Count) modules")
        [void]$combinedSb.AppendLine("")

        # Ordered module source (bas -> cls -> frm -> others)
        $extOrder = @('bas','cls','frm')
        $sortedModNames = $allModLines.Keys | Sort-Object { $o = [Array]::IndexOf($extOrder, $allModLines[$_].Ext); if($o -lt 0){99}else{$o} }, { $_ }
        foreach ($modName in $sortedModNames) {
            $ml = $allModLines[$modName]
            [void]$combinedSb.AppendLine("=" * 80)
            [void]$combinedSb.AppendLine(" $modName.$($ml.Ext)")
            [void]$combinedSb.AppendLine("=" * 80)
            [void]$combinedSb.AppendLine("")
            [void]$combinedSb.AppendLine(($ml.Lines -join "`r`n").TrimStart("`r`n"))
            [void]$combinedSb.AppendLine("")
        }
        $combinedPath = Join-Path $macroSubDir "${outPrefix}_combined.txt"
        Write-VbaLog 'Analyze' $filePath "Writing combined source: $combinedPath ($cTotalLines lines)"
        [IO.File]::WriteAllText($combinedPath, $combinedSb.ToString(), [System.Text.Encoding]::UTF8)

        # === Build deduplicated tooltip entries ===
        $tooltipEntries = [ordered]@{}
        foreach ($apiName in @($analysis.ApiCallNames) + @($analysis.ApiDecls | ForEach-Object { $_.Name })) {
            $info = $replacements[$apiName]
            if (-not $info) { continue }
            $key = "API: $apiName"
            if (-not $tooltipEntries.Contains($key)) {
                $tooltipEntries[$key] = $info
            }
        }
        foreach ($patName in @($patternDefs.Patterns.Keys) + @($patternDefs.CompatPatterns.Keys) + @($patternDefs.EnvPatterns.Keys) + @($patternDefs.EnvInfoPatterns.Keys) + @($patternDefs.BizPatterns.Keys)) {
            $info = $replacements[$patName]
            if (-not $info) { continue }
            if (-not $tooltipEntries.Contains($patName)) {
                $tooltipEntries[$patName] = $info
            }
        }

        # === Generate _analyze.html in subfolder ===
        Write-VbaLog 'Analyze' $filePath "Writing HTML report to $macroSubDir"
        $htmlPath = Build-AnalyzeHtml -AllModLines $allModLines -ModHighlights $modHighlights -TooltipEntries $tooltipEntries `
            -OutPrefix $outPrefix -FileName $fileName -CsvRow $csvRow -OutDir $macroSubDir -PatternDefs $patternDefs

        # Open HTML for single-file runs
        if ($files.Count -eq 1) { Start-Process "$htmlPath" }
        Write-VbaLog 'Analyze' $filePath "Processing complete"

    } catch {
        $csvRow = [ordered]@{}
        foreach ($col in $csvColumnNames) { $csvRow[$col] = '' }
        $csvRow.Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $csvRow.RelativePath = [IO.Path]::GetDirectoryName($relPath)
        $csvRow.FileName = $fileName
        $csvRow.Error = $_.Exception.Message
        Write-VbaError 'Analyze' $fileName $_.Exception.Message
        Write-VbaLog 'Analyze' $filePath "EXCEPTION: $($_.Exception.Message)" 'ERROR'
        Write-VbaLog 'Analyze' $filePath "Stack: $($_.ScriptStackTrace)" 'ERROR'
    }

    [void]$csvRows.Add($csvRow)

    $fileSw = $fileTimer.Elapsed.TotalSeconds
    Write-VbaResult 'Analyze' $fileName "$($csvRow.EdrHits) EDR, $($csvRow.CompatHits) compat, $($csvRow.PathHits) path" $outDir $fileSw
    Write-VbaLog 'Analyze' $filePath "$($csvRow.TotalModules) modules, $($csvRow.EdrHits) EDR, $($csvRow.CompatHits) compat, $($csvRow.PathHits) path | -> $outDir"
}

# === Write analyze.csv (summary CSV) ===
Write-VbaLog 'Analyze' $outDir "Writing analyze.csv ($($csvRows.Count) rows)"
$csvPath = Join-Path $outDir 'analyze.csv'
$csvSb = [System.Text.StringBuilder]::new()
[void]$csvSb.AppendLine($csvColumnNames -join ',')

foreach ($row in $csvRows) {
    $fields = foreach ($col in $csvColumnNames) {
        $val = $row[$col]
        if ($val -is [int] -or $val -is [long] -or $val -is [double]) {
            $val
        } else {
            '"' + ([string]$val -replace '"','""') + '"'
        }
    }
    [void]$csvSb.AppendLine($fields -join ',')
}

$utf8Bom = New-Object System.Text.UTF8Encoding $true
[IO.File]::WriteAllText($csvPath, $csvSb.ToString(), $utf8Bom)

# === Write hits.csv (line-level findings CSV) ===
Write-VbaLog 'Analyze' $outDir "Writing hits.csv ($($allHits.Count) total hits)"
$hitsCsvPath = Join-Path $outDir 'hits.csv'
$hitsSb = [System.Text.StringBuilder]::new()
[void]$hitsSb.AppendLine($hitsCsvColumns -join ',')

foreach ($hit in $allHits) {
    $fields = foreach ($col in $hitsCsvColumns) {
        $val = $hit[$col]
        if ($val -is [int] -or $val -is [long] -or $val -is [double]) {
            $val
        } else {
            '"' + ([string]$val -replace '"','""') + '"'
        }
    }
    [void]$hitsSb.AppendLine($fields -join ',')
}

[IO.File]::WriteAllText($hitsCsvPath, $hitsSb.ToString(), $utf8Bom)

$sw.Stop()
Write-VbaLog 'Analyze' $outDir "=== Analyze session complete: $($files.Count) files, $($allHits.Count) hits, $([Math]::Round($sw.Elapsed.TotalSeconds, 1))s ==="
if ($files.Count -gt 1) {
    Write-Host "`n  Total: $($files.Count) files analyzed" -ForegroundColor Green
}
Write-Host "  Summary CSV: $csvPath" -ForegroundColor Gray
Write-Host "  Hits CSV: $hitsCsvPath" -ForegroundColor Gray
Write-Host "  Output: $outDir" -ForegroundColor Gray
Write-Host "  Log: $(Join-Path (Split-Path "$PSScriptRoot" -Parent) 'vba-toolkit.log')" -ForegroundColor Gray
Write-Host "  Done ($([Math]::Round($sw.Elapsed.TotalSeconds, 1))s)" -ForegroundColor Gray
