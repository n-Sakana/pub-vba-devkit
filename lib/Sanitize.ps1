param(
    [Parameter(Mandatory = $true)]
    [string[]]$Path
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot\VBAToolkit.psm1" -Force -DisableNameChecking

$script:SupportedExtensions = @('.xls', '.xlsm', '.xlam')
$script:EdrRuleNames = @(
    'Win32 API (Declare)',
    'Shell / process',
    'PowerShell / WScript'
)
$script:SafeFallbackLine = "' *** disabled: unsafe statement"
$script:SafeContinuationLine = "' *** disabled: continued unsafe statement"

function Get-SanitizeTargets {
    param([string[]]$InputPaths)

    $seen = New-Object 'System.Collections.Hashtable' ([System.StringComparer]::OrdinalIgnoreCase)
    $targets = [System.Collections.ArrayList]::new()

    foreach ($item in $InputPaths) {
        $resolved = Resolve-Path -LiteralPath $item -ErrorAction Stop
        foreach ($rp in $resolved) {
            $fsPath = $rp.ProviderPath
            if ([IO.Directory]::Exists($fsPath)) {
                $files = Get-ChildItem -LiteralPath $fsPath -Recurse -File |
                    Where-Object { $script:SupportedExtensions -contains $_.Extension.ToLowerInvariant() }
                foreach ($file in $files) {
                    if (-not $seen.ContainsKey($file.FullName)) {
                        $seen[$file.FullName] = $true
                        [void]$targets.Add($file.FullName)
                    }
                }
            } elseif ([IO.File]::Exists($fsPath)) {
                $ext = [IO.Path]::GetExtension($fsPath).ToLowerInvariant()
                if ($script:SupportedExtensions -contains $ext) {
                    if (-not $seen.ContainsKey($fsPath)) {
                        $seen[$fsPath] = $true
                        [void]$targets.Add($fsPath)
                    }
                } else {
                    Write-VbaLog 'Sanitize' $fsPath "SKIP: unsupported extension $ext" 'WARN'
                }
            }
        }
    }

    return $targets.ToArray()
}

function Get-CommonBaseDirectory {
    param([string[]]$Files)

    if (-not $Files -or $Files.Count -eq 0) { return (Get-Location).Path }
    $dirs = @($Files | ForEach-Object { [IO.Path]::GetDirectoryName($_) })
    $base = $dirs[0]

    foreach ($dir in $dirs) {
        while ($base -and -not $dir.StartsWith($base, [System.StringComparison]::OrdinalIgnoreCase)) {
            $parent = [IO.Directory]::GetParent($base)
            if (-not $parent) { return $base }
            $base = $parent.FullName
        }
    }
    return $base
}

function Get-RelativePathText {
    param([string]$BaseDir, [string]$FilePath)

    if ($FilePath.StartsWith($BaseDir, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $FilePath.Substring($BaseDir.Length).TrimStart('\', '/')
    }
    return [IO.Path]::GetFileName($FilePath)
}

function New-SanitizedOutputPath {
    param(
        [string]$OutDir,
        [string]$BaseDir,
        [string]$FilePath,
        [hashtable]$UsedNames
    )

    $baseName = [IO.Path]::GetFileNameWithoutExtension($FilePath).Trim()
    $ext = [IO.Path]::GetExtension($FilePath)
    $outStem = $baseName
    $fileDir = [IO.Path]::GetDirectoryName($FilePath)

    if ($fileDir -and $BaseDir -and -not $fileDir.Equals($BaseDir, [System.StringComparison]::OrdinalIgnoreCase)) {
        $rel = Get-RelativePathText $BaseDir $FilePath
        $relDir = [IO.Path]::GetDirectoryName($rel)
        if ($relDir) {
            $prefix = $relDir -replace '[\\/]', '_'
            $outStem = "${prefix}_$baseName"
        }
    }

    $candidate = "${outStem}_sanitized$ext"
    $n = 2
    while ($UsedNames.ContainsKey($candidate)) {
        $candidate = "${outStem}_sanitized_$n$ext"
        $n++
    }
    $UsedNames[$candidate] = $true
    return (Join-Path $OutDir $candidate)
}

function Get-EdrRules {
    $defs = Get-VbaAnalysis -Project @{ Modules = [ordered]@{}; Ole2 = $null }
    $rules = [System.Collections.ArrayList]::new()

    foreach ($name in $script:EdrRuleNames) {
        $def = $defs.Patterns[$name]
        if (-not $def) { throw "Missing EDR rule from VBAToolkit: $name" }
        [void]$rules.Add([ordered]@{
            Name = $name
            Pattern = $def.Pattern
        })
    }
    return ,$rules
}

function Test-VbaLineContinues {
    param([string]$Line)
    if ($null -eq $Line) { return $false }
    return ($Line -match '(^|[ \t])_\s*(?:''.*)?$')
}

function Get-VbaStatementGroups {
    param([string[]]$Lines)

    $groups = [System.Collections.ArrayList]::new()
    if (-not $Lines) { return ,$groups }

    $start = 0
    for ($i = 0; $i -lt $Lines.Count; $i++) {
        if (Test-VbaLineContinues $Lines[$i]) { continue }
        [void]$groups.Add([ordered]@{
            Start = $start
            End = $i
        })
        $start = $i + 1
    }

    if ($start -lt $Lines.Count) {
        [void]$groups.Add([ordered]@{
            Start = $start
            End = $Lines.Count - 1
        })
    }
    return ,$groups
}

function Get-StatementText {
    param(
        [string[]]$Lines,
        [hashtable]$Group
    )

    $parts = [System.Collections.ArrayList]::new()
    for ($i = $Group.Start; $i -le $Group.End; $i++) {
        [void]$parts.Add($Lines[$i])
    }
    return ($parts -join "`n")
}

function Find-EdrRuleHit {
    param(
        [string]$StatementText,
        [System.Collections.IEnumerable]$Rules
    )

    foreach ($rule in $Rules) {
        if ([regex]::IsMatch($StatementText, $rule.Pattern)) {
            return $rule.Name
        }
    }
    return $null
}

function Get-FlatStatementText {
    param([string]$StatementText)
    $flat = $StatementText -replace "[ \t]_\s*(`r?`n)", ' '
    $flat = $flat -replace "(`r?`n)", ' '
    return ($flat -replace '\s+', ' ').Trim()
}

function Get-DeclareMetadata {
    param([string]$StatementText)

    $flat = Get-FlatStatementText $StatementText
    $meta = [ordered]@{
        Kind = 'api-decl'
        Names = @()
    }

    if ($flat -match '(?i)\bDeclare\s+(?:PtrSafe\s+)?(Function|Sub)\s+([A-Za-z_][A-Za-z0-9_]*)\b') {
        $meta.Names = @($Matches[2])
    }

    $aliases = [System.Collections.ArrayList]::new()
    foreach ($m in [regex]::Matches($flat, '(?i)\bAlias\s+"([^"]+)"')) {
        [void]$aliases.Add($m.Groups[1].Value)
    }
    if ($aliases.Count -gt 0) {
        foreach ($alias in $aliases) {
            $meta.Names = @($meta.Names + $alias)
        }
    }

    return $meta
}

function Add-NameToSet {
    param(
        [hashtable]$Set,
        [string]$Name,
        $Metadata
    )

    if ([string]::IsNullOrWhiteSpace($Name)) { return }
    if ($Name -notmatch '^[A-Za-z_][A-Za-z0-9_]*$') { return }
    if (-not $Set.ContainsKey($Name)) { $Set[$Name] = $Metadata }
}

function Add-DeclareNames {
    param(
        [hashtable]$NameSet,
        $Metadata
    )

    if (-not $Metadata -or -not $Metadata.Names) { return }
    foreach ($name in $Metadata.Names) {
        Add-NameToSet $NameSet $name $Metadata
    }
}

function Get-RelevantMaskTokens {
    param(
        [string]$StatementText,
        $Metadata
    )

    $tokens = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($token in @(
        'Declare', 'PtrSafe', 'Lib', 'Alias',
        'Shell', 'WScript.Shell', 'WScript', 'PowerShell', 'powershell', 'pwsh',
        'cscript', 'wscript', 'mshta', 'cmd'
    )) {
        [void]$tokens.Add($token)
    }

    if ($Metadata -and $Metadata.Names) {
        foreach ($name in $Metadata.Names) { [void]$tokens.Add([string]$name) }
    }

    $flat = Get-FlatStatementText $StatementText
    foreach ($m in [regex]::Matches($flat, '(?i)\bLib\s+"([^"]+)"')) {
        $lib = $m.Groups[1].Value
        [void]$tokens.Add($lib)
        [void]$tokens.Add([IO.Path]::GetFileNameWithoutExtension($lib))
    }
    foreach ($m in [regex]::Matches($flat, '(?i)\bAlias\s+"([^"]+)"')) {
        [void]$tokens.Add($m.Groups[1].Value)
    }
    foreach ($m in [regex]::Matches($flat, '(?i)"([^"]*(?:\.exe|powershell|pwsh|cmd|wscript|cscript|mshta)[^"]*)"')) {
        foreach ($part in ($m.Groups[1].Value -split '[^A-Za-z0-9_.$-]+')) {
            if ($part.Length -gt 0) { [void]$tokens.Add($part) }
            $base = [IO.Path]::GetFileNameWithoutExtension($part)
            if ($base) { [void]$tokens.Add($base) }
        }
    }

    return @($tokens | Where-Object { $_ -and $_.Length -gt 0 } | Sort-Object Length -Descending)
}

function Mask-TokenText {
    param([string]$Token)

    if ([string]::IsNullOrEmpty($Token)) { return $Token }
    $len = $Token.Length
    if ($len -le 1) { return '*' }

    $left = [Math]::Min(3, [Math]::Max(1, $len - 3))
    $right = if ($len -gt 5) { 2 } else { 1 }
    if ($left + $right -ge $len) {
        $left = 1
        $right = 0
    }
    $middle = [Math]::Max(2, $len - $left - $right)
    return $Token.Substring(0, $left) + ('*' * $middle) + $(if ($right -gt 0) { $Token.Substring($len - $right) } else { '' })
}

function Get-MaskedStatementText {
    param(
        [string]$StatementText,
        $Metadata
    )

    $text = Get-FlatStatementText $StatementText
    foreach ($token in (Get-RelevantMaskTokens $text $Metadata)) {
        $replacement = Mask-TokenText $token
        $text = [regex]::Replace($text, [regex]::Escape($token), [System.Text.RegularExpressions.MatchEvaluator]{ param($m) $replacement }, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    }
    return $text
}

function Get-MaskedReplacementLine {
    param(
        [string]$Kind,
        [string]$StatementText,
        $Metadata
    )

    $masked = Get-MaskedStatementText $StatementText $Metadata
    return "' *** masked kind=${Kind}: $masked"
}

function New-ReplacementComment {
    param(
        [string]$Kind,
        [string]$StatementText,
        $Metadata
    )

    return Get-MaskedReplacementLine $Kind $StatementText $Metadata
}

function Get-CompactReplacementComment {
    param([string]$ReplacementComment)

    if ([string]::IsNullOrWhiteSpace($ReplacementComment)) {
        return $script:SafeFallbackLine
    }
    if ($ReplacementComment -match 'masked kind=([^:]+):') {
        return "' *** masked kind=$($Matches[1])"
    }
    return $script:SafeFallbackLine
}

function Remove-VbaCommentsAndStrings {
    param([string]$Text)

    $sb = [System.Text.StringBuilder]::new()
    $inString = $false
    $i = 0

    while ($i -lt $Text.Length) {
        $ch = $Text[$i]

        if ($ch -eq '"') {
            if ($inString -and ($i + 1) -lt $Text.Length -and $Text[$i + 1] -eq '"') {
                [void]$sb.Append(' ')
                [void]$sb.Append(' ')
                $i += 2
                continue
            }
            $inString = -not $inString
            [void]$sb.Append(' ')
            $i++
            continue
        }

        if (-not $inString -and $ch -eq "'") {
            while ($i -lt $Text.Length -and $Text[$i] -ne "`n") {
                [void]$sb.Append(' ')
                $i++
            }
            continue
        }

        if ($inString) {
            if ($ch -eq "`r" -or $ch -eq "`n") {
                [void]$sb.Append($ch)
            } else {
                [void]$sb.Append(' ')
            }
        } else {
            [void]$sb.Append($ch)
        }
        $i++
    }

    return $sb.ToString()
}

function Find-ApiCallHit {
    param(
        [string]$StatementText,
        [hashtable]$NameSet
    )

    if (-not $NameSet -or $NameSet.Count -eq 0) { return $null }
    $searchText = Remove-VbaCommentsAndStrings $StatementText

    foreach ($name in $NameSet.Keys) {
        $escaped = [regex]::Escape([string]$name)
        if ([regex]::IsMatch($searchText, "(?i)(?<![A-Za-z0-9_])$escaped(?![A-Za-z0-9_])")) {
            return $NameSet[$name]
        }
    }
    return $null
}

function Get-Ole2StreamCapacity {
    param($Ole2, $Entry)

    if ($Entry.Size -lt $Ole2.MiniStreamCutoff) {
        $count = 0
        $s = $Entry.Start
        $visited = @{}
        while ($s -ge 0 -and $s -ne -2 -and -not $visited.ContainsKey($s)) {
            $visited[$s] = $true
            $count++
            if ($s -lt $Ole2.MiniFat.Length) { $s = $Ole2.MiniFat[$s] } else { break }
        }
        return ($count * $Ole2.MiniSectorSize)
    }

    $sectorCount = 0
    $sector = $Entry.Start
    $seen = @{}
    while ($sector -ge 0 -and $sector -ne -2 -and -not $seen.ContainsKey($sector)) {
        $seen[$sector] = $true
        $sectorCount++
        if ($sector -lt $Ole2.Fat.Length) { $sector = $Ole2.Fat[$sector] } else { break }
    }
    return ($sectorCount * $Ole2.SectorSize)
}

function New-ModuleStream {
    param(
        [string[]]$Lines,
        [System.Text.Encoding]$Encoding,
        [hashtable]$ModuleData,
        $Ole2
    )

    $text = $Lines -join "`r`n"
    $capacity = Get-Ole2StreamCapacity $Ole2 $ModuleData.Entry
    $minStreamSize = 0
    if ($ModuleData.Entry.Size -ge $Ole2.MiniStreamCutoff) {
        $minStreamSize = [int]$Ole2.MiniStreamCutoff
    }

    $attempt = 0
    while ($true) {
        $raw = $Encoding.GetBytes($text)
        $compressed = Compress-VBA $raw
        $streamLength = $ModuleData.Offset + $compressed.Length

        if (($minStreamSize -eq 0 -or $streamLength -ge $minStreamSize) -and $streamLength -le $capacity) {
            $newStream = New-Object byte[] $streamLength
            [Array]::Copy($ModuleData.StreamData, 0, $newStream, 0, $ModuleData.Offset)
            [Array]::Copy($compressed, 0, $newStream, $ModuleData.Offset, $compressed.Length)
            return ,$newStream
        }

        if ($streamLength -gt $capacity) {
            throw "Sanitized module stream exceeds existing OLE2 chain capacity. Stream=$streamLength Capacity=$capacity"
        }

        $attempt++
        $fill = "' *** pad $attempt " + ('x' * 96) + " $attempt"
        $text = $text + "`r`n" + $fill
    }
}

function New-ModulePlan {
    param(
        [string]$ModuleName,
        [hashtable]$Module,
        [System.Collections.IEnumerable]$Rules,
        [hashtable]$ApiNameSet
    )

    $lines = [string[]]@($Module.Lines)
    $groups = Get-VbaStatementGroups $lines
    $break = New-Object bool[] $groups.Count
    $replacement = New-Object string[] $groups.Count
    $directStatements = 0

    for ($i = 0; $i -lt $groups.Count; $i++) {
        $group = $groups[$i]
        $statement = Get-StatementText $lines $group
        $ruleName = Find-EdrRuleHit $statement $Rules
        if ($ruleName) {
            $break[$i] = $true
            $directStatements++
            if ($ruleName -eq 'Win32 API (Declare)') {
                $metadata = Get-DeclareMetadata $statement
                Add-DeclareNames $ApiNameSet $metadata
                $replacement[$i] = New-ReplacementComment 'api-decl' $statement $metadata
            } elseif ($ruleName -eq 'Shell / process') {
                $replacement[$i] = New-ReplacementComment 'process-launch' $statement $null
            } elseif ($ruleName -eq 'PowerShell / WScript') {
                $replacement[$i] = New-ReplacementComment 'script-host' $statement $null
            } else {
                $replacement[$i] = New-ReplacementComment 'unknown' $statement $null
            }
        }
    }

    return [ordered]@{
        ModuleName = $ModuleName
        Module = $Module
        Lines = $lines
        Groups = $groups
        Break = $break
        Replacement = $replacement
        DirectStatements = $directStatements
        ApiCallStatements = 0
        ChangedLines = 0
    }
}

function Complete-ModulePlan {
    param(
        $Plan,
        [hashtable]$ApiNameSet
    )

    if (-not $ApiNameSet -or $ApiNameSet.Count -eq 0) { return }

    for ($i = 0; $i -lt $Plan.Groups.Count; $i++) {
        if ($Plan.Break[$i]) { continue }
        $group = $Plan.Groups[$i]
        $statement = Get-StatementText $Plan.Lines $group
        $apiHit = Find-ApiCallHit $statement $ApiNameSet
        if ($apiHit) {
            $Plan.Break[$i] = $true
            $Plan['ApiCallStatements'] = [int]$Plan['ApiCallStatements'] + 1
            $Plan.Replacement[$i] = New-ReplacementComment 'api-call' $statement $apiHit
        }
    }
}

function Apply-ModulePlan {
    param(
        $Plan,
        [switch]$Compact
    )

    $newLines = [string[]]$Plan.Lines.Clone()
    $changed = 0
    for ($i = 0; $i -lt $Plan.Groups.Count; $i++) {
        if (-not $Plan.Break[$i]) { continue }
        $group = $Plan.Groups[$i]
        $firstLine = if ($Plan.Replacement[$i]) { $Plan.Replacement[$i] } else { $script:SafeFallbackLine }
        if ($Compact) {
            $firstLine = Get-CompactReplacementComment $firstLine
        }
        for ($lineNo = $group.Start; $lineNo -le $group.End; $lineNo++) {
            if ($lineNo -eq $group.Start) {
                $newLines[$lineNo] = $firstLine
            } else {
                $newLines[$lineNo] = $script:SafeContinuationLine
            }
            $changed++
        }
    }
    $Plan['ChangedLines'] = $changed
    return ,$newLines
}

function Assert-SanitizedSource {
    param(
        [hashtable]$Project,
        [System.Collections.IEnumerable]$Rules,
        [hashtable]$ApiNameSet
    )

    foreach ($moduleName in $Project.Modules.Keys) {
        $module = $Project.Modules[$moduleName]
        $lines = [string[]]@($module.Lines)
        $groups = Get-VbaStatementGroups $lines
        foreach ($group in $groups) {
            $statement = Get-StatementText $lines $group
            $ruleName = Find-EdrRuleHit $statement $Rules
            if ($ruleName) {
                throw "Verification failed: EDR rule remains in $moduleName"
            }
            if (Find-ApiCallHit $statement $ApiNameSet) {
                throw "Verification failed: declared API call remains in $moduleName"
            }
        }
    }
}

function Invoke-SanitizeFile {
    param(
        [string]$FilePath,
        [string]$OutputPath,
        [string]$BaseDir,
        [System.Collections.IEnumerable]$Rules
    )

    $row = [ordered]@{
        Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        RelativePath = Get-RelativePathText $BaseDir $FilePath
        FileName = [IO.Path]::GetFileName($FilePath)
        OutputFile = [IO.Path]::GetFileName($OutputPath)
        Status = ''
        Modules = 0
        DirectStatements = 0
        ApiCallStatements = 0
        ChangedLines = 0
        Error = ''
    }

    Copy-Item -LiteralPath $FilePath -Destination $OutputPath -Force

    $project = Get-AllModuleCode $OutputPath -IncludeRawData
    if (-not $project) {
        $row.Status = 'copied-no-vba'
        return $row
    }

    $row.Modules = $project.Modules.Count
    $encoding = [System.Text.Encoding]::GetEncoding($project.Codepage)
    $apiNames = New-Object 'System.Collections.Hashtable' ([System.StringComparer]::OrdinalIgnoreCase)
    $plans = [ordered]@{}

    foreach ($moduleName in $project.Modules.Keys) {
        $module = $project.Modules[$moduleName]
        if (-not $module.Entry -or $null -eq $module.Offset -or -not $module.StreamData) { continue }
        $plan = New-ModulePlan $moduleName $module $Rules $apiNames
        $plans[$moduleName] = $plan
        $row.DirectStatements += $plan.DirectStatements
    }

    foreach ($moduleName in $plans.Keys) {
        Complete-ModulePlan $plans[$moduleName] $apiNames
        $row.ApiCallStatements += $plans[$moduleName].ApiCallStatements
    }

    $changedModules = 0
    $ole2Bytes = [byte[]]$project.Ole2Bytes.Clone()

    foreach ($moduleName in $plans.Keys) {
        $plan = $plans[$moduleName]
        $changedGroupCount = 0
        foreach ($flag in $plan.Break) { if ($flag) { $changedGroupCount++ } }
        if ($changedGroupCount -eq 0) { continue }

        $changedModules++
        $newLines = Apply-ModulePlan $plan
        $row.ChangedLines += $plan.ChangedLines

        try {
            $newStream = New-ModuleStream $newLines $encoding $plan.Module $project.Ole2
        } catch {
            if ($_.Exception.Message -notmatch 'exceeds existing OLE2 chain capacity') { throw }
            $newLines = Apply-ModulePlan $plan -Compact
            $newStream = New-ModuleStream $newLines $encoding $plan.Module $project.Ole2
        }
        Write-Ole2Stream $ole2Bytes $project.Ole2 $plan.Module.Entry $newStream
    }

    if ($changedModules -gt 0) {
        Save-VbaProjectBytes $OutputPath $ole2Bytes $project.IsZip
        $verified = Get-AllModuleCode $OutputPath -IncludeRawData
        Assert-SanitizedSource $verified $Rules $apiNames
        $row.Status = 'sanitized'
    } else {
        $row.Status = 'copied-clean'
    }

    return $row
}

$rules = Get-EdrRules
$files = @(Get-SanitizeTargets $Path)
if ($files.Count -eq 0) {
    Write-VbaError 'Sanitize' '-' 'No supported Excel/VBA files found'
    exit 1
}

$timestamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
$devkitRoot = Split-Path "$PSScriptRoot" -Parent
$outputRoot = Join-Path $devkitRoot 'output'
$outDir = Join-Path $outputRoot "${timestamp}_sanitize"
[void][IO.Directory]::CreateDirectory($outDir)

$baseDir = Get-CommonBaseDirectory $files
$usedOutputNames = New-Object 'System.Collections.Hashtable' ([System.StringComparer]::OrdinalIgnoreCase)
$rows = [System.Collections.ArrayList]::new()
$processed = 0

Write-VbaLog 'Sanitize' $baseDir "=== Sanitize session started: $($files.Count) files ==="
Write-VbaLog 'Sanitize' $baseDir "Output dir: $outDir"

foreach ($file in $files) {
    $processed++
    $fileName = [IO.Path]::GetFileName($file)
    $outputPath = New-SanitizedOutputPath $outDir $baseDir $file $usedOutputNames
    Write-VbaHeader 'Sanitize' $fileName
    Write-VbaStatus 'Sanitize' $fileName "Processing $processed of $($files.Count)"

    try {
        $row = Invoke-SanitizeFile $file $outputPath $baseDir $rules
        [void]$rows.Add([pscustomobject]$row)
        Write-VbaResult 'Sanitize' $fileName "$($row.Status): $($row.ChangedLines) lines" $outDir 0
        Write-VbaLog 'Sanitize' $file "$($row.Status): direct=$($row.DirectStatements) apiCalls=$($row.ApiCallStatements) lines=$($row.ChangedLines) -> $outputPath"
    } catch {
        $err = $_.Exception.Message
        $row = [ordered]@{
            Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
            RelativePath = Get-RelativePathText $baseDir $file
            FileName = $fileName
            OutputFile = [IO.Path]::GetFileName($outputPath)
            Status = 'error'
            Modules = 0
            DirectStatements = 0
            ApiCallStatements = 0
            ChangedLines = 0
            Error = $err
        }
        [void]$rows.Add([pscustomobject]$row)
        Write-VbaError 'Sanitize' $fileName $err
        Write-VbaLog 'Sanitize' $file "ERROR: $err" 'ERROR'
    }
}

$summaryPath = Join-Path $outDir 'sanitize.csv'
$rows | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "Sanitize output: $outDir" -ForegroundColor Green
Write-Host "Summary: $summaryPath" -ForegroundColor Gray
