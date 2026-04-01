param(
    [string]$Path,
    [string]$OutputPath,
    [switch]$DryRun
)
$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot\VBAToolkit.psm1" -Force -DisableNameChecking

# ============================================================================
# EDR patterns to sanitize (probe-BLOCKED only)
# ============================================================================

$sanitizePatterns = [ordered]@{
    'Win32 API (Declare)' = @{
        LinePattern = '(?i)^[^'']*\bDeclare\s+(PtrSafe\s+)?(Function|Sub)\s+(\w+)\s+Lib\s+"([^"]+)"'
        Rewrite = {
            param($line, $m)
            $scope = ''
            if ($line -match '(?i)^\s*(Private|Public)\s') { $scope = "$($Matches[1]) " }
            $type = $m.Groups[2].Value    # Function or Sub
            $name = $m.Groups[3].Value    # API name
            $dll  = $m.Groups[4].Value    # DLL name
            $alias = ''
            if ($line -match 'Alias\s+"([^"]+)"') { $alias = " (alias: $($Matches[1]))" }
            $sig = ''
            if ($line -match '\(([^)]*)\)\s*(As\s+\w+)?\s*$') {
                $params = $Matches[1].Trim()
                $ret = if ($Matches[2]) { " -> $($Matches[2] -replace 'As\s+', '')" } else { '' }
                $sig = " | params: $params$ret"
            }
            return "' [sanitized:API] ${scope}${type}: ${name} @ ${dll}${alias}${sig}"
        }
    }
    'Shell / process' = @{
        LinePattern = '(?i)^[^'']*\b(Shell\s*[\("]|WScript\.Shell|cmd\s*/[ck])'
        Rewrite = {
            param($line, $m)
            $trigger = $m.Groups[1].Value.Trim()
            $indent = if ($line -match '^(\s+)') { $Matches[1] } else { '' }
            return "${indent}' [sanitized:Process] (original line contained process invocation: $trigger...)"
        }
    }
    'PowerShell / WScript' = @{
        LinePattern = '(?i)^[^'']*\b(powershell|wscript|cscript|mshta)\b'
        Rewrite = {
            param($line, $m)
            $trigger = $m.Groups[1].Value
            $indent = if ($line -match '^(\s+)') { $Matches[1] } else { '' }
            return "${indent}' [sanitized:Script] (original line contained script host invocation: $trigger)"
        }
    }
}

# ============================================================================
# Chain capacity check
# ============================================================================

function Get-StreamChainCapacity {
    param($ole2, $entry)
    if ($entry.Size -lt $ole2.MiniStreamCutoff) {
        $s = $entry.Start; $len = 0
        while ($s -ge 0 -and $s -ne -2) {
            $len++
            $s = if ($s -lt $ole2.MiniFat.Length) { $ole2.MiniFat[$s] } else { -1 }
        }
        return $len * $ole2.MiniSectorSize
    } else {
        $s = $entry.Start; $len = 0; $visited = @{}
        while ($s -ge 0 -and $s -ne -2 -and -not $visited.ContainsKey($s)) {
            $visited[$s] = $true; $len++; $s = $ole2.Fat[$s]
        }
        return $len * $ole2.SectorSize
    }
}

# ============================================================================
# Input validation
# ============================================================================

if (-not $Path) {
    Write-Host 'Usage: Sanitize.ps1 -Path <xlsm file> [-OutputPath <path>] [-DryRun]' -ForegroundColor Yellow
    Write-Host '  Masks EDR-triggering VBA patterns in xlsm files at the binary level.' -ForegroundColor Gray
    Write-Host '  -DryRun  Show what would be masked without modifying any files.' -ForegroundColor Gray
    exit 1
}

$resolved = (Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue).Path
if (-not $resolved) {
    Write-VbaError 'Sanitize' $Path 'File not found'
    exit 1
}
$ext = [IO.Path]::GetExtension($resolved).ToLower()
if ($ext -notin '.xlsm', '.xlam', '.xls') {
    Write-VbaError 'Sanitize' $Path "Unsupported extension: $ext"
    exit 1
}

$fileName = [IO.Path]::GetFileName($resolved)
$sw = [System.Diagnostics.Stopwatch]::StartNew()

# ============================================================================
# Output path
# ============================================================================

if ($DryRun) {
    Write-VbaHeader 'Sanitize' "$fileName (DryRun)"
} else {
    if (-not $OutputPath) {
        $outDir = New-VbaOutputDir $resolved 'sanitize'
        $baseName = [IO.Path]::GetFileNameWithoutExtension($fileName)
        $ext = [IO.Path]::GetExtension($fileName)
        $OutputPath = Join-Path $outDir "${baseName}_sanitized${ext}"
    } else {
        $outDir = [IO.Path]::GetDirectoryName($OutputPath)
        if (-not (Test-Path $outDir)) { [void][IO.Directory]::CreateDirectory($outDir) }
    }
    Copy-Item -LiteralPath $resolved -Destination $OutputPath -Force
    Write-VbaHeader 'Sanitize' $fileName
}

# ============================================================================
# Load project
# ============================================================================

$targetPath = if ($DryRun) { $resolved } else { $OutputPath }
$project = Get-AllModuleCode $targetPath -IncludeRawData
if (-not $project) {
    Write-VbaError 'Sanitize' $fileName 'No vbaProject.bin found'
    exit 1
}
$encoding = [System.Text.Encoding]::GetEncoding($project.Codepage)
$ole2Bytes = $null
if (-not $DryRun) {
    $ole2Bytes = [byte[]]$project.Ole2Bytes.Clone()
}

# ============================================================================
# Scan and mask
# ============================================================================

$totalMasked = 0
$report = [System.Collections.ArrayList]::new()

foreach ($modName in $project.Modules.Keys) {
    $md = $project.Modules[$modName]
    if (-not $md.Entry) { continue }

    $lines = @($md.Lines)
    $modified = $false
    $modMasked = 0

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        # Skip lines that are already comments
        if ($line -match '^\s*''') { continue }

        foreach ($catName in $sanitizePatterns.Keys) {
            $pat = $sanitizePatterns[$catName]
            if ($line -match $pat.LinePattern) {
                $original = $line
                $regexMatch = [regex]::Match($line, $pat.LinePattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)

                # Handle line continuation — join into single line first
                $fullLine = $original
                $contCount = 0
                while ($fullLine -match '_\s*$' -and ($i + $contCount + 1) -lt $lines.Count) {
                    $contCount++
                    $fullLine = $fullLine -replace '_\s*$', ''
                    $fullLine += ' ' + $lines[$i + $contCount].TrimStart()
                }
                if ($contCount -gt 0) {
                    $regexMatch = [regex]::Match($fullLine, $pat.LinePattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
                }

                $lines[$i] = & $pat.Rewrite $fullLine $regexMatch

                # Blank out continuation lines
                for ($c = 1; $c -le $contCount; $c++) {
                    $lines[$i + $c] = "' [sanitized-cont]"
                    $modMasked++
                }
                $i += $contCount

                $modified = $true
                $modMasked++
                [void]$report.Add(@{ Module = "$modName.$($md.Ext)"; Line = $i + 1; Category = $catName; Original = $original; Masked = $lines[$i] })
                break  # one match per line
            }
        }
    }

    if ($DryRun) {
        if ($modMasked -gt 0) {
            Write-VbaStatus 'Sanitize' $fileName "$modName.$($md.Ext): $modMasked line(s) to sanitize"
        }
        $totalMasked += $modMasked
        continue
    }

    if (-not $modified) { continue }
    $totalMasked += $modMasked

    # Re-encode to bytes
    $newText = $lines -join "`r`n"
    $newBytes = $encoding.GetBytes($newText)

    # Recompress
    $recompressed = Compress-VBA $newBytes

    # Rebuild stream: [p-code up to Offset] + [new compressed data]
    $newStream = New-Object byte[] ($md.Offset + $recompressed.Length)
    [Array]::Copy($md.StreamData, 0, $newStream, 0, $md.Offset)
    [Array]::Copy($recompressed, 0, $newStream, $md.Offset, $recompressed.Length)

    # Check capacity
    $capacity = Get-StreamChainCapacity $project.Ole2 $md.Entry
    if ($newStream.Length -gt $capacity) {
        Write-Host "  WARN: $modName stream exceeds capacity ($($newStream.Length) > $capacity). Using minimal mask." -ForegroundColor Yellow
        # Fallback: minimal comment only
        $lines2 = @($md.Lines)
        for ($j = 0; $j -lt $lines2.Count; $j++) {
            if ($lines2[$j] -match '^\s*''') { continue }
            foreach ($catName in $sanitizePatterns.Keys) {
                if ($lines2[$j] -match $sanitizePatterns[$catName].LinePattern) {
                    $lines2[$j] = "' [sanitized:$catName]"
                    if ($lines2[$j] -match '_\s*$') {
                        while (($j + 1) -lt $lines2.Count -and $lines2[$j] -match '_\s*$') {
                            $j++; $lines2[$j] = "' [sanitized-cont]"
                        }
                    }
                    break
                }
            }
        }
        $newBytes2 = $encoding.GetBytes($lines2 -join "`r`n")
        $recompressed2 = Compress-VBA $newBytes2
        $newStream = New-Object byte[] ($md.Offset + $recompressed2.Length)
        [Array]::Copy($md.StreamData, 0, $newStream, 0, $md.Offset)
        [Array]::Copy($recompressed2, 0, $newStream, $md.Offset, $recompressed2.Length)
    }

    # Write back to OLE2
    Write-Ole2Stream $ole2Bytes $project.Ole2 $md.Entry $newStream
    Write-VbaStatus 'Sanitize' $fileName "$modName.$($md.Ext): $modMasked line(s) sanitized"
}

# ============================================================================
# Save and report
# ============================================================================

if ($DryRun) {
    Write-Host ''
    if ($totalMasked -eq 0) {
        Write-Host '  No EDR patterns found.' -ForegroundColor Green
    } else {
        Write-Host "  $totalMasked line(s) would be sanitized" -ForegroundColor Yellow
        Write-Host ''
        foreach ($r in $report) {
            Write-Host "  $($r.Module):$($r.Line)  [$($r.Category)]" -ForegroundColor Gray
            Write-Host "    $($r.Original)" -ForegroundColor DarkGray
        }
    }
} else {
    if ($totalMasked -gt 0) {
        # Invalidate p-code so Excel re-decompiles from source text
        # _VBA_PROJECT stream offset 2-3 = version; changing it forces recompile
        $vbaProjEntry = $project.Ole2.Entries | Where-Object { $_.Name -eq '_VBA_PROJECT' -and $_.ObjType -eq 2 } | Select-Object -First 1
        if ($vbaProjEntry -and $vbaProjEntry.Size -gt 4) {
            $vbaProjData = Read-Ole2Stream $project.Ole2 $vbaProjEntry
            # Flip the version to an invalid value
            $vbaProjData[2] = 0x01
            $vbaProjData[3] = 0x00
            Write-Ole2Stream $ole2Bytes $project.Ole2 $vbaProjEntry $vbaProjData
        }

        Save-VbaProjectBytes $OutputPath $ole2Bytes $project.IsZip
    }

    $sw.Stop()
    if ($totalMasked -eq 0) {
        Write-VbaResult 'Sanitize' $fileName 'No EDR patterns found' $null $sw.Elapsed.TotalSeconds
    } else {
        Write-VbaResult 'Sanitize' $fileName "$totalMasked line(s) sanitized" $outDir $sw.Elapsed.TotalSeconds
        Write-Host "  File: $OutputPath" -ForegroundColor Gray
    }
    Write-VbaLog 'Sanitize' $resolved "$totalMasked lines sanitized -> $OutputPath"
}
