$ErrorActionPreference = 'Stop'

function Invoke-DevkitScript {
    param(
        [string]$ScriptPath,
        [string[]]$Arguments = @()
    )

    # Quote all values that may contain spaces (OneDrive sync paths, output dirs)
    $quotedArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', "`"$ScriptPath`"")
    foreach ($arg in $Arguments) {
        if ($arg.StartsWith('-')) {
            $quotedArgs += $arg
        } else {
            $quotedArgs += "`"$arg`""
        }
    }

    $process = Start-Process -FilePath 'powershell.exe' `
        -ArgumentList $quotedArgs `
        -NoNewWindow `
        -Wait `
        -PassThru

    if ($process -and $process.HasExited) {
        return $process.ExitCode
    }
    return 0
}

function Read-ProbeScenario {
    Write-Host ''
    Write-Host 'Scenario'
    Write-Host '  1 = Local (desktop, local drive)'
    Write-Host '  2 = SharePoint folder (OneDrive sync / local synced copy)'
    Write-Host '  used by: Probe / Full modes' -ForegroundColor Gray
    $scenarioNum = (Read-Host 'Scenario (1/2)').Trim()
    switch ($scenarioNum) {
        '2' { return 'OneDrive_Synced' }
        default { return 'Local' }
    }
}

function New-EnvTestReportDirectory {
    param([string]$BaseRoot)

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $path = Join-Path $BaseRoot "${timestamp}_envtest"
    New-Item -Path $path -ItemType Directory -Force | Out-Null
    return $path
}

$scriptRoot = Split-Path -Parent $PSScriptRoot
$internalRoot = Join-Path $PSScriptRoot 'internal'
$surveyPath = Join-Path $internalRoot 'Survey.ps1'
$probePath = Join-Path $internalRoot 'Probe.ps1'
$outputRoot = Join-Path $scriptRoot 'output'

Write-Host '=== Environment Test ===' -ForegroundColor Cyan
$scenario = Read-ProbeScenario

Write-Host ''
Write-Host 'Mode'
Write-Host '  S = Survey only'
Write-Host '  B = Probe (Basic)'
Write-Host '  E = Probe (Basic + Extended)'
Write-Host '  G = Generate SharePoint probe xlsm'
Write-Host '  F = Full (Survey + Probe Basic)'
Write-Host '  X = Full (Survey + Probe Basic + Extended)'
Write-Host '  Q = Quit'
$mode = (Read-Host 'Run mode? (S/B/E/G/F/X/Q)').Trim()
if ($mode -match '^(?i)q$') {
    exit 0
}

$mode = $mode.ToUpperInvariant()
$exitCode = 0
$reportDir = New-EnvTestReportDirectory -BaseRoot $outputRoot

Write-Host ''
Write-Host "Report folder: $reportDir" -ForegroundColor Gray

# Temp files for capturing script output text (placed in reportDir to avoid TEMP path issues)
$utf8Bom = New-Object System.Text.UTF8Encoding $true
$surveyTmp = Join-Path $reportDir '.survey.tmp'
$probeTmp  = Join-Path $reportDir '.probe.tmp'

$parts = [System.Collections.Generic.List[string]]::new()

try {
    switch ($mode) {
        'S' {
            $exitCode = Invoke-DevkitScript -ScriptPath $surveyPath -Arguments @('-ReturnTextFile', $surveyTmp, '-OutputDirectory', $reportDir)
        }
        'B' {
            $exitCode = Invoke-DevkitScript -ScriptPath $probePath -Arguments @('-ReturnTextFile', $probeTmp, '-Mode', 'Basic', '-Scenario', $scenario, '-OutputDirectory', $reportDir)
        }
        'E' {
            $exitCode = Invoke-DevkitScript -ScriptPath $probePath -Arguments @('-ReturnTextFile', $probeTmp, '-Mode', 'Extended', '-Scenario', $scenario, '-OutputDirectory', $reportDir)
        }
        'G' {
            $exitCode = Invoke-DevkitScript -ScriptPath $probePath -Arguments @('-Mode', 'GenerateStorage', '-OutputDirectory', $reportDir)
        }
        'F' {
            $surveyExit = Invoke-DevkitScript -ScriptPath $surveyPath -Arguments @('-ReturnTextFile', $surveyTmp, '-OutputDirectory', $reportDir)
            if ($surveyExit -ne 0) { $exitCode = $surveyExit }
            else {
                $exitCode = Invoke-DevkitScript -ScriptPath $probePath -Arguments @('-ReturnTextFile', $probeTmp, '-Mode', 'Basic', '-Scenario', $scenario, '-OutputDirectory', $reportDir)
            }
        }
        'X' {
            $surveyExit = Invoke-DevkitScript -ScriptPath $surveyPath -Arguments @('-ReturnTextFile', $surveyTmp, '-OutputDirectory', $reportDir)
            if ($surveyExit -ne 0) { $exitCode = $surveyExit }
            else {
                $exitCode = Invoke-DevkitScript -ScriptPath $probePath -Arguments @('-ReturnTextFile', $probeTmp, '-Mode', 'Extended', '-Scenario', $scenario, '-OutputDirectory', $reportDir)
            }
        }
        default {
            Write-Host "Unknown mode: $mode" -ForegroundColor Red
            $exitCode = 1
        }
    }
} catch {
    Write-Host "ERROR in script execution: $_" -ForegroundColor Red
    $exitCode = 1
}

# Collect results from temp files
try {
    foreach ($tmp in @($surveyTmp, $probeTmp)) {
        if (Test-Path $tmp) {
            $parts.Add([IO.File]::ReadAllText($tmp, $utf8Bom))
            Remove-Item $tmp -Force -ErrorAction SilentlyContinue
        }
    }
} catch {
    Write-Host "ERROR collecting results: $_" -ForegroundColor Red
}

# Write single report file
if ($parts.Count -gt 0) {
    $reportPath = Join-Path $reportDir 'envtest.txt'
    [IO.File]::WriteAllText($reportPath, ($parts -join "`r`n`r`n"), $utf8Bom)

    Write-Host ''
    Write-Host "Report: $reportPath" -ForegroundColor Green
} else {
    Write-Host ''
    Write-Host "WARNING: No output was captured from sub-scripts." -ForegroundColor Yellow
    Write-Host "  Survey tmp: $surveyTmp (exists=$(Test-Path $surveyTmp))" -ForegroundColor Yellow
    Write-Host "  Probe tmp:  $probeTmp (exists=$(Test-Path $probeTmp))" -ForegroundColor Yellow
}

Write-Host "Done. Exit code: $exitCode" -ForegroundColor $(if ($exitCode -eq 0) { 'Green' } else { 'Red' })
exit $exitCode
