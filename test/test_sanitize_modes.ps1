param(
    [string]$Fixture = (Join-Path $PSScriptRoot 'test_sample.xlsm')
)

$ErrorActionPreference = 'Stop'

$root = Split-Path $PSScriptRoot -Parent
$sanitizer = Join-Path $root 'lib\Sanitize.ps1'
$toolkit = Join-Path $root 'lib\VBAToolkit.psm1'
$outputRoot = Join-Path $root 'output'

if (-not (Test-Path -LiteralPath $Fixture)) {
    throw "Fixture not found: $Fixture"
}

Import-Module $toolkit -Force -DisableNameChecking

$base = [IO.Path]::GetFileNameWithoutExtension($Fixture)
$ext = [IO.Path]::GetExtension($Fixture)
$forbidden = @(
    '\bDeclare\b',
    '\bGetTickCount\b',
    '\bSleep\b',
    '\bGetUserName\b',
    '\bGetUserNameA\b',
    '\bkernel32\b',
    '\badvapi32\b',
    '\bShell\b',
    '\bpowershell\b',
    '\bwscript\b',
    '\bcscript\b',
    '\bmshta\b',
    '\bcmd\s*/[ck]\b'
)

foreach ($mode in 1..10) {
    $started = Get-Date
    & $sanitizer -Mode $mode -Path $Fixture | Out-Null

    $outDir = Get-ChildItem -LiteralPath $outputRoot -Directory -Filter '*_sanitize' |
        Where-Object { $_.LastWriteTime -ge $started.AddSeconds(-2) } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if (-not $outDir) {
        throw "Mode $mode did not create an output directory."
    }

    $summary = Import-Csv (Join-Path $outDir.FullName 'sanitize.csv') | Select-Object -First 1
    if (-not $summary) {
        throw "Mode $mode did not write sanitize.csv."
    }
    if ([int]$summary.Mode -ne $mode) {
        throw "Mode $mode wrote wrong summary mode: $($summary.Mode)"
    }
    $expectedStatus = if (@(3, 4, 5) -contains $mode) { 'sanitized-experimental' } else { 'sanitized' }
    if ($summary.Status -ne $expectedStatus) {
        throw "Mode $mode status mismatch: $($summary.Status), expected $expectedStatus"
    }

    $sanitized = Join-Path $outDir.FullName "${base}_sanitized$ext"
    if (-not (Test-Path -LiteralPath $sanitized)) {
        throw "Mode $mode did not create sanitized workbook: $sanitized"
    }

    $project = Get-AllModuleCode $sanitized -IncludeRawData
    if (-not $project) {
        throw "Mode $mode sanitized workbook has no readable VBA project."
    }

    $allText = ($project.Modules.Keys | ForEach-Object { $project.Modules[$_].Code }) -join "`n"
    if ($allText -notmatch '\*\*\*') {
        throw "Mode $mode did not include replacement markers."
    }

    switch ($mode) {
        1 {
            if ($allText -notmatch 'role=') { throw 'Mode 1 metadata marker missing.' }
        }
        2 {
            if ($allText -notmatch 'original-vba' -or $allText -notmatch '\bDeclare\b') { throw 'Mode 2 original text missing.' }
        }
        3 {
            if ($allText -notmatch 'original-rem' -or $allText -notmatch '\bDeclare\b') { throw 'Mode 3 original text missing.' }
        }
        4 {
            if ($allText -notmatch 'original-slash' -or $allText -notmatch '\bDeclare\b') { throw 'Mode 4 original text missing.' }
        }
        5 {
            if ($allText -notmatch 'original-block' -or $allText -notmatch '\bDeclare\b') { throw 'Mode 5 original text missing.' }
        }
        default {
            if ($allText -notmatch "masked$mode") { throw "Mode $mode mask marker missing." }
        }
    }

    if (@(1, 6, 7, 8, 9, 10) -contains $mode) {
        foreach ($pattern in $forbidden) {
            if ([regex]::IsMatch($allText, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
                throw "Mode $mode left forbidden token in sanitized source: $pattern"
            }
        }
    }

    Write-Host "mode $mode OK"
}

Write-Host 'sanitize modes test OK'
