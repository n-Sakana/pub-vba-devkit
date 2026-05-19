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

$started = Get-Date
& $sanitizer -Path $Fixture

$outDir = Get-ChildItem -LiteralPath $outputRoot -Directory -Filter '*_sanitize' |
    Where-Object { $_.LastWriteTime -ge $started.AddSeconds(-2) } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

if (-not $outDir) {
    throw 'Sanitize output directory was not created.'
}

$base = [IO.Path]::GetFileNameWithoutExtension($Fixture)
$ext = [IO.Path]::GetExtension($Fixture)
$sanitized = Join-Path $outDir.FullName "${base}_sanitized$ext"
if (-not (Test-Path -LiteralPath $sanitized)) {
    throw "Sanitized workbook was not created: $sanitized"
}

$summary = Import-Csv (Join-Path $outDir.FullName 'sanitize.csv') | Select-Object -First 1
if (-not $summary -or $summary.Status -ne 'sanitized') {
    throw "Unexpected sanitize status: $($summary.Status)"
}
if ([int]$summary.ChangedLines -le 0) {
    throw 'Sanitize did not replace any lines in the fixture.'
}

Import-Module $toolkit -Force -DisableNameChecking
$project = Get-AllModuleCode $sanitized -IncludeRawData
if (-not $project) {
    throw 'Sanitized workbook has no readable VBA project.'
}

$defs = Get-VbaAnalysis -Project @{ Modules = [ordered]@{}; Ole2 = $null }
$edrRuleNames = @('Win32 API (Declare)', 'Shell / process', 'PowerShell / WScript')
$allText = ($project.Modules.Keys | ForEach-Object { $project.Modules[$_].Code }) -join "`n"

foreach ($ruleName in $edrRuleNames) {
    $pattern = $defs.Patterns[$ruleName].Pattern
    if ([regex]::IsMatch($allText, $pattern)) {
        throw "EDR rule remains after sanitize: $ruleName"
    }
}

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

foreach ($pattern in $forbidden) {
    if ([regex]::IsMatch($allText, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
        throw "Forbidden token remains in sanitized source: $pattern"
    }
}

if ($allText -notmatch '\*\*\*') {
    throw 'Replacement marker was not found in sanitized source.'
}

Write-Host 'sanitize static test OK'
