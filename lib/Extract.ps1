param([Parameter(Mandatory)][string[]]$Paths)
$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot\VBAToolkit.psm1" -Force -DisableNameChecking

$sw = [System.Diagnostics.Stopwatch]::StartNew()

# Collect all xlsm/xlam/xls files
$files = [System.Collections.ArrayList]::new()
$baseDir = $null
foreach ($p in $Paths) {
    $p = $p.Trim().Trim('"')
    $resolved = (Resolve-Path -LiteralPath $p -ErrorAction SilentlyContinue).Path
    if (-not $resolved) { Write-VbaError 'Extract' $p 'Path not found'; continue }

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

Write-VbaLog 'Extract' $baseDir "=== Extract session started: $($files.Count) files from $baseDir ==="

# Single output directory for the entire run
$outDir = New-VbaOutputDir ($files[0]) 'extract'
Write-VbaLog 'Extract' $baseDir "Output dir: $outDir"

$totalExtracted = 0
foreach ($FilePath in $files) {
    $fileName = [IO.Path]::GetFileName($FilePath)
    $baseName = [IO.Path]::GetFileNameWithoutExtension($FilePath).Trim()
    # Prefix with subfolder name when file is not directly under baseDir
    $outPrefix = $baseName
    $fileDir = [IO.Path]::GetDirectoryName($FilePath)
    if ($fileDir -ne $baseDir) {
        $relDir = ''
        if ($FilePath.StartsWith($baseDir)) {
            $relDir = [IO.Path]::GetDirectoryName($FilePath.Substring($baseDir.Length).TrimStart('\', '/'))
        }
        if ($relDir) {
            $outPrefix = ($relDir -replace '[\\/]', '_') + "_$baseName"
        } else {
            $parentDir = Split-Path (Split-Path "$FilePath" -Parent) -Leaf
            $outPrefix = "${parentDir}_${baseName}"
        }
    }
    $fileSw = [System.Diagnostics.Stopwatch]::StartNew()

    Write-VbaHeader 'Extract' $fileName
    Write-VbaLog 'Extract' $FilePath 'Started'

    Write-VbaLog 'Extract' $FilePath 'Loading project (StripAttributes)'
    $project = Get-AllModuleCode $FilePath -StripAttributes
    if (-not $project) {
        Write-VbaError 'Extract' $fileName 'No vbaProject.bin found'
        Write-VbaLog 'Extract' $FilePath 'SKIP: No vbaProject.bin' 'WARN'
        continue
    }
    Write-VbaLog 'Extract' $FilePath "Project loaded: $($project.Modules.Count) modules ($($project.Modules.Keys -join ', '))"

    # Per-file subfolder to avoid module name collisions across files
    $modulesDir = Join-Path $outDir "modules/$outPrefix"
    Write-VbaLog 'Extract' $FilePath "Creating module dir: $modulesDir"
    [void][IO.Directory]::CreateDirectory($modulesDir)

    # Write individual module files
    $extracted = 0
    foreach ($modName in $project.Modules.Keys) {
        $mod = $project.Modules[$modName]
        $outPath = Join-Path $modulesDir "$modName.$($mod.Ext)"
        $lineCount = if ($mod.Lines -is [array]) { $mod.Lines.Count } else { 1 }
        Write-VbaLog 'Extract' $FilePath "  Writing $modName.$($mod.Ext) ($lineCount lines)"
        [IO.File]::WriteAllText($outPath, ($mod.Lines -join "`r`n"), [System.Text.Encoding]::UTF8)
        Write-VbaStatus 'Extract' $fileName "  $modName.$($mod.Ext)"
        $extracted++
    }
    Write-VbaStatus 'Extract' $fileName "$extracted module(s) extracted"
    Write-VbaLog 'Extract' $FilePath "$extracted modules written to $modulesDir"

    # Combined source with module index
    $allFiles = Get-ChildItem "$modulesDir" -File
    $totalLines = 0

    $combined = [System.Text.StringBuilder]::new()
    [void]$combined.AppendLine("=" * 80)
    [void]$combined.AppendLine(" $fileName - VBA Source Code")
    [void]$combined.AppendLine(" Extracted: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')")
    [void]$combined.AppendLine("=" * 80)
    [void]$combined.AppendLine("")

    [void]$combined.AppendLine("MODULE INDEX")
    [void]$combined.AppendLine("-" * 40)
    [void]$combined.AppendLine("")

    $stdMods = @(); $clsMods = @(); $frmMods = @(); $docMods = @()
    foreach ($f in $allFiles) {
        $fExt = $f.Extension.TrimStart('.')
        $lc = (Get-Content "$($f.FullName)" -Encoding UTF8).Count
        $totalLines += $lc
        $entry = "    $($f.Name) ($lc lines)"
        switch ($fExt) {
            'bas' { $stdMods += $entry }
            'cls' { $clsMods += $entry }
            'frm' { $frmMods += $entry }
            default { $docMods += $entry }
        }
    }
    if ($stdMods.Count -gt 0) { [void]$combined.AppendLine("  Standard Modules:"); foreach ($e in $stdMods) { [void]$combined.AppendLine($e) }; [void]$combined.AppendLine("") }
    if ($clsMods.Count -gt 0) { [void]$combined.AppendLine("  Class Modules:"); foreach ($e in $clsMods) { [void]$combined.AppendLine($e) }; [void]$combined.AppendLine("") }
    if ($frmMods.Count -gt 0) { [void]$combined.AppendLine("  UserForms:"); foreach ($e in $frmMods) { [void]$combined.AppendLine($e) }; [void]$combined.AppendLine("") }
    if ($docMods.Count -gt 0) { [void]$combined.AppendLine("  Document Modules:"); foreach ($e in $docMods) { [void]$combined.AppendLine($e) }; [void]$combined.AppendLine("") }
    [void]$combined.AppendLine("  Total: $totalLines lines across $($allFiles.Count) modules")
    [void]$combined.AppendLine("")

    $order = @('bas','cls','frm')
    $sorted = $allFiles | Sort-Object { $o = [Array]::IndexOf($order, $_.Extension.TrimStart('.')); if($o -lt 0){99}else{$o} }, Name
    foreach ($f in $sorted) {
        $c = [IO.File]::ReadAllText($f.FullName, [System.Text.Encoding]::UTF8)
        [void]$combined.AppendLine("=" * 80)
        [void]$combined.AppendLine(" $($f.Name)")
        [void]$combined.AppendLine("=" * 80)
        [void]$combined.AppendLine("")
        [void]$combined.AppendLine($c.TrimStart("`r`n"))
        [void]$combined.AppendLine("")
    }
    $combinedPath = Join-Path $outDir "${outPrefix}_combined.txt"
    Write-VbaLog 'Extract' $FilePath "Writing combined source: $combinedPath ($totalLines total lines)"
    [IO.File]::WriteAllText($combinedPath, $combined.ToString(), [System.Text.Encoding]::UTF8)

    $fileSw.Stop()
    $totalExtracted += $extracted
    Write-VbaResult 'Extract' $fileName "$extracted module(s), $totalLines lines" $outDir $fileSw.Elapsed.TotalSeconds
    Write-VbaLog 'Extract' $FilePath "$extracted modules, $totalLines lines ($([Math]::Round($fileSw.Elapsed.TotalSeconds, 1))s) | -> $outDir"
}

$sw.Stop()
Write-VbaLog 'Extract' $baseDir "=== Extract session complete: $($files.Count) files, $totalExtracted modules ($([Math]::Round($sw.Elapsed.TotalSeconds, 1))s) ==="
if ($files.Count -gt 1) {
    Write-Host "`n  Total: $($files.Count) files, $totalExtracted modules" -ForegroundColor Green
}
