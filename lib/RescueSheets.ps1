param(
    [Parameter(Mandatory = $true)]
    [string[]]$Paths,

    [string]$OutputPath,

    # Default for .xlsm/.xltm is to remove the entire VBA package and emit
    # a macro-free .xlsx/.xltx rescue copy. Use this only when the
    # macro-enabled container must be preserved.
    [switch]$KeepMacroContainer,

    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot\VBAToolkit.psm1" -Force -DisableNameChecking

function Add-RescueFile {
    param([System.Collections.ArrayList]$Files, [string]$FilePath)
    $ext = [IO.Path]::GetExtension($FilePath).ToLower()
    if ($ext -in @('.xlsm', '.xltm', '.xlam', '.xls')) { [void]$Files.Add($FilePath) }
}

function Get-RescueFiles {
    param([string[]]$InputPaths)
    $files = [System.Collections.ArrayList]::new()
    foreach ($p in $InputPaths) {
        $p = $p.Trim().Trim('"')
        $resolved = (Resolve-Path -LiteralPath $p -ErrorAction SilentlyContinue).Path
        if (-not $resolved) {
            Write-VbaError 'RescueSheets' $p 'Path not found'
            continue
        }
        if (Test-Path -LiteralPath $resolved -PathType Container) {
            Get-ChildItem ([WildcardPattern]::Escape($resolved)) -Recurse -File -Include '*.xlsm','*.xltm','*.xlam','*.xls' |
                Where-Object { $_.FullName -notmatch '[\\/](output|debug_output)[\\/]' } |
                ForEach-Object { Add-RescueFile $files $_.FullName }
        } else {
            Add-RescueFile $files $resolved
        }
    }
    return ,$files
}

function New-RescueOutputPath {
    param([string]$FilePath, [string]$OutDir, [string]$Mode)
    $baseName = [IO.Path]::GetFileNameWithoutExtension($FilePath)
    $ext = [IO.Path]::GetExtension($FilePath).ToLower()
    if ($Mode -eq 'StripPackage') {
        if ($ext -eq '.xlsm') { return (Join-Path $OutDir "${baseName}_macrofree.xlsx") }
        if ($ext -eq '.xltm') { return (Join-Path $OutDir "${baseName}_macrofree.xltx") }
    }
    return (Join-Path $OutDir "${baseName}_code_removed${ext}")
}

function Read-ZipEntryText {
    param($Zip, [string]$EntryName)
    $entry = $Zip.GetEntry($EntryName)
    if (-not $entry) { return $null }
    $stream = $entry.Open()
    try {
        $reader = New-Object IO.StreamReader($stream, [Text.Encoding]::UTF8, $true)
        try { return $reader.ReadToEnd() } finally { $reader.Close() }
    } finally { $stream.Close() }
}

function Write-ZipEntryText {
    param($Zip, [string]$EntryName, [string]$Text)
    $old = $Zip.GetEntry($EntryName)
    if ($old) { $old.Delete() }
    $entry = $Zip.CreateEntry($EntryName, [System.IO.Compression.CompressionLevel]::Optimal)
    $stream = $entry.Open()
    try {
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        $writer = New-Object IO.StreamWriter($stream, $utf8NoBom)
        try { $writer.Write($Text) } finally { $writer.Close() }
    } finally { $stream.Close() }
}

function Update-ContentTypesForMacroFreeWorkbook {
    param([xml]$ContentTypesXml)
    $removed = 0
    foreach ($node in @($ContentTypesXml.DocumentElement.ChildNodes)) {
        if ($node.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
        $partName = if ($node.Attributes['PartName']) { $node.Attributes['PartName'].Value } else { '' }
        $contentType = if ($node.Attributes['ContentType']) { $node.Attributes['ContentType'].Value } else { '' }
        if (($partName -match '(?i)vba(Project|Data)') -or ($contentType -match '(?i)vba(Project|ProjectSignature|Data)')) {
            [void]$ContentTypesXml.DocumentElement.RemoveChild($node)
            $removed++
            continue
        }
        if ($node.LocalName -eq 'Override' -and $partName -eq '/xl/workbook.xml') {
            if ($contentType -eq 'application/vnd.ms-excel.sheet.macroEnabled.main+xml') {
                $node.Attributes['ContentType'].Value = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
            } elseif ($contentType -eq 'application/vnd.ms-excel.template.macroEnabled.main+xml') {
                $node.Attributes['ContentType'].Value = 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml'
            }
        }
    }
    return $removed
}

function Update-WorkbookRelsForMacroFreeWorkbook {
    param([xml]$RelsXml)
    $removed = 0
    foreach ($node in @($RelsXml.DocumentElement.ChildNodes)) {
        if ($node.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
        $type = if ($node.Attributes['Type']) { $node.Attributes['Type'].Value } else { '' }
        $target = if ($node.Attributes['Target']) { $node.Attributes['Target'].Value } else { '' }
        if ($type -match '(?i)vba(Project|Data)' -or $target -match '(?i)vba(Project|Data)') {
            [void]$RelsXml.DocumentElement.RemoveChild($node)
            $removed++
        }
    }
    return $removed
}

function Remove-OoxmlVbaPackage {
    param([string]$InputPath, [string]$OutputPath, [switch]$DryRun)
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $report = [ordered]@{ RemovedEntries = @(); RemovedContentTypes = 0; RemovedRelationships = 0 }

    if ($DryRun) {
        $zipRead = [System.IO.Compression.ZipFile]::OpenRead($InputPath)
        try {
            foreach ($entry in $zipRead.Entries) {
                if ($entry.FullName -match '(?i)(^|/)vba(Project|Data)[^/]*\.(bin|xml)$') { $report.RemovedEntries += $entry.FullName }
            }
        } finally { $zipRead.Dispose() }
        return $report
    }

    Copy-Item -LiteralPath $InputPath -Destination $OutputPath -Force
    $zip = [System.IO.Compression.ZipFile]::Open($OutputPath, [System.IO.Compression.ZipArchiveMode]::Update)
    try {
        $entriesToDelete = @()
        foreach ($entry in $zip.Entries) {
            if ($entry.FullName -match '(?i)(^|/)vba(Project|Data)[^/]*\.(bin|xml)$') { $entriesToDelete += $entry }
        }
        foreach ($entry in $entriesToDelete) {
            $report.RemovedEntries += $entry.FullName
            $entry.Delete()
        }

        $ctText = Read-ZipEntryText $zip '[Content_Types].xml'
        if ($ctText) {
            $ctXml = New-Object xml
            $ctXml.PreserveWhitespace = $true
            $ctXml.LoadXml($ctText)
            $report.RemovedContentTypes = Update-ContentTypesForMacroFreeWorkbook $ctXml
            Write-ZipEntryText $zip '[Content_Types].xml' $ctXml.OuterXml
        }

        $relsText = Read-ZipEntryText $zip 'xl/_rels/workbook.xml.rels'
        if ($relsText) {
            $relsXml = New-Object xml
            $relsXml.PreserveWhitespace = $true
            $relsXml.LoadXml($relsText)
            $report.RemovedRelationships = Update-WorkbookRelsForMacroFreeWorkbook $relsXml
            Write-ZipEntryText $zip 'xl/_rels/workbook.xml.rels' $relsXml.OuterXml
        }
    } finally { $zip.Dispose() }
    return $report
}

function Get-MinimalModuleSource {
    param($Module)
    $lines = @()
    foreach ($line in $Module.Lines) {
        if ($line -match '^\s*Attribute\s+VB_') { $lines += $line }
    }
    $hasName = $false
    foreach ($line in $lines) {
        if ($line -match '^\s*Attribute\s+VB_Name\s*=') { $hasName = $true; break }
    }
    if (-not $hasName) {
        $safeName = $Module.Name -replace '"', '""'
        $lines = @("Attribute VB_Name = `"$safeName`"") + $lines
    }
    $lines += ''
    $lines += "' code removed for safe sheet recovery"
    return ,$lines
}

function Get-StreamChainCapacity {
    param($ole2, $entry)
    if ($entry.Size -lt $ole2.MiniStreamCutoff) {
        $s = $entry.Start; $len = 0
        while ($s -ge 0 -and $s -ne -2) {
            $len++
            $s = if ($s -lt $ole2.MiniFat.Length) { $ole2.MiniFat[$s] } else { -1 }
        }
        return $len * $ole2.MiniSectorSize
    }
    $s = $entry.Start; $len = 0; $visited = @{}
    while ($s -ge 0 -and $s -ne -2 -and -not $visited.ContainsKey($s)) {
        $visited[$s] = $true
        $len++
        $s = $ole2.Fat[$s]
    }
    return $len * $ole2.SectorSize
}

function Invalidate-VbaProjectPCode {
    param([byte[]]$Ole2Bytes, $Project)
    $vbaProjEntry = $Project.Ole2.Entries | Where-Object { $_.Name -eq '_VBA_PROJECT' -and $_.ObjType -eq 2 } | Select-Object -First 1
    if ($vbaProjEntry -and $vbaProjEntry.Size -gt 4) {
        $vbaProjData = Read-Ole2Stream $Project.Ole2 $vbaProjEntry
        $vbaProjData[2] = 0x01
        $vbaProjData[3] = 0x00
        Write-Ole2Stream $Ole2Bytes $Project.Ole2 $vbaProjEntry $vbaProjData
    }
}

function Clear-VbaProjectCode {
    param([string]$InputPath, [string]$OutputPath, [switch]$DryRun)
    $project = Get-AllModuleCode $InputPath -IncludeRawData
    if (-not $project) { return @{ Modules = 0; Written = $false; Reason = 'No vbaProject.bin found' } }
    $encoding = [System.Text.Encoding]::GetEncoding($project.Codepage)
    $ole2Bytes = $null
    if (-not $DryRun) {
        Copy-Item -LiteralPath $InputPath -Destination $OutputPath -Force
        $ole2Bytes = [byte[]]$project.Ole2Bytes.Clone()
    }

    $moduleCount = 0
    foreach ($modName in $project.Modules.Keys) {
        $md = $project.Modules[$modName]
        if (-not $md.Entry) { continue }
        $minimalLines = Get-MinimalModuleSource $md
        $newBytes = $encoding.GetBytes(($minimalLines -join "`r`n"))
        $recompressed = Compress-VBA $newBytes
        $newStream = New-Object byte[] ($md.Offset + $recompressed.Length)
        [Array]::Copy($recompressed, 0, $newStream, $md.Offset, $recompressed.Length)
        $capacity = Get-StreamChainCapacity $project.Ole2 $md.Entry
        if ($newStream.Length -gt $capacity) {
            throw "Module $modName cannot be cleared in-place ($($newStream.Length) > $capacity bytes)."
        }
        if (-not $DryRun) { Write-Ole2Stream $ole2Bytes $project.Ole2 $md.Entry $newStream }
        $moduleCount++
    }

    if (-not $DryRun) {
        Invalidate-VbaProjectPCode $ole2Bytes $project
        Save-VbaProjectBytes $OutputPath $ole2Bytes $project.IsZip
    }
    return @{ Modules = $moduleCount; Written = (-not $DryRun); Reason = '' }
}

$files = Get-RescueFiles $Paths
if ($files.Count -eq 0) {
    Write-Host 'No supported Excel macro files found.' -ForegroundColor Yellow
    exit 0
}
if ($OutputPath -and $files.Count -ne 1) {
    Write-VbaError 'RescueSheets' $OutputPath '-OutputPath can be used only with a single input file'
    exit 1
}

$sw = [System.Diagnostics.Stopwatch]::StartNew()
$outDir = $null
if (-not $DryRun) {
    if ($OutputPath) {
        $outDir = [IO.Path]::GetDirectoryName($OutputPath)
        if (-not $outDir) { $outDir = (Get-Location).Path }
        if (-not (Test-Path $outDir)) { [void][IO.Directory]::CreateDirectory($outDir) }
    } else {
        $outDir = New-VbaOutputDir $files[0] 'rescue'
    }
}

$ok = 0; $failed = 0
foreach ($file in $files) {
    $fileName = [IO.Path]::GetFileName($file)
    $ext = [IO.Path]::GetExtension($file).ToLower()
    Write-VbaHeader 'RescueSheets' $fileName
    Write-VbaLog 'RescueSheets' $file 'Started'
    try {
        $useStripPackage = ($ext -in @('.xlsm', '.xltm')) -and (-not $KeepMacroContainer)
        if ($useStripPackage) {
            $dest = if ($OutputPath) { $OutputPath } else { New-RescueOutputPath $file $outDir 'StripPackage' }
            $report = Remove-OoxmlVbaPackage $file $dest -DryRun:$DryRun
            if ($DryRun) {
                Write-VbaStatus 'RescueSheets' $fileName "$($report.RemovedEntries.Count) VBA package part(s) would be removed"
            } else {
                Write-VbaStatus 'RescueSheets' $fileName "$($report.RemovedEntries.Count) VBA package part(s) removed"
                Write-VbaStatus 'RescueSheets' $fileName "$($report.RemovedRelationships) workbook relationship(s) removed"
                Write-VbaStatus 'RescueSheets' $fileName "$($report.RemovedContentTypes) content type entry(ies) removed/updated"
                Write-VbaStatus 'RescueSheets' $fileName "File: $dest"
            }
            Write-VbaLog 'RescueSheets' $file "Macro package stripped -> $dest"
        } else {
            $dest = if ($OutputPath) { $OutputPath } else { New-RescueOutputPath $file $outDir 'ClearCode' }
            if ($ext -eq '.xlam' -and (-not $KeepMacroContainer)) {
                Write-VbaStatus 'RescueSheets' $fileName '.xlam has no macro-free workbook extension; clearing VBA code in-place copy'
            }
            $result = Clear-VbaProjectCode $file $dest -DryRun:$DryRun
            if ($DryRun) {
                Write-VbaStatus 'RescueSheets' $fileName "$($result.Modules) module(s) would be cleared"
            } else {
                Write-VbaStatus 'RescueSheets' $fileName "$($result.Modules) module(s) cleared"
                Write-VbaStatus 'RescueSheets' $fileName "File: $dest"
            }
            Write-VbaLog 'RescueSheets' $file "VBA code cleared: $($result.Modules) modules -> $dest"
        }
        $ok++
    } catch {
        $failed++
        Write-VbaError 'RescueSheets' $fileName $_.Exception.Message
    }
}

$sw.Stop()
Write-Host ''
if ($failed -eq 0) {
    Write-VbaResult 'RescueSheets' "$($files.Count) file(s)" "ok=$ok failed=0" $outDir $sw.Elapsed.TotalSeconds
} else {
    Write-Host "  ok=$ok failed=$failed" -ForegroundColor Yellow
    if ($outDir) { Write-Host "  Output: $outDir" -ForegroundColor Gray }
    Write-Host "  Done ($([Math]::Round($sw.Elapsed.TotalSeconds, 1))s)" -ForegroundColor Gray
    exit 1
}
