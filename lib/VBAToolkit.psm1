# ============================================================================
# VBAToolkit - Common module for binary-level VBA file manipulation
# OLE2 parser, VBA compression/decompression (MS-OVBA 2.4.1), ZIP helpers
# ============================================================================

$ErrorActionPreference = 'Stop'

# ============================================================================
# C# Native Implementation (high-performance byte operations)
# ============================================================================

if (-not ([System.Management.Automation.PSTypeName]'VbaToolkitNative').Type) {
    Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Collections.Generic;

public static class VbaToolkitNative
{
    public static byte[] ReadSectorChain(byte[] bytes, int startSector, int sectorSize, int[] fat)
    {
        var ms = new MemoryStream();
        var visited = new HashSet<int>();
        int s = startSector;
        while (s >= 0 && s != -2 && s != -1 && !visited.Contains(s))
        {
            visited.Add(s);
            int off = (s + 1) * sectorSize;
            if (off + sectorSize > bytes.Length) break;
            ms.Write(bytes, off, sectorSize);
            s = (s < fat.Length) ? fat[s] : -1;
        }
        return ms.ToArray();
    }

    public static byte[] ReadMiniStream(byte[] miniStreamData, int startSector, int size, int miniSectorSize, int[] miniFat)
    {
        var data = new byte[size];
        int s = startSector;
        int written = 0;
        while (s >= 0 && s != -2 && written < size)
        {
            int off = s * miniSectorSize;
            int toRead = Math.Min(miniSectorSize, size - written);
            if (off + toRead <= miniStreamData.Length)
                Array.Copy(miniStreamData, off, data, written, toRead);
            written += miniSectorSize;
            s = (s < miniFat.Length) ? miniFat[s] : -1;
        }
        return data;
    }

    public static byte[] DecompressVba(byte[] data, int offset)
    {
        if (offset >= data.Length || data[offset] != 1) return new byte[0];
        var result = new List<byte>(data.Length * 2);
        int pos = offset + 1;
        while (pos < data.Length - 1)
        {
            if (pos + 1 >= data.Length) break;
            ushort header = BitConverter.ToUInt16(data, pos); pos += 2;
            int chunkSize = (header & 0x0FFF) + 3;
            bool isCompressed = (header & 0x8000) != 0;
            if (!isCompressed)
            {
                int toCopy = Math.Min(chunkSize - 2, data.Length - pos);
                for (int c = 0; c < toCopy; c++) result.Add(data[pos + c]);
                pos += toCopy;
                continue;
            }
            int chunkEnd = pos + chunkSize - 2;
            if (chunkEnd > data.Length) chunkEnd = data.Length;
            int decompStart = result.Count;
            while (pos < chunkEnd)
            {
                if (pos >= data.Length) break;
                byte flagByte = data[pos]; pos++;
                for (int bit = 0; bit < 8 && pos < chunkEnd; bit++)
                {
                    if ((flagByte & (1 << bit)) == 0)
                    {
                        result.Add(data[pos]); pos++;
                    }
                    else
                    {
                        if (pos + 1 >= data.Length) { pos = chunkEnd; break; }
                        ushort token = BitConverter.ToUInt16(data, pos); pos += 2;
                        int dPos = result.Count - decompStart;
                        if (dPos < 1) dPos = 1;
                        int bitCount = 4;
                        while ((1 << bitCount) < dPos) bitCount++;
                        if (bitCount > 12) bitCount = 12;
                        int lengthMask = 0xFFFF >> bitCount;
                        int copyLen = (token & lengthMask) + 3;
                        int copyOff = (token >> (16 - bitCount)) + 1;
                        for (int c = 0; c < copyLen; c++)
                        {
                            int srcIdx = result.Count - copyOff;
                            if (srcIdx >= 0 && srcIdx < result.Count)
                                result.Add(result[srcIdx]);
                        }
                    }
                }
            }
            pos = chunkEnd;
        }
        return result.ToArray();
    }

    public static byte[] CompressVba(byte[] data)
    {
        var result = new MemoryStream();
        result.WriteByte(1);
        int srcPos = 0;
        while (srcPos < data.Length)
        {
            int chunkStart = srcPos;
            int chunkEnd = Math.Min(srcPos + 4096, data.Length);
            var chunkBuf = new MemoryStream();
            int dPos = srcPos;
            while (dPos < chunkEnd)
            {
                long flagPos = chunkBuf.Position;
                chunkBuf.WriteByte(0);
                byte flagByte = 0;
                for (int bit = 0; bit < 8 && dPos < chunkEnd; bit++)
                {
                    int bestLen = 0, bestOff = 0;
                    int decompPos = dPos - chunkStart;
                    if (decompPos < 1) decompPos = 1;
                    int bitCount = 4;
                    while ((1 << bitCount) < decompPos) bitCount++;
                    if (bitCount > 12) bitCount = 12;
                    int maxOff = 1 << bitCount;
                    int maxLen = (0xFFFF >> bitCount) + 3;
                    for (int off = 1; off <= Math.Min(maxOff, dPos - chunkStart); off++)
                    {
                        int matchLen = 0;
                        while (matchLen < maxLen && dPos + matchLen < chunkEnd)
                        {
                            if (data[dPos - off + (matchLen % off)] != data[dPos + matchLen]) break;
                            matchLen++;
                        }
                        if (matchLen >= 3 && matchLen > bestLen) { bestLen = matchLen; bestOff = off; }
                    }
                    if (bestLen >= 3)
                    {
                        flagByte |= (byte)(1 << bit);
                        int token = ((bestOff - 1) << (16 - bitCount)) | (bestLen - 3);
                        chunkBuf.WriteByte((byte)(token & 0xFF));
                        chunkBuf.WriteByte((byte)((token >> 8) & 0xFF));
                        dPos += bestLen;
                    }
                    else
                    {
                        chunkBuf.WriteByte(data[dPos]); dPos++;
                    }
                }
                long savedPos = chunkBuf.Position;
                chunkBuf.Position = flagPos;
                chunkBuf.WriteByte(flagByte);
                chunkBuf.Position = savedPos;
            }
            byte[] compressed = chunkBuf.ToArray();
            srcPos = dPos;
            ushort hdr = (ushort)(0x8000 | 0x3000 | (compressed.Length + 2 - 3));
            result.WriteByte((byte)(hdr & 0xFF));
            result.WriteByte((byte)((hdr >> 8) & 0xFF));
            result.Write(compressed, 0, compressed.Length);
        }
        return result.ToArray();
    }

    public static int FindPattern(byte[] data, byte[] pattern)
    {
        for (int i = 0; i <= data.Length - pattern.Length; i++)
        {
            bool match = true;
            for (int j = 0; j < pattern.Length; j++)
            {
                if (data[i + j] != pattern[j]) { match = false; break; }
            }
            if (match) return i;
        }
        return -1;
    }
}
'@
}

# ============================================================================
# OLE2 Compound Document Parser
# ============================================================================

function Read-SectorChain([byte[]]$bytes, [int]$startSector, [int]$sectorSize, [int[]]$fat) {
    return ,[VbaToolkitNative]::ReadSectorChain($bytes, $startSector, $sectorSize, $fat)
}

function Read-MiniStream([byte[]]$miniStreamData, [int]$startSector, [int]$size, [int]$miniSectorSize, [int[]]$miniFat) {
    return ,[VbaToolkitNative]::ReadMiniStream($miniStreamData, $startSector, $size, $miniSectorSize, $miniFat)
}

function Read-Ole2([byte[]]$bytes) {
    $sectorPow = [BitConverter]::ToUInt16($bytes, 30)
    $sectorSize = [int][Math]::Pow(2, $sectorPow)
    $miniSectorPow = [BitConverter]::ToUInt16($bytes, 32)
    $miniSectorSize = [int][Math]::Pow(2, $miniSectorPow)
    $firstDirSector = [BitConverter]::ToInt32($bytes, 48)
    $miniStreamCutoff = [BitConverter]::ToUInt32($bytes, 56)
    $firstMiniFatSector = [BitConverter]::ToInt32($bytes, 60)
    $firstDifatSector = [BitConverter]::ToInt32($bytes, 68)

    $difat = [System.Collections.ArrayList]::new()
    for ($i = 0; $i -lt 109; $i++) {
        $v = [BitConverter]::ToInt32($bytes, 76 + $i * 4)
        if ($v -ge 0) { [void]$difat.Add($v) }
    }
    $nextDifat = $firstDifatSector
    while ($nextDifat -ge 0 -and $nextDifat -ne -2) {
        $off = ($nextDifat + 1) * $sectorSize
        $ePS = $sectorSize / 4 - 1
        for ($i = 0; $i -lt $ePS; $i++) {
            $v = [BitConverter]::ToInt32($bytes, $off + $i * 4)
            if ($v -ge 0) { [void]$difat.Add($v) }
        }
        $nextDifat = [BitConverter]::ToInt32($bytes, $off + $ePS * 4)
    }

    $fatEntries = [int]($bytes.Length / $sectorSize)
    [int[]]$fat = New-Object int[] $fatEntries
    for ($i = 0; $i -lt $fatEntries; $i++) { $fat[$i] = -1 }
    $idx = 0
    foreach ($ds in $difat) {
        $off = ($ds + 1) * $sectorSize
        for ($i = 0; $i -lt ($sectorSize / 4) -and $idx -lt $fatEntries; $i++) {
            $fat[$idx++] = [BitConverter]::ToInt32($bytes, $off + $i * 4)
        }
    }

    $dirData = Read-SectorChain $bytes $firstDirSector $sectorSize $fat
    $entries = [System.Collections.ArrayList]::new()
    for ($i = 0; $i -lt [int]($dirData.Length / 128); $i++) {
        $eOff = $i * 128
        $nameLen = [BitConverter]::ToUInt16($dirData, $eOff + 64)
        $name = ''
        if ($nameLen -gt 2) {
            $name = [System.Text.Encoding]::Unicode.GetString($dirData, $eOff, $nameLen - 2)
        }
        [void]$entries.Add([PSCustomObject]@{
            Name = $name; ObjType = $dirData[$eOff + 66]
            Start = [BitConverter]::ToInt32($dirData, $eOff + 116)
            Size = [BitConverter]::ToUInt32($dirData, $eOff + 120)
            DirOffset = $eOff
        })
    }

    [int[]]$miniFat = @()
    if ($firstMiniFatSector -ge 0 -and $firstMiniFatSector -ne -2) {
        $mfData = Read-SectorChain $bytes $firstMiniFatSector $sectorSize $fat
        $miniFat = New-Object int[] ([int]($mfData.Length / 4))
        for ($i = 0; $i -lt $miniFat.Length; $i++) {
            $miniFat[$i] = [BitConverter]::ToInt32($mfData, $i * 4)
        }
    }

    $rootEntry = $entries | Where-Object { $_.ObjType -eq 5 } | Select-Object -First 1
    [byte[]]$miniStreamData = @()
    if ($rootEntry -and $rootEntry.Start -ge 0) {
        $miniStreamData = Read-SectorChain $bytes $rootEntry.Start $sectorSize $fat
    }

    return @{
        Entries = $entries; Bytes = $bytes; SectorSize = $sectorSize
        MiniSectorSize = $miniSectorSize; MiniStreamCutoff = $miniStreamCutoff
        Fat = $fat; MiniFat = $miniFat; MiniStreamData = $miniStreamData
        FirstDirSector = $firstDirSector
    }
}

function Read-Ole2Stream($ole2, $entry) {
    if ($entry.Size -lt $ole2.MiniStreamCutoff -and $ole2.MiniStreamData.Length -gt 0) {
        return Read-MiniStream $ole2.MiniStreamData $entry.Start $entry.Size $ole2.MiniSectorSize $ole2.MiniFat
    } else {
        $raw = Read-SectorChain $ole2.Bytes $entry.Start $ole2.SectorSize $ole2.Fat
        if ($raw.Length -gt $entry.Size) {
            $trimmed = New-Object byte[] $entry.Size
            [Array]::Copy($raw, $trimmed, $entry.Size)
            return ,$trimmed
        }
        return ,$raw
    }
}

function Write-Ole2Stream([byte[]]$ole2Bytes, $ole2, $entry, [byte[]]$newData) {
    $sectorSize = $ole2.SectorSize
    $fat = $ole2.Fat

    if ($entry.Size -lt $ole2.MiniStreamCutoff) {
        $miniSectorSize = $ole2.MiniSectorSize
        $miniFat = $ole2.MiniFat
        $s = $entry.Start; $written = 0
        while ($s -ge 0 -and $s -ne -2 -and $written -lt $newData.Length) {
            $off = $s * $miniSectorSize
            $toWrite = [Math]::Min($miniSectorSize, $newData.Length - $written)
            [Array]::Copy($newData, $written, $ole2.MiniStreamData, $off, $toWrite)
            if ($toWrite -lt $miniSectorSize) {
                for ($p = $toWrite; $p -lt $miniSectorSize; $p++) { $ole2.MiniStreamData[$off + $p] = 0 }
            }
            $written += $miniSectorSize
            if ($s -lt $miniFat.Length) { $s = $miniFat[$s] } else { break }
        }
        $rootEntry = $ole2.Entries | Where-Object { $_.ObjType -eq 5 } | Select-Object -First 1
        $s2 = $rootEntry.Start; $written2 = 0; $visited = @{}
        while ($s2 -ge 0 -and $s2 -ne -2 -and -not $visited.ContainsKey($s2) -and $written2 -lt $ole2.MiniStreamData.Length) {
            $visited[$s2] = $true
            $off2 = ($s2 + 1) * $sectorSize
            [Array]::Copy($ole2.MiniStreamData, $written2, $ole2Bytes, $off2, [Math]::Min($sectorSize, $ole2.MiniStreamData.Length - $written2))
            $written2 += $sectorSize; $s2 = $fat[$s2]
        }
        # Validate: count sectors actually used
        $chainLen = 0; $sv = $entry.Start
        while ($sv -ge 0 -and $sv -ne -2) { $chainLen++; $sv = if ($sv -lt $miniFat.Length) { $miniFat[$sv] } else { -1 } }
        $chainCapacity = $chainLen * $miniSectorSize
        if ($newData.Length -gt $chainCapacity) {
            throw "Write-Ole2Stream: data truncated (mini stream). Data=$($newData.Length) bytes, chain capacity=$chainCapacity bytes."
        }
    } else {
        $s = $entry.Start; $written = 0; $visited = @{}
        while ($s -ge 0 -and $s -ne -2 -and -not $visited.ContainsKey($s) -and $written -lt $newData.Length) {
            $visited[$s] = $true
            $off = ($s + 1) * $sectorSize
            $toWrite = [Math]::Min($sectorSize, $newData.Length - $written)
            [Array]::Copy($newData, $written, $ole2Bytes, $off, $toWrite)
            if ($toWrite -lt $sectorSize) {
                for ($p = $toWrite; $p -lt $sectorSize; $p++) { $ole2Bytes[$off + $p] = 0 }
            }
            $written += $sectorSize; $s = $fat[$s]
        }
        # Validate: count sectors actually used
        $chainLen = 0; $sv = $entry.Start; $visitedV = @{}
        while ($sv -ge 0 -and $sv -ne -2 -and -not $visitedV.ContainsKey($sv)) { $visitedV[$sv]=$true; $chainLen++; $sv = $fat[$sv] }
        $chainCapacity = $chainLen * $sectorSize
        if ($newData.Length -gt $chainCapacity) {
            throw "Write-Ole2Stream: data truncated. Data=$($newData.Length) bytes, chain capacity=$chainCapacity bytes."
        }
    }

    # Update size in directory
    $dirSectorData = Read-SectorChain $ole2Bytes $ole2.FirstDirSector $sectorSize $fat
    [Array]::Copy([BitConverter]::GetBytes([uint32]$newData.Length), 0, $dirSectorData, $entry.DirOffset + 120, 4)
    $s3 = $ole2.FirstDirSector; $written3 = 0; $visited3 = @{}
    while ($s3 -ge 0 -and $s3 -ne -2 -and -not $visited3.ContainsKey($s3)) {
        $visited3[$s3] = $true
        [Array]::Copy($dirSectorData, $written3, $ole2Bytes, ($s3 + 1) * $sectorSize, [Math]::Min($sectorSize, $dirSectorData.Length - $written3))
        $written3 += $sectorSize; $s3 = $fat[$s3]
    }
}

# ============================================================================
# VBA Decompression (MS-OVBA 2.4.1)
# ============================================================================

function Decompress-VBA([byte[]]$data, [int]$offset) {
    return ,[VbaToolkitNative]::DecompressVba($data, $offset)
}

function Compress-VBA([byte[]]$data) {
    return ,[VbaToolkitNative]::CompressVba($data)
}

# ============================================================================
# High-level helpers
# ============================================================================

function Get-VbaProjectBytes([string]$filePath) {
    $ext = [IO.Path]::GetExtension($filePath).ToLower()
    if ($ext -eq '.xls') {
        return @{ Bytes = [IO.File]::ReadAllBytes($filePath); IsZip = $false }
    }
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $block = [ScriptBlock]::Create('
        param($path)
        $zip = [System.IO.Compression.ZipFile]::OpenRead($path)
        try {
            $entry = $zip.Entries | Where-Object { $_.Name -eq "vbaProject.bin" } | Select-Object -First 1
            if (-not $entry) { return $null }
            $s = $entry.Open(); $ms = New-Object IO.MemoryStream; $s.CopyTo($ms); $s.Close()
            return ,$ms.ToArray()
        } finally { $zip.Dispose() }
    ')
    $bytes = & $block $filePath
    return @{ Bytes = $bytes; IsZip = $true }
}

function Save-VbaProjectBytes([string]$filePath, [byte[]]$ole2Bytes, [bool]$isZip) {
    if (-not $ole2Bytes -or $ole2Bytes.Length -eq 0) {
        throw "Save-VbaProjectBytes: ole2Bytes is null or empty"
    }
    if (-not (Test-Path "$filePath")) {
        throw "Save-VbaProjectBytes: target file does not exist: $filePath"
    }
    if ($isZip) {
        Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
        $zip = [System.IO.Compression.ZipFile]::Open("$filePath", [System.IO.Compression.ZipArchiveMode]::Update)
        try {
            $entry = $zip.Entries | Where-Object { $_.Name -eq 'vbaProject.bin' } | Select-Object -First 1
            if (-not $entry) { throw "Save-VbaProjectBytes: vbaProject.bin not found in ZIP" }
            $stream = $entry.Open()
            $stream.SetLength(0)
            $stream.Write($ole2Bytes, 0, $ole2Bytes.Length)
            $stream.Flush()
            $stream.Close()
        } finally {
            $zip.Dispose()
        }
    } else {
        [IO.File]::WriteAllBytes("$filePath", $ole2Bytes)
    }
    # Post-write validation: file must exist and have nonzero size
    if (-not (Test-Path "$filePath")) {
        throw "Save-VbaProjectBytes: file missing after write: $filePath"
    }
    $savedSize = (Get-Item "$filePath").Length
    if ($savedSize -eq 0) {
        throw "Save-VbaProjectBytes: file is 0 bytes after write: $filePath"
    }
}

function Get-VbaModuleList($ole2, [int]$codepage = 932) {
    $projEntry = $ole2.Entries | Where-Object { $_.Name -eq 'PROJECT' -and $_.ObjType -eq 2 } | Select-Object -First 1
    if (-not $projEntry) { return @() }
    $projData = Read-Ole2Stream $ole2 $projEntry
    $projText = [System.Text.Encoding]::GetEncoding($codepage).GetString($projData)
    $modules = [System.Collections.ArrayList]::new()
    foreach ($line in $projText -split "`r`n|`n") {
        if ($line -match '^Module=(.+)$') { [void]$modules.Add(@{ Name = $Matches[1]; Ext = 'bas'; Type = 'Module' }) }
        elseif ($line -match '^Class=(.+)$') { [void]$modules.Add(@{ Name = $Matches[1]; Ext = 'cls'; Type = 'Class' }) }
        elseif ($line -match '^BaseClass=(.+)$') { [void]$modules.Add(@{ Name = $Matches[1]; Ext = 'frm'; Type = 'Form' }) }
        elseif ($line -match '^Document=(.+?)/') { [void]$modules.Add(@{ Name = $Matches[1]; Ext = 'cls'; Type = 'Document' }) }
    }
    # Sort: Document → Form → Module → Class (matches VBA Editor tree order)
    $typeOrder = @{ 'Document' = 0; 'Form' = 1; 'Module' = 2; 'Class' = 3 }
    $modules = [System.Collections.ArrayList]@($modules | Sort-Object { $typeOrder[$_.Type] }, { $_.Name })
    return ,$modules
}

function Get-VbaModuleCode($ole2, [string]$moduleName, [int]$codepage = 932) {
    $streamEntry = $ole2.Entries | Where-Object { $_.Name -eq $moduleName -and $_.ObjType -eq 2 } | Select-Object -First 1
    if (-not $streamEntry) {
        $streamEntry = $ole2.Entries | Where-Object { $_.Name -ieq $moduleName -and $_.ObjType -eq 2 } | Select-Object -First 1
    }
    if (-not $streamEntry -or $streamEntry.Size -eq 0) { return $null }
    $streamData = Read-Ole2Stream $ole2 $streamEntry
    for ($tryOff = $streamData.Length - 2; $tryOff -ge 0; $tryOff--) {
        if ($streamData[$tryOff] -eq 0x01 -and $tryOff + 2 -lt $streamData.Length) {
            $hdr = [BitConverter]::ToUInt16($streamData, $tryOff + 1)
            if ((($hdr -shr 12) -band 0x07) -eq 3) {
                $code = Decompress-VBA $streamData $tryOff
                if ($code.Length -gt 0) {
                    $text = [System.Text.Encoding]::GetEncoding($codepage).GetString($code)
                    if ($text -match 'Attribute\s+VB_Name') {
                        return @{ Code = $text; Offset = $tryOff; Entry = $streamEntry; StreamData = $streamData }
                    }
                }
            }
        }
    }
    return $null
}

# ============================================================================
# HTML Base Template (shared dark-theme shell: CSS, sidebar, content, minimap, JS)
# ============================================================================

function New-HtmlBase {
    param(
        [string]$Title,
        [string]$Subtitle,
        [string]$ExtraCss = '',           # additional CSS rules
        [string]$SidebarHtml = '',        # sidebar inner HTML
        [string]$ContentHtml = '',        # content area inner HTML
        [string]$ExtraHtml = '',          # extra HTML inside .main after .content (e.g. outline, tooltip divs)
        [string]$ExtraJs = '',            # additional JS code
        [string]$HighlightSelector = '',  # CSS selector for minimap marks (e.g. 'tr.hl-edr')
        [int]$FirstTabIndex = 0,          # initial tab to show
        [string]$OutputPath               # file path to write
    )

    $he = { param($s) [System.Net.WebUtility]::HtmlEncode($s) }

    $html = @"
<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>$(& $he $Title)</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: Consolas, 'Courier New', monospace; font-size: 13px; background: #1e1e1e; color: #d4d4d4; overflow: hidden; }
.sidebar, .content { scrollbar-width: thin; scrollbar-color: #555 #1e1e1e; }
.header { background: #252526; padding: 10px 20px; border-bottom: 1px solid #3c3c3c; }
.header h1 { font-size: 15px; font-weight: normal; color: #cccccc; }
.header .sub { margin-top: 4px; font-size: 12px; color: #888; }
.main { display: flex; height: calc(100vh - 52px); }
.sidebar { width: 200px; min-width: 80px; background: #252526; border-right: 1px solid #3c3c3c; overflow-y: auto; padding: 8px 0; }
.sidebar .item { padding: 5px 16px; cursor: pointer; color: #888; font-size: 13px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
.sidebar .item:hover { color: #d4d4d4; background: #2a2d2e; }
.sidebar .item.active { color: #ffffff; background: #37373d; border-left: 2px solid #0078d4; }
.resizer { width: 4px; cursor: col-resize; background: #3c3c3c; flex-shrink: 0; }
.resizer:hover, .resizer.active { background: #0078d4; }
.content { flex: 1; overflow: auto; position: relative; }
.module { display: none; }
.module.active { display: block; }
.minimap { position: fixed; right: 0; top: 52px; width: 14px; bottom: 0; background: #1e1e1e; border-left: 1px solid #3c3c3c; z-index: 20; cursor: pointer; }
.minimap .mark { position: absolute; right: 2px; width: 10px; height: 3px; border-radius: 1px; }
.minimap .viewport { position: absolute; right: 0; width: 14px; background: rgba(255,255,255,0.25); border-radius: 2px; pointer-events: none; }
$ExtraCss
</style>
</head>
<body>
<div class="header">
  <h1>$(& $he $Title)</h1>
  <div class="sub">$(& $he $Subtitle)</div>
</div>
<div class="main">
<div class="sidebar" id="sidebar">
$SidebarHtml
</div>
<div class="resizer" id="resizer"></div>
<div class="content" id="content">
$ContentHtml
<div class="minimap" id="minimap"><div class="viewport" id="viewport"></div></div>
</div>
$ExtraHtml
</div>
<script>
const content = document.querySelector('.content');
const minimap = document.getElementById('minimap');
const viewport = document.getElementById('viewport');
function showTab(idx) {
  document.querySelectorAll('.module').forEach(m => m.classList.remove('active'));
  document.querySelectorAll('.item').forEach(t => t.classList.remove('active'));
  var modEl = document.getElementById('mod' + idx);
  var tabEl = document.getElementById('tab' + idx);
  if (modEl) modEl.classList.add('active');
  if (tabEl) tabEl.classList.add('active');
  content.scrollTop = 0;
  updateMinimap();
}
function updateMinimap() {
  minimap.querySelectorAll('.mark').forEach(m => m.remove());
  const mod = document.querySelector('.module.active');
  if (!mod) return;
  const hlSelector = '$HighlightSelector';
  if (!hlSelector) return;
  const rows = mod.querySelectorAll(hlSelector);
  const allRows = mod.querySelectorAll('tr');
  if (allRows.length === 0) return;
  const mapH = minimap.clientHeight;
  rows.forEach(r => {
    const idx = Array.from(allRows).indexOf(r);
    const mark = document.createElement('div');
    mark.className = 'mark' + (r.className ? ' m-' + r.className.split(' ')[0] : '');
    mark.style.top = (idx / allRows.length * mapH) + 'px';
    mark.addEventListener('click', () => r.scrollIntoView({block:'center'}));
    minimap.appendChild(mark);
  });
  updateViewport();
}
function updateViewport() {
  const sh = content.scrollHeight, ch = content.clientHeight, st = content.scrollTop;
  const mapH = minimap.clientHeight;
  if (sh <= ch) { viewport.style.display = 'none'; return; }
  viewport.style.display = '';
  viewport.style.top = (st / sh * mapH) + 'px';
  viewport.style.height = (ch / sh * mapH) + 'px';
}
content.addEventListener('scroll', updateViewport);
minimap.addEventListener('click', (e) => {
  if (e.target.classList.contains('mark')) return;
  content.scrollTop = e.offsetY / minimap.clientHeight * content.scrollHeight - content.clientHeight / 2;
});
// Resizer drag
(function() {
  const resizer = document.getElementById('resizer');
  const sidebar = document.getElementById('sidebar');
  if (!resizer || !sidebar) return;
  let startX, startW;
  resizer.addEventListener('mousedown', function(e) {
    startX = e.clientX;
    startW = sidebar.offsetWidth;
    resizer.classList.add('active');
    document.addEventListener('mousemove', onDrag);
    document.addEventListener('mouseup', onUp);
    e.preventDefault();
  });
  function onDrag(e) {
    const w = startW + (e.clientX - startX);
    sidebar.style.width = Math.max(80, Math.min(w, 600)) + 'px';
  }
  function onUp() {
    resizer.classList.remove('active');
    document.removeEventListener('mousemove', onDrag);
    document.removeEventListener('mouseup', onUp);
  }
})();
$ExtraJs
showTab($FirstTabIndex);
</script>
</body></html>
"@

    [IO.File]::WriteAllText($OutputPath, $html, [System.Text.Encoding]::UTF8)
}

# ============================================================================
# HTML Code Viewer (shared by Extract, Sanitize) — delegates to New-HtmlBase
# ============================================================================

# $moduleData: ordered hashtable of name -> @{ Ext; Lines = string[]; Highlights = @{ lineIndex -> cssClass } }
function New-HtmlCodeView {
    param(
        [string]$title,
        [string]$subtitle,
        [System.Collections.Specialized.OrderedDictionary]$moduleData,
        [string]$highlightClass,   # CSS class for highlighted lines (e.g. 'hl-edr', 'hl-sanitized')
        [string]$highlightColor,   # CSS color (e.g. '#1b3a5c' for blue, '#4b3a00' for yellow)
        [string]$highlightText,    # CSS text color
        [string]$markerColor,      # minimap marker color
        [string]$outputPath
    )

    $he = { param($s) [System.Net.WebUtility]::HtmlEncode($s) }

    # --- Extra CSS ---
    $extraCss = @"
.sidebar .item.has-hl { color: $markerColor; }
.sidebar .item.no-hl { color: #606060; }
.code-table { width: 100%; border-collapse: collapse; }
.code-table td { padding: 0 8px; line-height: 20px; vertical-align: top; white-space: pre; overflow: hidden; text-overflow: ellipsis; }
.code-table .ln { width: 50px; min-width: 50px; text-align: right; color: #606060; padding-right: 12px; user-select: none; border-right: 1px solid #3c3c3c; }
.code-table .code { color: #d4d4d4; }
tr.$highlightClass td.code { background: $highlightColor; color: $highlightText; }
tr.$highlightClass td.ln { color: #cccccc; }
.minimap .mark { background: $markerColor; }
"@

    # --- Sidebar ---
    $sidebarSb = [System.Text.StringBuilder]::new()
    $tabIdx = 0; $firstHlIdx = -1
    foreach ($modName in $moduleData.Keys) {
        $md = $moduleData[$modName]
        $hlCount = 0
        if ($md.Highlights) { $hlCount = $md.Highlights.Count }
        $cls = if ($hlCount -gt 0) { 'has-hl' } else { 'no-hl' }
        if ($firstHlIdx -eq -1 -and $hlCount -gt 0) { $firstHlIdx = $tabIdx }
        $label = "$modName.$($md.Ext)"
        if ($hlCount -gt 0) { $label += " ($hlCount)" }
        [void]$sidebarSb.Append("<div class=`"item $cls`" onclick=`"showTab($tabIdx)`" id=`"tab$tabIdx`">$(& $he $label)</div>")
        $tabIdx++
    }
    if ($firstHlIdx -eq -1) { $firstHlIdx = 0 }

    # --- Content ---
    $contentSb = [System.Text.StringBuilder]::new()
    $tabIdx = 0
    foreach ($modName in $moduleData.Keys) {
        $md = $moduleData[$modName]
        [void]$contentSb.Append("<div class=`"module`" id=`"mod$tabIdx`"><table class=`"code-table`">")
        for ($i = 0; $i -lt $md.Lines.Count; $i++) {
            $trClass = ''
            if ($md.Highlights -and $md.Highlights.ContainsKey($i)) {
                $hlVal = $md.Highlights[$i]
                if ($hlVal -is [string]) { $trClass = $hlVal } else { $trClass = $highlightClass }
            }
            $ln = $i + 1
            $code = & $he $md.Lines[$i]
            [void]$contentSb.Append("<tr class=`"$trClass`"><td class=`"ln`">$ln</td><td class=`"code`">$code</td></tr>")
        }
        [void]$contentSb.Append("</table></div>")
        $tabIdx++
    }

    # Collect all unique highlight classes from module data for minimap selector
    $hlClasses = [System.Collections.ArrayList]::new()
    foreach ($modName in $moduleData.Keys) {
        $md = $moduleData[$modName]
        if ($md.Highlights) {
            foreach ($v in $md.Highlights.Values) {
                $cls = if ($v -is [string]) { $v } else { $highlightClass }
                if ($hlClasses -notcontains $cls) { [void]$hlClasses.Add($cls) }
            }
        }
    }
    if ($hlClasses.Count -eq 0) { [void]$hlClasses.Add($highlightClass) }
    $hlSelector = ($hlClasses | ForEach-Object { "tr.$_" }) -join ', '

    New-HtmlBase -Title $title -Subtitle $subtitle `
        -ExtraCss $extraCss -SidebarHtml $sidebarSb.ToString() -ContentHtml $contentSb.ToString() `
        -HighlightSelector $hlSelector -FirstTabIndex $firstHlIdx -OutputPath $outputPath
}

# ============================================================================
# Input validation
# ============================================================================

function Resolve-VbaFilePath {
    param([string]$Path, [string[]]$Supported = @('.xls','.xlsm','.xlam'))
    if (-not (Test-Path -LiteralPath $Path)) { throw "File not found: $Path" }
    $resolved = (Resolve-Path -LiteralPath $Path).Path
    $ext = [IO.Path]::GetExtension($resolved).ToLower()
    if ($ext -notin $Supported) { throw "Unsupported format: $ext (supported: $($Supported -join ', '))" }
    return $resolved
}

# ============================================================================
# Output management
# ============================================================================

function New-VbaOutputDir {
    param([string]$InputFilePath, [string]$ToolName)
    $scriptRoot = Split-Path "$PSScriptRoot" -Parent  # vba-devkit root
    $outputRoot = Join-Path $scriptRoot 'output'
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $runDir = Join-Path $outputRoot "${timestamp}_${ToolName}"
    [void][IO.Directory]::CreateDirectory($runDir)
    return $runDir
}

# ============================================================================
# Logging and terminal display
# ============================================================================

function Write-VbaLog {
    param([string]$ToolName, [string]$InputFile, [string]$Message, [string]$Level = 'INFO')
    $logPath = Join-Path $PSScriptRoot '..\vba-toolkit.log'
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $fileName = [IO.Path]::GetFileName($InputFile)
    $entry = "[$timestamp] [$Level] [$ToolName] $fileName | $Message"
    [IO.File]::AppendAllText($logPath, "$entry`r`n", [System.Text.Encoding]::UTF8)
}

function Write-VbaStatus {
    param([string]$ToolName, [string]$FileName, [string]$Message)
    Write-Host "  $Message" -ForegroundColor Gray
}

function Write-VbaResult {
    param([string]$ToolName, [string]$FileName, [string]$Summary, [string]$OutputDir, [double]$ElapsedSec)
    Write-Host "  $Summary" -ForegroundColor Green
    if ($OutputDir) { Write-Host "  Output: $OutputDir" -ForegroundColor Gray }
    if ($ElapsedSec -gt 0) { Write-Host "  Done ($([Math]::Round($ElapsedSec, 1))s)" -ForegroundColor Gray }
}

function Write-VbaError {
    param([string]$ToolName, [string]$FileName, [string]$Message)
    Write-Host "  ERROR: $Message" -ForegroundColor Red
    Write-VbaLog $ToolName $FileName $Message 'ERROR'
}

function Write-VbaHeader {
    param([string]$ToolName, [string]$FileName)
    Write-Host "[$ToolName] $FileName" -ForegroundColor Cyan
}

# ============================================================================
# Codepage detection
# ============================================================================

function Get-VbaCodepage($ole2) {
    $codepage = 932  # default fallback (Shift-JIS)
    try {
        $dirEntry = $ole2.Entries | Where-Object { $_.Name -eq 'dir' } | Select-Object -First 1
        if (-not $dirEntry) { return $codepage }
        $dirRaw = Read-Ole2Stream $ole2 $dirEntry
        $dirData = Decompress-VBA $dirRaw 0
        if ($dirData.Length -lt 8) { return $codepage }
        # Scan for PROJECTCODEPAGE record (ID=0x0003)
        $pos = 0
        while ($pos + 6 -le $dirData.Length) {
            $id = [BitConverter]::ToUInt16($dirData, $pos)
            $size = [BitConverter]::ToInt32($dirData, $pos + 2)
            if ($size -lt 0 -or $pos + 6 + $size -gt $dirData.Length) { break }
            if ($id -eq 0x0003 -and $size -ge 2) {
                $codepage = [BitConverter]::ToUInt16($dirData, $pos + 6)
                break
            }
            $pos += 6 + $size
            # Handle PROJECTVERSION (0x0009) non-standard format: 2 extra bytes for MinorVersion
            if ($id -eq 0x0009) { $pos += 2 }
        }
    } catch {}
    # Validate codepage
    try { [System.Text.Encoding]::GetEncoding($codepage) | Out-Null } catch { $codepage = 932 }
    return $codepage
}

# ============================================================================
# Bulk module extraction
# ============================================================================

function Get-AllModuleCode {
    param(
        [string]$FilePath,
        [switch]$StripAttributes,
        [switch]$IncludeRawData
    )
    $proj = Get-VbaProjectBytes $FilePath
    if (-not $proj.Bytes) { return $null }
    $ole2 = Read-Ole2 $proj.Bytes
    $codepage = Get-VbaCodepage $ole2
    $encoding = [System.Text.Encoding]::GetEncoding($codepage)
    $modules = Get-VbaModuleList $ole2 $codepage
    $result = [ordered]@{}
    foreach ($mod in $modules) {
        $mc = Get-VbaModuleCode $ole2 $mod.Name $codepage
        if (-not $mc) { continue }
        $rawBytes = Decompress-VBA $mc.StreamData $mc.Offset
        $code = $encoding.GetString($rawBytes)
        $lines = $code -split "`r`n|`n"
        if ($StripAttributes) {
            $lines = @($lines | Where-Object { $_ -notmatch '^\s*Attribute\s+VB_' })
        }
        $entry = @{ Code = ($lines -join "`n"); Ext = $mod.Ext; Lines = $lines; Name = $mod.Name }
        if ($IncludeRawData) {
            $entry.Entry = $mc.Entry
            $entry.Offset = $mc.Offset
            $entry.StreamData = $mc.StreamData
        }
        $result[$mod.Name] = $entry
    }
    return @{ Modules = $result; Ole2 = $ole2; Ole2Bytes = $proj.Bytes; IsZip = $proj.IsZip; Codepage = $codepage; FilePath = $FilePath }
}

# ============================================================================
# Analysis engine (shared by Extract, Cheatsheet)
# ============================================================================

function Get-VbaAnalysis {
    param(
        [hashtable]$Project  # result of Get-AllModuleCode
    )

    $allCode = @{}  # fileName -> lines array
    foreach ($modName in $Project.Modules.Keys) {
        $mod = $Project.Modules[$modName]
        $allCode["$modName.$($mod.Ext)"] = $mod.Lines
    }

    # --- EDR detection patterns ---
    $patterns = [ordered]@{
        'Win32 API (Declare)' = @{
            Pattern = '(?m)^[^''\r\n]*\bDeclare\s+(PtrSafe\s+)?(Function|Sub)\s+(\w+)'
            Extract = { param($m) "$($m.Groups[3].Value) ($(if($m.Groups[1].Value){'PtrSafe'} else {'Legacy'}))" }
        }
        'COM / CreateObject' = @{ Pattern = '(?m)^[^''\r\n]*\bCreateObject\s*\(\s*"([^"]+)"'; Extract = { param($m) $m.Groups[1].Value } }
        'COM / GetObject' = @{ Pattern = '(?m)^[^''\r\n]*\bGetObject\s*\(\s*"?([^")\s]+)"?'; Extract = { param($m) $m.Groups[1].Value } }
        'Shell / process' = @{ Pattern = '(?m)^[^''\r\n]*\b(Shell\s*[\("]|WScript\.Shell|cmd\s*/[ck])'; Extract = { param($m) $m.Groups[1].Value.Trim() } }
        'File I/O' = @{
            Pattern = '(?m)^[^''\r\n]*\b(Open\s+\S+\s+For\s+(Input|Output|Append|Binary|Random)|Kill\s|FileCopy\s|MkDir\s|RmDir\s)'
            Extract = { param($m) if ($m.Groups[2].Value) { "Open For $($m.Groups[2].Value)" } else { $m.Groups[1].Value.Trim() } }
        }
        'FileSystemObject' = @{ Pattern = '(?m)^[^''\r\n]*\b(Scripting\.FileSystemObject)\b'; Extract = { param($m) $m.Groups[1].Value } }
        'Registry' = @{ Pattern = '(?m)^[^''\r\n]*\b(GetSetting|SaveSetting|DeleteSetting|RegRead|RegWrite|RegDelete)\b'; Extract = { param($m) $m.Groups[1].Value } }
        'SendKeys' = @{ Pattern = '(?m)^[^''\r\n]*\b(SendKeys)\b'; Extract = { param($m) $m.Groups[1].Value } }
        'Network / HTTP' = @{ Pattern = '(?m)^[^''\r\n]*\b(MSXML2\.XMLHTTP|WinHttp\.WinHttpRequest|URLDownloadToFile|MSXML2\.ServerXMLHTTP)\b'; Extract = { param($m) $m.Groups[1].Value } }
        'PowerShell / WScript' = @{ Pattern = '(?mi)^[^''\r\n]*\b(powershell|wscript|cscript|mshta)\b'; Extract = { param($m) $m.Groups[1].Value } }
        'Process / WMI' = @{ Pattern = '(?m)^[^''\r\n]*\b(winmgmts|Win32_Process|WbemScripting|ExecQuery)\b'; Extract = { param($m) $m.Groups[1].Value } }
        'DLL loading' = @{ Pattern = '(?m)^[^''\r\n]*\b(LoadLibrary|GetProcAddress|FreeLibrary|CallByName)\b'; Extract = { param($m) $m.Groups[1].Value } }
        'Clipboard' = @{ Pattern = '(?m)^[^''\r\n]*\b(MSForms\.DataObject|GetClipboardData|SetClipboardData)\b'; Extract = { param($m) $m.Groups[1].Value } }
        'Environment' = @{ Pattern = '(?m)^[^''\r\n]*\b(Environ\s*\$?\s*\()'; Extract = { param($m) "Environ" } }
        'Auto-execution' = @{ Pattern = '(?m)^\s*(Sub\s+(Auto_Open|Auto_Close|Workbook_Open|Workbook_BeforeClose|Document_Open|Document_Close)\b)'; Extract = { param($m) $m.Groups[2].Value } }
        'Encoding / obfuscation' = @{ Pattern = '(?m)^[^''\r\n]*\b(Chr\s*\$?\s*\(\s*\d+\s*\))'; Extract = { param($m) $m.Groups[1].Value }; Aggregate = $true }
    }

    # --- Compatibility risk patterns ---
    $compatPatterns = [ordered]@{
        '64-bit: Missing PtrSafe' = @{
            Pattern = '(?m)^[^''\r\n]*\bDeclare\s+(?!PtrSafe\b)(Function|Sub)\s+(\w+)'
            Extract = { param($m) "$($m.Groups[2].Value) -- missing PtrSafe" }
        }
        '64-bit: Long for handles' = @{
            Pattern = '(?mi)^[^''\r\n]*\bDeclare\s+PtrSafe\s+(?:Function|Sub)\s+\w+[^''\n]*\bAs\s+Long\b'
            Extract = { param($m) "As Long in PtrSafe Declare -- review for LongPtr" }
        }
        '64-bit: VarPtr/ObjPtr/StrPtr' = @{
            Pattern = '(?mi)^[^''\r\n]*\b(VarPtr|ObjPtr|StrPtr)\s*\('
            Extract = { param($m) "$($m.Groups[1].Value) -- returns LongPtr on 64-bit" }
        }
        'Deprecated: DDE' = @{
            Pattern = '(?mi)^[^''\r\n]*\b(DDEInitiate|DDEExecute|DDEPoke|DDERequest|DDETerminate|DDETerminateAll)\b'
            Extract = { param($m) "$($m.Groups[1].Value)" }
        }
        'Deprecated: IE Automation' = @{
            Pattern = '(?mi)^[^''\r\n]*\bInternetExplorer\.Application\b'
            Extract = { param($m) "InternetExplorer.Application -- IE removed" }
        }
        'Deprecated: Legacy Controls' = @{
            Pattern = '(?mi)^[^''\r\n]*\b(MSCAL\.Calendar|MSComDlg\.CommonDialog|MSComctlLib\.\w+|COMDLG32\.OCX)\b'
            Extract = { param($m) "$($m.Groups[1].Value) -- no 64-bit version" }
        }
        'Deprecated: DAO' = @{
            Pattern = '(?mi)^[^''\r\n]*\b(DAO\.Database|DAO\.Recordset|DBEngine\b|DAO\.QueryDef)'
            Extract = { param($m) "$($m.Groups[1].Value)" }
        }
        'Legacy: DefType' = @{
            Pattern = '(?mi)^\s*(Def(Bool|Byte|Int|Lng|LngLng|LngPtr|Cur|Sng|Dbl|Dec|Date|Str|Obj|Var))\s+'
            Extract = { param($m) "$($m.Groups[1].Value)" }
        }
        'Legacy: GoSub' = @{
            Pattern = '(?mi)^[^''\r\n]*\bGoSub\b'
            Extract = { param($m) "GoSub -- refactor to Sub/Function" }
        }
        'Legacy: While/Wend' = @{
            Pattern = '(?mi)^\s*Wend\b'
            Extract = { param($m) "While...Wend -- use Do While...Loop" }
        }
    }

    # --- Environment dependency patterns (Risk) ---
    $envPatterns = [ordered]@{
        'Fixed drive letter' = @{
            Pattern = '(?mi)^[^''\r\n]*"[A-Z]:\\"'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Risk'
        }
        'UNC path' = @{
            Pattern = '(?mi)^[^''\r\n]*"\\\\[^"]+\\'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Risk'
        }
        'User folder' = @{
            Pattern = '(?mi)^[^''\r\n]*C:\\Users\\'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Risk'
        }
        'Desktop / Documents' = @{
            Pattern = '(?mi)^[^''\r\n]*\\(Desktop|Documents)\\'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Risk'
        }
        'AppData' = @{
            Pattern = '(?mi)^[^''\r\n]*\\AppData\\'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Risk'
        }
        'Program Files' = @{
            Pattern = '(?mi)^[^''\r\n]*\\Program Files'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Risk'
        }
        'Fixed printer name' = @{
            Pattern = '(?mi)^[^''\r\n]*\.ActivePrinter\s*=\s*"'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Risk'
        }
        'Fixed IP address' = @{
            Pattern = '(?mi)^[^''\r\n]*"\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}"'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Risk'
        }
        'Fixed connection host' = @{
            Pattern = '(?mi)^[^''\r\n]*(Server\s*=|Host\s*=)'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Risk'
        }
        'localhost' = @{
            Pattern = '(?mi)^[^''\r\n]*\blocalhost\b'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Review'
        }
        'Connection string' = @{
            Pattern = '(?mi)^[^''\r\n]*(Provider\s*=|DSN\s*=|Data\s+Source\s*=)'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Risk'
        }
        'External workbook open (literal)' = @{
            Pattern = '(?mi)^[^''\r\n]*\bWorkbooks\.Open\s*\(\s*"'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Risk'
        }
        'Dir() path check' = @{
            Pattern = '(?mi)^[^''\r\n]*\bDir\s*\('
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Risk'
        }
        'Path concatenation' = @{
            Pattern = '(?mi)^[^''\r\n]*\.Path\s*&\s*"[/\\]'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Risk'
        }
        'SaveAs call' = @{
            Pattern = '(?mi)^[^''\r\n]*\.SaveAs\b'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Risk'
        }
        'External workbook ref' = @{
            Pattern = '(?mi)^[^''\r\n]*\[''[^''\]]+\.xls[^''\]]*''\]'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Risk'
        }
        'BeforeSave event' = @{
            Pattern = '(?mi)^\s*(Sub|Private\s+Sub)\s+(Workbook_BeforeSave|BeforeSave)\b'
            Extract = { param($m) $m.Groups[2].Value }
            Severity = 'Review'
        }
        'AfterSave event' = @{
            Pattern = '(?mi)^\s*(Sub|Private\s+Sub)\s+(Workbook_AfterSave|AfterSave)\b'
            Extract = { param($m) $m.Groups[2].Value }
            Severity = 'Review'
        }
        'LinkSources / UpdateLink' = @{
            Pattern = '(?mi)^[^''\r\n]*\.(LinkSources|UpdateLink)\b'
            Extract = { param($m) $m.Groups[1].Value }
            Severity = 'Review'
        }
        'Workbooks.Open (variable)' = @{
            Pattern = '(?mi)^[^''\r\n]*\bWorkbooks\.Open\s*[\( ]\s*[A-Za-z_]\w*'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Review'
        }
        'CurDir' = @{
            Pattern = '(?mi)^[^''\r\n]*\bCurDir\b'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Review'
        }
        'ChDir' = @{
            Pattern = '(?mi)^[^''\r\n]*\bChDir\b'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Review'
        }
    }

    # --- Environment Info patterns (no highlight, text report only) ---
    $envInfoPatterns = [ordered]@{
        'ThisWorkbook.Path' = @{
            Pattern = '(?mi)^[^''\r\n]*\bThisWorkbook\.Path\b'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Info'
        }
        'ActiveWorkbook.Path' = @{
            Pattern = '(?mi)^[^''\r\n]*\bActiveWorkbook\.Path\b'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Info'
        }
        'ThisWorkbook.FullName' = @{
            Pattern = '(?mi)^[^''\r\n]*\bThisWorkbook\.FullName\b'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Info'
        }
        'ActiveWorkbook.FullName' = @{
            Pattern = '(?mi)^[^''\r\n]*\bActiveWorkbook\.FullName\b'
            Extract = { param($m) $m.Value.Trim() }
            Severity = 'Info'
        }
    }

    # --- Business dependency patterns ---
    $bizPatterns = [ordered]@{
        'Outlook integration' = @{
            Pattern = '(?mi)^[^''\r\n]*\bOutlook\.Application\b'
            Extract = { param($m) $m.Value.Trim() }
        }
        'Word integration' = @{
            Pattern = '(?mi)^[^''\r\n]*\bWord\.Application\b'
            Extract = { param($m) $m.Value.Trim() }
        }
        'Access / DB integration' = @{
            Pattern = '(?mi)^[^''\r\n]*\b(Access\.Application|CurrentDb|DoCmd)\b'
            Extract = { param($m) $m.Groups[1].Value }
        }
        'PDF export' = @{
            Pattern = '(?mi)^[^''\r\n]*\.ExportAsFixedFormat\b'
            Extract = { param($m) $m.Value.Trim() }
        }
        'Print' = @{
            Pattern = '(?mi)^[^''\r\n]*\.(PrintOut|PrintPreview)\b'
            Extract = { param($m) $m.Groups[1].Value }
        }
        'External EXE' = @{
            Pattern = '(?mi)^[^''\r\n]*\bShell\s.*\.exe'
            Extract = { param($m) $m.Value.Trim() }
        }
    }

    # --- Pattern matching ---
    $findings = [ordered]@{}
    $issueCount = 0
    foreach ($cat in $patterns.Keys) {
        $p = $patterns[$cat]
        $catFindings = [System.Collections.ArrayList]::new()
        foreach ($fn in $allCode.Keys) {
            $content = $allCode[$fn] -join "`n"
            foreach ($m in [regex]::Matches($content, $p.Pattern)) {
                [void]$catFindings.Add("${fn}: $(& $p.Extract $m)")
            }
        }
        if ($catFindings.Count -gt 0) {
            $issueCount += $catFindings.Count
            $findings[$cat] = @{ Findings = $catFindings; Aggregate = [bool]$p.Aggregate }
        }
    }

    # --- Compatibility pattern matching ---
    $compatFindings = [ordered]@{}
    $compatIssueCount = 0
    foreach ($cat in $compatPatterns.Keys) {
        $p = $compatPatterns[$cat]
        $catFindings = [System.Collections.ArrayList]::new()
        foreach ($fn in $allCode.Keys) {
            $content = $allCode[$fn] -join "`n"
            foreach ($m in [regex]::Matches($content, $p.Pattern)) {
                [void]$catFindings.Add("${fn}: $(& $p.Extract $m)")
            }
        }
        if ($catFindings.Count -gt 0) {
            $compatIssueCount += $catFindings.Count
            $compatFindings[$cat] = @{ Findings = $catFindings; Aggregate = [bool]$p.Aggregate }
        }
    }

    # --- Environment dependency pattern matching (Risk + Review) ---
    $envFindings = [ordered]@{}
    $envIssueCount = 0
    $envRiskCount = 0
    $envReviewCount = 0
    foreach ($cat in $envPatterns.Keys) {
        $p = $envPatterns[$cat]
        $catFindings = [System.Collections.ArrayList]::new()
        foreach ($fn in $allCode.Keys) {
            $content = $allCode[$fn] -join "`n"
            foreach ($m in [regex]::Matches($content, $p.Pattern)) {
                [void]$catFindings.Add("${fn}: $(& $p.Extract $m)")
            }
        }
        if ($catFindings.Count -gt 0) {
            $envIssueCount += $catFindings.Count
            $sev = if ($p.Severity) { $p.Severity } else { 'Risk' }
            if ($sev -eq 'Risk') { $envRiskCount += $catFindings.Count }
            elseif ($sev -eq 'Review') { $envReviewCount += $catFindings.Count }
            $envFindings[$cat] = @{ Findings = $catFindings; Aggregate = [bool]$p.Aggregate; Severity = $sev }
        }
    }

    # --- Environment Info pattern matching ---
    $infoFindings = [ordered]@{}
    $infoCount = 0
    foreach ($cat in $envInfoPatterns.Keys) {
        $p = $envInfoPatterns[$cat]
        $catFindings = [System.Collections.ArrayList]::new()
        foreach ($fn in $allCode.Keys) {
            $content = $allCode[$fn] -join "`n"
            foreach ($m in [regex]::Matches($content, $p.Pattern)) {
                [void]$catFindings.Add("${fn}: $(& $p.Extract $m)")
            }
        }
        if ($catFindings.Count -gt 0) {
            $infoCount += $catFindings.Count
            $infoFindings[$cat] = @{ Findings = $catFindings; Aggregate = $false; Severity = 'Info' }
        }
    }

    # --- Business dependency pattern matching ---
    $bizFindings = [ordered]@{}
    $bizIssueCount = 0
    foreach ($cat in $bizPatterns.Keys) {
        $p = $bizPatterns[$cat]
        $catFindings = [System.Collections.ArrayList]::new()
        foreach ($fn in $allCode.Keys) {
            $content = $allCode[$fn] -join "`n"
            foreach ($m in [regex]::Matches($content, $p.Pattern)) {
                [void]$catFindings.Add("${fn}: $(& $p.Extract $m)")
            }
        }
        if ($catFindings.Count -gt 0) {
            $bizIssueCount += $catFindings.Count
            $bizFindings[$cat] = @{ Findings = $catFindings; Aggregate = [bool]$p.Aggregate }
        }
    }

    # --- COM object tracking ---
    $comBindings = [System.Collections.ArrayList]::new()
    foreach ($fn in $allCode.Keys) {
        $lines = $allCode[$fn]
        for ($li = 0; $li -lt $lines.Count; $li++) {
            if ($lines[$li] -match '^\s*''') { continue }
            if ($lines[$li] -match '\bSet\s+(\w+)\s*=\s*CreateObject\s*\(\s*"([^"]+)"') {
                [void]$comBindings.Add(@{ VarName = $Matches[1]; ProgId = $Matches[2]; File = $fn; Line = $li + 1 })
            } elseif ($lines[$li] -match '\bSet\s+(\w+)\s*=\s*GetObject\s*\(') {
                [void]$comBindings.Add(@{ VarName = $Matches[1]; ProgId = '(GetObject)'; File = $fn; Line = $li + 1 })
            }
        }
    }

    # --- API declarations + call sites ---
    $apiDecls = [System.Collections.ArrayList]::new()
    $apiCallNames = [System.Collections.ArrayList]::new()
    foreach ($fn in $allCode.Keys) {
        $lines = $allCode[$fn]
        for ($li = 0; $li -lt $lines.Count; $li++) {
            if ($lines[$li] -match '^\s*''') { continue }
            if ($lines[$li] -match '(?i)\bDeclare\s+(PtrSafe\s+)?(Function|Sub)\s+(\w+)\s+Lib\s+(.*)') {
                $trimmed = $lines[$li].Trim(); if ($trimmed.Length -gt 100) { $trimmed = $trimmed.Substring(0, 97) + '...' }
                [void]$apiDecls.Add(@{ Name = $Matches[3]; File = $fn; Line = $li + 1; Sig = $trimmed })
                if ($apiCallNames -notcontains $Matches[3]) { [void]$apiCallNames.Add($Matches[3]) }
            }
        }
    }

    # COM variable names for highlight
    $comVarNames = [System.Collections.ArrayList]::new()
    foreach ($b in $comBindings) { if ($comVarNames -notcontains $b.VarName) { [void]$comVarNames.Add($b.VarName) } }

    # --- External references from PROJECT stream ---
    $externalRefs = [System.Collections.ArrayList]::new()
    if ($Project.Ole2) {
        $projEntry = $Project.Ole2.Entries | Where-Object { $_.Name -eq 'PROJECT' -and $_.ObjType -eq 2 } | Select-Object -First 1
        if ($projEntry) {
            $cp = if ($Project.Codepage) { $Project.Codepage } else { 932 }
            $projData = Read-Ole2Stream $Project.Ole2 $projEntry
            $projText = [System.Text.Encoding]::GetEncoding($cp).GetString($projData)
            foreach ($line in $projText -split "`r`n|`n") {
                if ($line -match '^Reference=' -and $line -match '#([^#]+)$') {
                    [void]$externalRefs.Add($Matches[1])
                }
            }
        }
    }

    return @{
        Patterns = $patterns
        Findings = $findings
        IssueCount = $issueCount
        CompatPatterns = $compatPatterns
        CompatFindings = $compatFindings
        CompatIssueCount = $compatIssueCount
        EnvPatterns = $envPatterns
        EnvFindings = $envFindings
        EnvIssueCount = $envIssueCount
        EnvRiskCount = $envRiskCount
        EnvReviewCount = $envReviewCount
        EnvInfoPatterns = $envInfoPatterns
        InfoFindings = $infoFindings
        InfoCount = $infoCount
        BizPatterns = $bizPatterns
        BizFindings = $bizFindings
        BizIssueCount = $bizIssueCount
        ComBindings = $comBindings
        ComVarNames = $comVarNames
        ApiDecls = $apiDecls
        ApiCallNames = $apiCallNames
        AllCode = $allCode
        ExternalRefs = $externalRefs
    }
}

# ============================================================================
# API Replacement Database (moved from Cheatsheet.ps1)
# Covers ALL 26 detection patterns (16 EDR + 10 compat)
# ============================================================================

$script:VbaApiReplacements = [ordered]@{
    # --- Timer / Sleep ---
    'GetTickCount' = @{
        Lib = 'kernel32'
        Alt = 'Timer (VBA built-in, Single type)'
        Example = @'
' Before:
Dim t As Long: t = GetTickCount()
DoSomething
Debug.Print "Elapsed: " & (GetTickCount() - t) & " ms"

' After:
Dim t As Single: t = Timer
DoSomething
Dim elapsed As Single: elapsed = Timer - t
If elapsed < 0 Then elapsed = elapsed + 86400  ' midnight rollover
Debug.Print "Elapsed: " & Format(elapsed * 1000, "0") & " ms"
'@
        Note = 'Timer is Single (~15ms resolution), resets at midnight. Add 86400 if elapsed < 0. GetTickCount wraps at ~49.7 days.'
    }
    'GetTickCount64' = @{
        Lib = 'kernel32'
        Alt = 'Timer (VBA built-in)'
        Example = '(Same as GetTickCount)'
        Note = ''
    }
    'Sleep' = @{
        Lib = 'kernel32'
        Alt = 'Application.Wait (Excel only) or DoEvents loop'
        Example = @'
' Before:
Sleep 1000  ' 1 second

' After (Option A - Excel only, 1sec resolution):
Application.Wait Now + TimeSerial(0, 0, 1)

' After (Option B - any host, sub-second, busy-wait):
Dim endTime As Single: endTime = Timer + 0.5  ' 500ms
Do While Timer < endTime: DoEvents: Loop
' Note: DoEvents loop uses 100% CPU on one core

' After (Option C - non-blocking delayed execution):
Application.OnTime Now + TimeSerial(0, 0, 1), "MyCallback"
'@
        Note = 'Application.Wait is Excel-only (not Word/Access/Outlook). DoEvents loop is a busy-wait. Application.OnTime is non-blocking but requires a callback Sub.'
    }
    'timeGetTime' = @{
        Lib = 'winmm'
        Alt = 'Timer (VBA built-in)'
        Example = '(Same as GetTickCount)'
        Note = ''
    }
    'QueryPerformanceCounter' = @{
        Lib = 'kernel32'
        Alt = 'No equivalent for high-resolution timing. Timer (~15ms) for rough measurements.'
        Example = @'
' Before:
QueryPerformanceCounter startCount
DoSomething
QueryPerformanceCounter endCount
elapsed = (endCount - startCount) / freq

' After (rough timing only):
Dim t As Single: t = Timer
DoSomething
Debug.Print "Elapsed: " & Format((Timer - t) * 1000, "0") & " ms"
'@
        Note = 'QPC provides sub-microsecond precision. Timer provides ~15ms at best. No pure VBA equivalent for high-resolution timing.'
    }
    'QueryPerformanceFrequency' = @{
        Lib = 'kernel32'
        Alt = '(Remove together with QueryPerformanceCounter)'
        Example = ''
        Note = ''
    }

    # --- String / Memory ---
    'CopyMemory' = @{
        Lib = 'kernel32 (RtlMoveMemory)'
        Alt = 'LSet (UDT copy), array assignment, or byte-by-byte copy'
        Example = @'
' Before:
CopyMemory ByVal dest, ByVal src, length

' After (byte arrays - direct assignment):
destBytes() = sourceBytes()

' After (byte-by-byte):
Dim i As Long
For i = 0 To length - 1
    dest(i) = src(i)
Next i

' After (UDT to UDT of same size):
LSet destUDT = sourceUDT

' After (in-place string modification):
Mid$(dest, pos, length) = Mid$(src, 1, length)
'@
        Note = 'LSet copies between UDTs of the same size without API. Mid$ as a statement (left-hand side) modifies strings in-place.'
    }
    'lstrlen' = @{
        Lib = 'kernel32'
        Alt = 'Len / LenB (VBA built-in)'
        Example = @'
' Before:
length = lstrlen(ByVal ptr)

' After:
length = Len(str)     ' character count
length = LenB(str)    ' byte count (= Len * 2 in VBA Unicode)
'@
        Note = 'If original code used lstrlen for ANSI buffer sizing, use LenB instead of Len.'
    }

    # --- User / System info ---
    'GetUserName' = @{
        Lib = 'advapi32'
        Alt = 'Environ$("USERNAME") for Windows login name'
        Example = @'
' Before:
Dim buf As String: buf = Space(256)
Dim sz As Long: sz = 256
GetUserName buf, sz
userName = Left$(buf, sz - 1)

' After (Windows login name):
userName = Environ$("USERNAME")

' CAUTION: Application.UserName is the Office display name,
' NOT the Windows login. These are often different in
' corporate environments.
'@
        Note = 'Environ$("USERNAME") = Windows login. Application.UserName = Office display name. These differ in corporate environments.'
    }
    'GetComputerName' = @{
        Lib = 'kernel32'
        Alt = 'Environ$("COMPUTERNAME")'
        Example = @'
' Before:
Dim buf As String: buf = Space(256)
Dim sz As Long: sz = 256
GetComputerName buf, sz
compName = Left$(buf, sz - 1)

' After:
compName = Environ$("COMPUTERNAME")
'@
        Note = 'Environ$ returns empty string if variable is not set. Always validate the result.'
    }
    'GetTempPath' = @{
        Lib = 'kernel32'
        Alt = 'Environ$("TEMP")'
        Example = @'
' Before:
Dim buf As String: buf = Space(260)
GetTempPath 260, buf
tmpPath = Left$(buf, InStr(buf, vbNullChar) - 1)

' After:
tmpPath = Environ$("TEMP") & "\"
' Note: API appends trailing "\", Environ$ does not.
'@
        Note = 'GetTempPath appends trailing backslash. Environ$("TEMP") does not. Add "\" when concatenating paths.'
    }
    'GetSystemDirectory' = @{
        Lib = 'kernel32'
        Alt = 'Environ$("WINDIR") & "\System32"'
        Example = ''
        Note = 'Caution: 32-bit Office on 64-bit Windows uses SysWOW64. Environ$ always returns System32. Behavior may differ from the original API call.'
    }
    'GetWindowsDirectory' = @{
        Lib = 'kernel32'
        Alt = 'Environ$("WINDIR")'
        Example = ''
        Note = ''
    }

    # --- Window / UI ---
    'FindWindow' = @{
        Lib = 'user32'
        Alt = 'Application.hWnd (own window only, Excel 2010+) or AppActivate'
        Example = @'
' Before:
hWnd = FindWindow(vbNullString, "Window Title")

' After (get own Excel window handle):
hWnd = Application.hWnd  ' Excel 2010+

' After (activate by title - no handle returned):
AppActivate "Window Title"
'@
        Note = 'Application.hWnd only returns the host app window handle. FindWindow for other app windows has no VBA equivalent. AppActivate does partial title matching - may activate wrong window.'
    }
    'SetWindowPos' = @{
        Lib = 'user32'
        Alt = 'UserForm position properties (positioning only, no topmost)'
        Example = @'
' Before:
SetWindowPos hWnd, HWND_TOPMOST, x, y, w, h, SWP_NOSIZE

' After (UserForm positioning only):
Me.StartUpPosition = 0
Me.Left = x: Me.Top = y

' For "stay visible" behavior:
frm.Show vbModeless
'@
        Note = 'No VBA equivalent for HWND_TOPMOST. Positioning alternative applies to UserForms only, not the application window. vbModeless keeps form visible while user works.'
    }
    'GetSystemMetrics' = @{
        Lib = 'user32'
        Alt = 'Application.UsableWidth/Height (Excel, in points) - no pixel equivalent'
        Example = @'
' Before:
screenW = GetSystemMetrics(SM_CXSCREEN)  ' pixels
screenH = GetSystemMetrics(SM_CYSCREEN)  ' pixels

' After (Excel - workspace size in points, excludes taskbar):
workW = Application.UsableWidth    ' points
workH = Application.UsableHeight   ' points
' Note: 1 point = 1/72 inch. NOT pixels.

' CAUTION: Application.Width/Height is the Excel
' WINDOW size, not the screen size.
'@
        Note = 'Application.UsableWidth/Height = workspace in points (Excel only). For pixels, no pure VBA equivalent. In Access: Screen.Width/Height (twips).'
    }
    'ShowWindow' = @{
        Lib = 'user32'
        Alt = 'Application.Visible, WindowState, or UserForm.Show/Hide'
        Example = @'
' Before:
ShowWindow hWnd, SW_SHOW
ShowWindow hWnd, SW_MINIMIZE
ShowWindow hWnd, SW_MAXIMIZE

' After:
Application.Visible = True          ' show
Application.WindowState = xlMinimized  ' minimize
Application.WindowState = xlMaximized  ' maximize
Application.WindowState = xlNormal     ' restore

' For UserForm:
frm.Show / frm.Hide
'@
        Note = ''
    }
    'SetForegroundWindow' = @{
        Lib = 'user32'
        Alt = 'AppActivate (VBA built-in)'
        Example = @'
' Before:
SetForegroundWindow hWnd

' After:
On Error Resume Next  ' raises error 5 if not found
AppActivate "Window Title"
On Error GoTo 0
'@
        Note = 'AppActivate does partial title matching - may activate the wrong window if titles are similar. Always wrap in error handling.'
    }
    'SendMessage' = @{
        Lib = 'user32'
        Alt = 'Depends on message type. Often no direct alternative.'
        Example = ''
        Note = 'SendMessage is highly versatile. Review each call site individually. Common uses: scrolling listboxes, setting control properties. If used for external app automation, the business process itself may need redesign.'
    }
    'PostMessage' = @{
        Lib = 'user32'
        Alt = '(Same as SendMessage - review individually)'
        Example = ''
        Note = ''
    }

    # --- File ---
    'SHFileOperation' = @{
        Lib = 'shell32'
        Alt = 'FileCopy / Kill / Name / MkDir / RmDir or FileSystemObject'
        Example = @'
' Before:
SHFileOperation fileOp  ' copy/move/delete with recycle bin

' After:
' Copy file:   FileCopy src, dst
' Move/rename: Name src As dst
' Delete:      Kill path  ' permanent, no recycle bin
' Create dir:  MkDir path
' Remove dir:  RmDir path  ' must be empty

' Or use FileSystemObject for folders:
' fso.CopyFolder / fso.DeleteFolder / fso.MoveFolder
' Note: fso.DeleteFile is also permanent (no recycle bin)

' Delete (recycle bin): no pure VBA equivalent
'@
        Note = 'Kill does not support wildcards in the path portion (only filename). Recycle bin delete has no VBA equivalent. FileSystemObject may also be restricted by EDR.'
    }
    'ShellExecute' = @{
        Lib = 'shell32'
        Alt = 'ThisWorkbook.FollowHyperlink (documents) or Shell (executables only)'
        Example = @'
' Before:
ShellExecute 0, "open", path, vbNullString, vbNullString, SW_SHOW

' After (open document/URL with default app - Excel only):
ThisWorkbook.FollowHyperlink path
' Note: may trigger security warnings

' After (run executable only - not documents):
Shell "notepad.exe C:\file.txt", vbNormalFocus
' CAUTION: Shell cannot open .pdf, .xlsx etc.
' by file association. Use FollowHyperlink instead.
'@
        Note = 'Shell only launches executables, not documents by association. FollowHyperlink is Excel-specific and may trigger security prompts.'
    }

    # --- Clipboard ---
    'OpenClipboard' = @{
        Lib = 'user32'
        Alt = 'MSForms.DataObject (text only)'
        Example = @'
' Before:
OpenClipboard 0
hData = GetClipboardData(CF_TEXT)
' ...
CloseClipboard

' After (text only):
Dim d As New MSForms.DataObject
d.GetFromClipboard
text = d.GetText
'@
        Note = 'MSForms.DataObject handles text only (no images/files). Requires Microsoft Forms 2.0 reference. Add error handling for clipboard lock failures.'
    }
    'GetClipboardData' = @{
        Lib = 'user32'
        Alt = '(See OpenClipboard - text only via MSForms.DataObject)'
        Example = ''
        Note = ''
    }
    'SetClipboardData' = @{
        Lib = 'user32'
        Alt = 'MSForms.DataObject.SetText / PutInClipboard (text only)'
        Example = ''
        Note = ''
    }
    'CloseClipboard' = @{
        Lib = 'user32'
        Alt = '(Remove together with OpenClipboard)'
        Example = ''
        Note = ''
    }

    # --- EDR pattern-level entries (keyed by detection pattern name) ---
    'COM / CreateObject' = @{
        Lib = '(COM)'
        Alt = 'Review each ProgID. FSO -> Dir$/FileCopy/MkDir. ADODB -> keep if needed.'
        Example = @'
' Before:
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")

' After (early binding - requires reference):
Dim fso As New Scripting.FileSystemObject

' After (use VBA built-in instead of FSO):
' Dir(), Open, Kill, FileCopy, MkDir, RmDir
' (preferred in EDR environments - no COM needed)
'@
        Note = 'CreateObject itself is usually not blocked by EDR. Review if the created object is problematic.'
    }
    'COM / GetObject' = @{
        Lib = '(COM)'
        Alt = 'Replace WMI queries with Application object properties where possible'
        Example = @'
' Before:
Set wmi = GetObject("winmgmts:\\.\root\cimv2")
Set xlApp = GetObject(, "Excel.Application")

' After (attach to running Excel):
' Use Application object directly (already in scope)
Dim wb As Workbook
Set wb = Workbooks("Book1.xlsx")
'@
        Note = 'GetObject("winmgmts:") for process info can often be eliminated.'
    }
    'Shell / process' = @{
        Lib = '(VBA)'
        Alt = 'ThisWorkbook.FollowHyperlink (documents) or redesign workflow'
        Example = @'
' Before:
Shell "notepad.exe C:\log.txt", vbNormalFocus
Shell "cmd /c del C:\temp\*.tmp"

' After (open document):
ThisWorkbook.FollowHyperlink "C:\log.txt"

' After (delete files with VBA):
Kill "C:\temp\*.tmp"

' After (run macro in another workbook):
Application.Run "'Other.xlsm'!MacroName"
'@
        Note = 'Shell and cmd are primary EDR targets. Avoid launching external processes.'
    }
    'File I/O' = @{
        Lib = '(VBA)'
        Alt = 'VBA standard file I/O is generally safe. No change needed.'
        Example = @'
' Before (binary read via API):
' CopyMemory / ReadFile API calls

' After (VBA native binary read):
Open path For Binary Access Read As #1
Dim buf() As Byte: ReDim buf(LOF(1) - 1)
Get #1, , buf: Close #1
'@
        Note = 'Detected for awareness. Open/Kill/FileCopy are not typically blocked by EDR.'
    }
    'FileSystemObject' = @{
        Lib = '(COM)'
        Alt = 'Dir$, FileCopy, MkDir, RmDir, Kill (VBA built-in)'
        Example = @'
' Before:
Dim fso As Object
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(path) Then fso.CopyFile src, dst

' After (VBA built-in):
If Dir(path) <> "" Then FileCopy src, dst
'@
        Note = 'FSO is more capable but VBA built-ins cover most use cases.'
    }
    'Registry' = @{
        Lib = '(VBA)'
        Alt = 'Store settings in Excel sheet, JSON file, or CustomDocumentProperties'
        Example = @'
' Before:
SaveSetting "MyApp", "Config", "LastRun", Now()
val = GetSetting("MyApp", "Config", "LastRun", "")

' After (INI-style config file):
Dim configPath As String
configPath = ThisWorkbook.Path & "\config.ini"
Open configPath For Output As #1
Print #1, "LastRun=" & Now(): Close #1
'@
        Note = 'GetSetting/SaveSetting use VBA-specific registry area. Consider file-based config.'
    }
    'SendKeys' = @{
        Lib = '(VBA)'
        Alt = 'No direct alternative. Redesign to avoid UI automation.'
        Example = @'
' Before:
SendKeys "^c"       ' Ctrl+C
SendKeys "{ENTER}"  ' press Enter

' After (copy via object model):
Selection.Copy
' After (activate/close):
ActiveWorkbook.Close SaveChanges:=True
'@
        Note = 'SendKeys is fragile and blocked by many EDR policies. Rethink the workflow.'
    }
    'Network / HTTP' = @{
        Lib = '(COM)'
        Alt = 'No alternative if network access is required. Verify EDR policy.'
        Example = @'
' Before:
Dim http As Object
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", url, False: http.Send
data = http.responseText

' After (Power Query - preferred, no VBA needed):
' Data tab > Get Data > From Web > enter URL
'@
        Note = 'XMLHTTP/WinHttp may be allowed. Test in target environment.'
    }
    'PowerShell / WScript' = @{
        Lib = '(Shell)'
        Alt = 'No alternative. Remove script execution or redesign.'
        Example = @'
' Before:
Shell "powershell -Command ""Get-Process"""
Set wsh = CreateObject("WScript.Shell")
wsh.Run "cscript //nologo script.vbs"

' After:
' Move .vbs logic into a VBA module and call it directly
Call MyConvertedSubroutine
'@
        Note = 'Script hosts are the highest-risk EDR target. Almost always blocked.'
    }
    'Process / WMI' = @{
        Lib = '(COM)'
        Alt = 'Use Application object properties instead of WMI queries'
        Example = @'
' Before:
Set wmi = GetObject("winmgmts:\\.\root\cimv2")
Set os = wmi.ExecQuery("SELECT Caption FROM Win32_OperatingSystem")

' After (system info via Environ$):
userName = Environ$("USERNAME")
compName = Environ$("COMPUTERNAME")
osInfo = Environ$("OS")
'@
        Note = 'WMI process enumeration is rarely needed in Excel VBA. Remove if possible.'
    }
    'DLL loading' = @{
        Lib = '(API)'
        Alt = 'Remove dynamic DLL loading. Use static Declare instead.'
        Example = @'
' Before (dynamic loading):
Dim hLib As LongPtr
hLib = LoadLibrary("custom.dll")
Dim pFunc As LongPtr
pFunc = GetProcAddress(hLib, "MyFunc")
FreeLibrary hLib

' After (static declaration):
Declare PtrSafe Function MyFunc Lib "custom.dll" () As Long
'@
        Note = 'LoadLibrary/GetProcAddress pattern is blocked by EDR.'
    }
    'Clipboard' = @{
        Lib = '(COM/API)'
        Alt = 'MSForms.DataObject for text clipboard (Forms 2.0 reference required)'
        Example = @'
' Before (API-based clipboard):
OpenClipboard 0
hData = GetClipboardData(CF_TEXT)
CloseClipboard

' After (text via MSForms.DataObject):
Dim d As New MSForms.DataObject
d.GetFromClipboard
txt = d.GetText
'@
        Note = 'Text only. No image/file clipboard support.'
    }
    'Environment' = @{
        Lib = '(VBA)'
        Alt = 'No change needed. Environ$ is VBA standard.'
        Example = @'
' Environ$() itself IS the recommended alternative for API calls.
' Common safe Environ$ values:
userName = Environ$("USERNAME")
compName = Environ$("COMPUTERNAME")
tempDir  = Environ$("TEMP")
'@
        Note = 'Detected for awareness. Not typically blocked by EDR.'
    }
    'Auto-execution' = @{
        Lib = '(VBA)'
        Alt = 'No change needed. Configure macro security settings.'
        Example = @'
' Before (heavy Auto_Open):
Sub Auto_Open()
    ConnectToDatabase
    DownloadUpdates
    Shell "powershell ..."
End Sub

' After (minimal Auto_Open + user-triggered):
Sub Auto_Open()
    MsgBox "Click [Initialize] on the ribbon to start.", vbInformation
End Sub
'@
        Note = 'Auto_Open/Workbook_Open are standard VBA entry points.'
    }
    'Encoding / obfuscation' = @{
        Lib = '(VBA)'
        Alt = 'Replace Chr$() with string literals'
        Example = @'
' Before (obfuscated):
path = Chr$(67) & Chr$(58) & Chr$(92)  ' "C:\"
cmd = Chr(112) & Chr(111) & Chr(119) & Chr(101) & Chr(114) & Chr(115) & Chr(104) & Chr(101) & Chr(108) & Chr(108)

' After (plain text):
path = "C:\"
cmd = "powershell"
'@
        Note = 'Heavy Chr$() usage looks suspicious to EDR. Use "..." instead.'
    }

    # --- Compatibility / Legacy pattern-level entries ---
    'Declare without PtrSafe' = @{
        Lib = '(any)'
        Alt = 'Add PtrSafe keyword + review parameter types for LongPtr'
        Example = @'
' Before:
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long) As Long

' After:
Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
'@
        Note = '64-bit Office requires PtrSafe on all Declare statements. Handle/pointer params must be LongPtr.'
    }
    '64-bit: Missing PtrSafe' = @{
        Lib = '(any)'
        Alt = 'Add PtrSafe keyword + review parameter types for LongPtr'
        Example = @'
' Before:
Declare Function GetWindow Lib "user32" (ByVal hWnd As Long) As Long

' After:
Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
'@
        Note = '64-bit Office requires PtrSafe on all Declare statements.'
    }
    'Long for handles' = @{
        Lib = '(any)'
        Alt = 'Change As Long to As LongPtr for handles, pointers, and HWNDs in PtrSafe Declare'
        Example = @'
' Before:
Declare PtrSafe Function FindWindow Lib "user32" _
    Alias "FindWindowA" (ByVal lpClass As String, ByVal lpWindow As String) As Long

' After:
Declare PtrSafe Function FindWindow Lib "user32" _
    Alias "FindWindowA" (ByVal lpClass As String, ByVal lpWindow As String) As LongPtr
'@
        Note = 'On 64-bit Office, handles and pointers are 8 bytes. Using Long (4 bytes) truncates them, causing crashes or silent corruption.'
    }
    '64-bit: Long for handles' = @{
        Lib = '(any)'
        Alt = 'Change handle/pointer parameters from Long to LongPtr'
        Example = @'
' Before:
Declare PtrSafe Function FindWindow Lib "user32" _
    Alias "FindWindowA" (ByVal lpClass As String, ByVal lpWindow As String) As Long

' After:
Declare PtrSafe Function FindWindow Lib "user32" _
    Alias "FindWindowA" (ByVal lpClass As String, ByVal lpWindow As String) As LongPtr
'@
        Note = 'HWND, HINSTANCE, pointers are 8 bytes on 64-bit. Long is always 4 bytes.'
    }
    'VarPtr' = @{
        Lib = '(VBA)'
        Alt = 'Assign VarPtr/ObjPtr/StrPtr results to LongPtr variables on 64-bit'
        Example = @'
' Before:
Dim addr As Long
addr = VarPtr(myVar)

' After (64-bit safe):
Dim addr As LongPtr
addr = VarPtr(myVar)
'@
        Note = 'VarPtr/ObjPtr/StrPtr return LongPtr on 64-bit VBA7. Storing in Long truncates the address.'
    }
    '64-bit: VarPtr/ObjPtr/StrPtr' = @{
        Lib = '(VBA)'
        Alt = 'Store result in LongPtr variable'
        Example = @'
' Before:
Dim addr As Long
addr = VarPtr(myVar)

' After:
Dim addr As LongPtr
addr = VarPtr(myVar)
'@
        Note = 'These functions return LongPtr on 64-bit.'
    }
    'DDEInitiate' = @{
        Lib = '(DDE)'
        Alt = 'COM Automation or Application.Run for inter-app communication'
        Example = @'
' Before:
ch = DDEInitiate("Excel", "Sheet1")
DDEExecute ch, "[OPEN(""file.xls"")]"
DDETerminate ch

' After (COM Automation):
Dim xlApp As Object
Set xlApp = GetObject(, "Excel.Application")
xlApp.Workbooks.Open "file.xls"
'@
        Note = 'DDE is deprecated and disabled by default in modern Office. Use COM Automation instead.'
    }
    'Deprecated: DDE' = @{
        Lib = '(DDE)'
        Alt = 'COM Automation or Application.Run for inter-app communication'
        Example = ''
        Note = 'DDE is deprecated and disabled by default in modern Office.'
    }
    'InternetExplorer.Application' = @{
        Lib = '(COM)'
        Alt = 'MSXML2.XMLHTTP for HTTP requests, or Selenium/Edge WebDriver for browser automation'
        Example = @'
' Before:
Dim ie As Object
Set ie = CreateObject("InternetExplorer.Application")
ie.Navigate "https://example.com"

' After (HTTP request only):
Dim http As Object
Set http = CreateObject("MSXML2.XMLHTTP")
http.Open "GET", "https://example.com", False
http.Send
Dim html As String: html = http.responseText
'@
        Note = 'Internet Explorer has been removed from Windows.'
    }
    'Deprecated: IE Automation' = @{
        Lib = '(COM)'
        Alt = 'MSXML2.XMLHTTP for HTTP requests, or Edge WebDriver for browser automation'
        Example = ''
        Note = 'Internet Explorer has been removed from Windows.'
    }
    'MSComctlLib' = @{
        Lib = '(OCX)'
        Alt = 'Use ListView/TreeView from MSComctlLib if available, or replace with ListBox/UserForm controls'
        Example = ''
        Note = 'Microsoft released 64-bit MSCOMCTL.OCX (KB2687441, updated). If available, re-register it.'
    }
    'MSComDlg.CommonDialog' = @{
        Lib = '(OCX)'
        Alt = 'Application.GetOpenFilename / GetSaveAsFilename (Excel) or Application.FileDialog'
        Example = ''
        Note = 'MSComDlg.CommonDialog (COMDLG32.OCX) has no 64-bit version. Use Application.FileDialog.'
    }
    'MSCAL.Calendar' = @{
        Lib = '(OCX)'
        Alt = 'MonthView control (MS Date and Time Picker) or custom UserForm'
        Example = ''
        Note = 'MSCAL.Calendar (mscal.ocx) has no 64-bit version and is removed from Office 2010+.'
    }
    'Deprecated: Legacy Controls' = @{
        Lib = '(OCX)'
        Alt = 'Use Application.FileDialog, ListBox, or updated 64-bit OCX'
        Example = ''
        Note = 'Legacy ActiveX controls (MSCAL, MSComDlg, MSComctlLib) may not have 64-bit versions.'
    }
    'DAO.Database' = @{
        Lib = '(DAO)'
        Alt = 'ADO (ADODB.Connection / ADODB.Recordset) or CurrentDb in Access'
        Example = @'
' Before:
Dim db As DAO.Database
Set db = DBEngine.OpenDatabase("C:\data.mdb")

' After (ADO):
Dim cn As Object
Set cn = CreateObject("ADODB.Connection")
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\data.mdb"
'@
        Note = 'DAO is still supported in Access (via CurrentDb) but standalone DAO references should migrate to ADO.'
    }
    'Deprecated: DAO' = @{
        Lib = '(DAO)'
        Alt = 'ADO (ADODB.Connection / ADODB.Recordset) or CurrentDb in Access'
        Example = ''
        Note = 'Standalone DAO references should migrate to ADO for better compatibility.'
    }
    'DefLng' = @{
        Lib = '(VBA)'
        Alt = 'Explicit variable declarations with Dim ... As Long'
        Example = @'
' Before:
DefLng A-Z

' After:
' Remove DefLng and add explicit type declarations:
Dim count As Long
Dim index As Long
'@
        Note = 'DefType statements make code harder to read and maintain. Use Option Explicit.'
    }
    'Legacy: DefType' = @{
        Lib = '(VBA)'
        Alt = 'Explicit variable declarations with Dim ... As Type'
        Example = ''
        Note = 'DefType statements make code harder to read. Use Option Explicit and declare each variable.'
    }
    'GoSub' = @{
        Lib = '(VBA)'
        Alt = 'Refactor into separate Sub or Function procedures'
        Example = @'
' Before:
Sub Main()
    GoSub DoWork
    Exit Sub
DoWork:
    MsgBox "Working"
    Return
End Sub

' After:
Sub Main()
    DoWork
End Sub
Private Sub DoWork()
    MsgBox "Working"
End Sub
'@
        Note = 'GoSub/Return is a legacy construct from BASIC. Refactor into standalone Sub/Function.'
    }
    'Legacy: GoSub' = @{
        Lib = '(VBA)'
        Alt = 'Refactor into separate Sub or Function procedures'
        Example = ''
        Note = 'GoSub/Return is a legacy construct. Refactor into standalone Sub/Function.'
    }
    'While/Wend' = @{
        Lib = '(VBA)'
        Alt = 'Do While ... Loop (supports Exit Do)'
        Example = @'
' Before:
While condition
    DoSomething
Wend

' After:
Do While condition
    DoSomething
    If needExit Then Exit Do
Loop
'@
        Note = 'While...Wend cannot be exited early. Do While...Loop supports Exit Do.'
    }
    'Legacy: While/Wend' = @{
        Lib = '(VBA)'
        Alt = 'Do While ... Loop (supports Exit Do)'
        Example = ''
        Note = 'While...Wend cannot be exited early. Do While...Loop supports Exit Do.'
    }

    # --- Environment dependency entries ---
    'Fixed drive letter' = @{
        Lib = '(Path)'
        Alt = 'Use Environ$("TEMP") for temp files, or externalize paths to config'
        Example = ''
        Note = 'Hardcoded drive letters break on migration. Environ$("TEMP") is stable across all scenarios (local/OneDrive/SharePoint). Do NOT use ThisWorkbook.Path as replacement — it returns URL on cloud.'
    }
    'UNC path' = @{ Lib = '(Path)'; Alt = 'Externalize to config file or CustomDocumentProperties'; Example = ''; Note = 'UNC paths change on server migration. SharePoint migration replaces UNC shares entirely.' }
    'User folder' = @{ Lib = '(Path)'; Alt = 'Use Environ$("USERPROFILE") for local user folder'; Example = ''; Note = 'Environ$("USERPROFILE") always returns local path (stable). But if the code then opens files relative to this path, verify those files exist locally and not only on cloud.' }
    'Desktop / Documents' = @{ Lib = '(Path)'; Alt = 'Use Environ$("USERPROFILE") & "\Desktop" or "\Documents"'; Example = ''; Note = 'Folder names differ between EN/JP locales. WScript.Shell.SpecialFolders is blocked by EDR. Environ$-based approach is safe.' }
    'AppData' = @{ Lib = '(Path)'; Alt = 'Use Environ$("APPDATA") or Environ$("LOCALAPPDATA")'; Example = ''; Note = 'Environ$ returns correct local path. Hardcoded AppData paths break across users.' }
    'Program Files' = @{ Lib = '(Path)'; Alt = 'Use Environ$("ProgramFiles")'; Example = ''; Note = 'Path differs between 32-bit and 64-bit Office. Environ$ handles this automatically.' }
    'Fixed printer name' = @{ Lib = '(Printer)'; Alt = 'Read Application.ActivePrinter dynamically, or externalize to config'; Example = ''; Note = 'Target environment may have completely different printers.' }
    'Fixed IP address' = @{ Lib = '(Network)'; Alt = 'Externalize to config file'; Example = ''; Note = 'IP addresses change on environment migration.' }
    'Fixed connection host' = @{ Lib = '(Network)'; Alt = 'Externalize Server=/Host= values to config'; Example = ''; Note = 'Database/service hosts change on migration.' }
    'localhost' = @{ Lib = '(Network)'; Alt = 'Confirm localhost backend is deployed (casedesk architecture uses localhost HTTP for VBA-PS communication)'; Example = ''; Note = 'Probe confirmed: XMLHTTP/WinHttp to localhost works. HttpListener can bind. localhost is expected for local backend communication.' }
    'Connection string' = @{ Lib = '(DB)'; Alt = 'Externalize to config. Check Provider version (ACE vs Jet)'; Example = ''; Note = 'ACE/Jet Provider version differs on 64-bit. Connection string paths may also need updating for cloud.' }
    'External workbook open (literal)' = @{ Lib = '(File)'; Alt = 'Externalize file paths to config. Do NOT rely on ThisWorkbook.Path — it returns URL on cloud.'; Example = ''; Note = 'Tested: creating adjacent files via ThisWorkbook.Path fails on OneDrive/SharePoint. Hardcoded paths also break. Use config or TEMP-based approach.' }

    # --- New env Risk/Review entries ---
    'Dir() path check' = @{
        Lib = '(Path)'
        Alt = 'Use FSO.FileExists or error handling instead of Dir()'
        Example = @'
' Before:
If Dir(ThisWorkbook.Path & "\data.csv") <> "" Then ...

' After (error handling):
On Error Resume Next
Open filePath For Input As #1
If Err.Number = 0 Then
    Close #1
    ' file exists
End If
On Error GoTo 0
'@
        Note = 'Dir() fails when path is a URL (cloud/SharePoint). Use FileSystemObject or error handling for cloud-safe existence checks.'
    }
    'Path concatenation' = @{
        Lib = '(Path)'
        Alt = 'Use Environ$("TEMP") or config-based path instead of Path & "\"'
        Example = @'
' Before:
savePath = ThisWorkbook.Path & "\output.xlsx"

' After (TEMP-based, cloud-safe):
savePath = Environ$("TEMP") & "\output.xlsx"

' After (config-based):
savePath = Range("ConfigSheet!B2").Value & "\output.xlsx"
'@
        Note = 'Path & "\" fails when Path is a URL (SharePoint/OneDrive). Use Environ$("TEMP") or externalize to config.'
    }
    'SaveAs call' = @{
        Lib = '(File)'
        Alt = 'Verify SaveAs destination is not URL-dependent'
        Example = @'
' Before:
ActiveWorkbook.SaveAs ThisWorkbook.Path & "\new.xlsx"

' After (TEMP-based, stable):
ActiveWorkbook.SaveAs Environ$("TEMP") & "\new.xlsx"
'@
        Note = 'SaveAs works but destination changes meaning on cloud. TEMP-based saving is stable.'
    }
    'External workbook ref' = @{
        Lib = '(File)'
        Alt = 'Review external workbook references for path changes'
        Example = ''
        Note = 'External workbook references in formulas break when paths become URLs on cloud.'
    }
    'BeforeSave event' = @{
        Lib = '(Event)'
        Alt = 'Review save-event logic for AutoSave compatibility'
        Example = @'
' Cloud: AutoSave=True fires BeforeSave frequently.
' Review if BeforeSave contains heavy operations or
' user prompts that assume manual save only.
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    ' Check AutoSave state:
    If Me.AutoSaveOn Then Exit Sub  ' skip for AutoSave
    ' ... original logic ...
End Sub
'@
        Note = 'AutoSave=True on cloud. BeforeSave fires frequently. Review save-event-dependent code.'
    }
    'AfterSave event' = @{
        Lib = '(Event)'
        Alt = 'Review save-event logic for AutoSave compatibility'
        Example = ''
        Note = 'AutoSave=True on cloud. AfterSave fires frequently. Review save-event-dependent code.'
    }
    'LinkSources / UpdateLink' = @{
        Lib = '(Link)'
        Alt = 'Review external link handling for cloud path changes'
        Example = ''
        Note = 'External links may resolve to URLs on cloud. Review LinkSources/UpdateLink usage.'
    }
    'Workbooks.Open (variable)' = @{
        Lib = '(File)'
        Alt = 'Verify the variable contains a valid path for the target environment'
        Example = @'
' Before:
Workbooks.Open someVariable

' After (validate path):
If Left(someVariable, 4) = "http" Then
    ' Handle URL path
Else
    Workbooks.Open someVariable
End If
'@
        Note = 'Workbooks.Open works but paths become URL on cloud. Check if input is fixed path or variable.'
    }
    'CurDir' = @{
        Lib = '(Path)'
        Alt = 'Use Environ$("TEMP") or explicit full paths instead'
        Example = ''
        Note = 'Tested: CurDir returns different values across scenarios (Local=depth4, SharePoint=depth2). Do NOT use ThisWorkbook.Path as replacement — it returns URL on cloud. Use Environ$("TEMP") or config-based paths.'
    }
    'ChDir' = @{
        Lib = '(Path)'
        Alt = 'Avoid ChDir; use full paths instead'
        Example = ''
        Note = 'ChDir may fail on URL paths (cloud condition). Use explicit full paths.'
    }

    # --- Info items (no highlight, report only) ---
    'ThisWorkbook.Path' = @{ Lib = '(Info)'; Alt = 'Reference only — safe alone, but dangerous when used with Dir/Open/SaveAs'; Example = ''; Note = 'Tested: returns URL (https://...) on both OneDrive synced AND SharePoint Open-in-App. Even from sync folder, NOT the local sync path. Safe to read, but do NOT pass to Dir(), Workbooks.Open, or SaveAs.' }
    'ActiveWorkbook.Path' = @{ Lib = '(Info)'; Alt = 'Same concern as ThisWorkbook.Path'; Example = ''; Note = 'Returns URL on cloud. Same behavior as ThisWorkbook.Path.' }
    'ThisWorkbook.FullName' = @{ Lib = '(Info)'; Alt = 'Reference only — may return URL on cloud'; Example = ''; Note = 'Tested: returns full URL on cloud (depth=10). Code that parses FullName (e.g. splitting by "\") will break because URL uses "/" as separator.' }
    'ActiveWorkbook.FullName' = @{ Lib = '(Info)'; Alt = 'Same concern as ThisWorkbook.FullName'; Example = ''; Note = 'Returns full URL on cloud. Same behavior.' }

    # --- Business dependency entries ---
    'Outlook integration' = @{ Lib = '(Office)'; Alt = 'Usually works, but test CreateObject in target environment'; Example = ''; Note = 'Check Outlook version compatibility.' }
    'Word integration' = @{ Lib = '(Office)'; Alt = 'Usually works, but test CreateObject in target environment'; Example = ''; Note = 'Check Word version compatibility.' }
    'Access / DB integration' = @{ Lib = '(DB)'; Alt = 'Consider DAO to ADO migration. Check ACE Provider'; Example = ''; Note = '64-bit environment may have different Provider.' }
    'PDF export' = @{ Lib = '(Print)'; Alt = 'ExportAsFixedFormat usually works'; Example = ''; Note = 'May depend on printer driver in some cases.' }
    'Print' = @{ Lib = '(Print)'; Alt = 'Check for hardcoded printer names'; Example = ''; Note = 'Verify printer configuration in target environment.' }
    'External EXE' = @{ Lib = '(Process)'; Alt = 'Likely blocked by EDR'; Example = ''; Note = 'Same risk level as Shell. Avoid launching external processes.' }
}

function Get-VbaApiReplacements {
    return $script:VbaApiReplacements
}

Export-ModuleMember -Function Read-Ole2, Read-Ole2Stream, Write-Ole2Stream,
    Decompress-VBA, Compress-VBA,
    Get-VbaProjectBytes, Save-VbaProjectBytes,
    Get-VbaModuleList, Get-VbaModuleCode,
    New-HtmlBase, New-HtmlCodeView,
    Resolve-VbaFilePath, New-VbaOutputDir,
    Write-VbaLog, Write-VbaStatus, Write-VbaResult, Write-VbaError, Write-VbaHeader,
    Get-VbaCodepage, Get-AllModuleCode, Get-VbaAnalysis, Get-VbaApiReplacements
