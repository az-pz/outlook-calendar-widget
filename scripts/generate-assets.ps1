<#
.SYNOPSIS
    Regenerates the app/tile/widget icon assets from vector primitives (no binary sources needed).
.NOTES
    Runs under Windows PowerShell 5.1 or PowerShell 7 on Windows (uses System.Drawing / GDI+).
    Widget picker screenshots are produced separately by scripts/widget-preview (npm install; npm run render).
#>
param(
    [string] $OutputDirectory = (Join-Path $PSScriptRoot '..\src\OutlookCalendarWidget\Assets')
)

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

$OutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null

$blueTop = [System.Drawing.Color]::FromArgb(255, 40, 137, 222)
$blueBottom = [System.Drawing.Color]::FromArgb(255, 15, 94, 176)

function New-RoundedRectPath([float] $x, [float] $y, [float] $w, [float] $h, [float] $r) {
    $path = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = [Math]::Max(1.0, $r * 2)
    $path.AddArc($x, $y, $d, $d, 180, 90)
    $path.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
    $path.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
    $path.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
    $path.CloseFigure()
    return $path
}

# Draws the calendar mark: a white page with a header band, binder rings and a 3x2 grid of day cells,
# with the "today" cell highlighted.
function Draw-CalendarMark([System.Drawing.Graphics] $g, [float] $cx, [float] $cy, [float] $size, [bool] $onBlue) {
    $w = $size
    $h = $size * 0.92
    $x = $cx - $w / 2
    $y = $cy - $h / 2 + $size * 0.04
    $radius = $size * 0.14

    $pageColor = if ($onBlue) { [System.Drawing.Color]::White } else { $blueBottom }
    $inkColor = if ($onBlue) { $blueBottom } else { [System.Drawing.Color]::White }

    $page = New-RoundedRectPath $x $y $w $h $radius
    $g.FillPath((New-Object System.Drawing.SolidBrush $pageColor), $page)

    $bandHeight = $h * 0.26
    $band = New-RoundedRectPath $x $y $w $bandHeight $radius
    $bandRect = New-Object System.Drawing.RectangleF $x, ($y + $bandHeight / 2), $w, ($bandHeight / 2)
    $bandColor = if ($onBlue) { [System.Drawing.Color]::FromArgb(255, 200, 226, 250) } else { [System.Drawing.Color]::FromArgb(255, 10, 70, 140) }
    $bandBrush = New-Object System.Drawing.SolidBrush $bandColor
    $g.FillPath($bandBrush, $band)
    $g.FillRectangle($bandBrush, $bandRect)

    $ringW = $size * 0.07
    $ringH = $size * 0.2
    foreach ($rx in @(($x + $w * 0.28), ($x + $w * 0.72))) {
        $ring = New-RoundedRectPath ($rx - $ringW / 2) ($y - $ringH * 0.45) $ringW $ringH ($ringW / 2)
        $g.FillPath((New-Object System.Drawing.SolidBrush $pageColor), $ring)
        $inner = New-RoundedRectPath ($rx - $ringW * 0.25) ($y - $ringH * 0.3) ($ringW * 0.5) ($ringH * 0.8) ($ringW * 0.25)
        $g.FillPath((New-Object System.Drawing.SolidBrush $bandColor), $inner)
    }

    $gridTop = $y + $bandHeight + $h * 0.1
    $cols = 3
    $rows = 2
    $gap = $w * 0.08
    $cellW = ($w - $gap * ($cols + 1)) / $cols
    $cellH = ($h - $bandHeight - $h * 0.1 - $gap * ($rows + 0.5)) / $rows
    for ($r = 0; $r -lt $rows; $r++) {
        for ($c = 0; $c -lt $cols; $c++) {
            $cellX = $x + $gap + $c * ($cellW + $gap)
            $cellY = $gridTop + $r * ($cellH + $gap)
            $isToday = ($r -eq 0 -and $c -eq 1)
            $color = if ($isToday) { $inkColor } else { [System.Drawing.Color]::FromArgb(70, $inkColor.R, $inkColor.G, $inkColor.B) }
            $cell = New-RoundedRectPath $cellX $cellY $cellW $cellH ($cellW * 0.22)
            $g.FillPath((New-Object System.Drawing.SolidBrush $color), $cell)
        }
    }
}

function New-Canvas([int] $width, [int] $height) {
    $bmp = New-Object System.Drawing.Bitmap $width, $height, ([System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.Clear([System.Drawing.Color]::Transparent)
    return @{ Bitmap = $bmp; Graphics = $g }
}

# App icon: blue rounded tile with the white calendar mark.
function Save-TileIcon([string] $name, [int] $size, [float] $markScale = 0.56, [bool] $plated = $true) {
    $canvas = New-Canvas $size $size
    $g = $canvas.Graphics
    if ($plated) {
        $inset = $size * 0.04
        $tile = New-RoundedRectPath $inset $inset ($size - 2 * $inset) ($size - 2 * $inset) ($size * 0.22)
        $rect = New-Object System.Drawing.RectangleF 0, 0, $size, $size
        $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush $rect, $blueTop, $blueBottom, 90.0
        $g.FillPath($brush, $tile)
        Draw-CalendarMark $g ($size / 2) ($size / 2) ($size * $markScale) $true
    }
    else {
        Draw-CalendarMark $g ($size / 2) ($size / 2) ($size * $markScale) $false
    }
    $canvas.Bitmap.Save((Join-Path $OutputDirectory $name), [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $canvas.Bitmap.Dispose()
}

# Wide/splash assets: transparent canvas with the plated icon centered.
function Save-CenteredIcon([string] $name, [int] $width, [int] $height, [float] $iconSize) {
    $canvas = New-Canvas $width $height
    $g = $canvas.Graphics
    $x = ($width - $iconSize) / 2
    $y = ($height - $iconSize) / 2
    $tile = New-RoundedRectPath $x $y $iconSize $iconSize ($iconSize * 0.22)
    $rect = New-Object System.Drawing.RectangleF $x, $y, $iconSize, $iconSize
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush $rect, $blueTop, $blueBottom, 90.0
    $g.FillPath($brush, $tile)
    Draw-CalendarMark $g ($width / 2) ($height / 2) ($iconSize * 0.56) $true
    $canvas.Bitmap.Save((Join-Path $OutputDirectory $name), [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $canvas.Bitmap.Dispose()
}

Save-TileIcon 'Square44x44Logo.scale-200.png' 88
Save-TileIcon 'Square44x44Logo.targetsize-24_altform-unplated.png' 24 0.9 $false
Save-TileIcon 'Square44x44Logo.targetsize-48_altform-unplated.png' 48 0.9 $false
Save-TileIcon 'Square44x44Logo.targetsize-256_altform-unplated.png' 256 0.9 $false
Save-TileIcon 'Square150x150Logo.scale-200.png' 300 0.5
Save-TileIcon 'LockScreenLogo.scale-200.png' 48
Save-TileIcon 'StoreLogo.png' 50
Save-TileIcon 'WidgetIcon.png' 64 0.62
Save-CenteredIcon 'Wide310x150Logo.scale-200.png' 620 300 220
Save-CenteredIcon 'SplashScreen.scale-200.png' 1240 600 320

# Multi-resolution .ico for the executable / window icon (PNG-compressed entries).
$icoSizes = @(16, 24, 32, 48, 64, 256)
$pngs = foreach ($s in $icoSizes) {
    $canvas = New-Canvas $s $s
    $inset = $s * 0.04
    $tile = New-RoundedRectPath $inset $inset ($s - 2 * $inset) ($s - 2 * $inset) ($s * 0.22)
    $rect = New-Object System.Drawing.RectangleF 0, 0, $s, $s
    $canvas.Graphics.FillPath((New-Object System.Drawing.Drawing2D.LinearGradientBrush $rect, $blueTop, $blueBottom, 90.0), $tile)
    Draw-CalendarMark $canvas.Graphics ($s / 2) ($s / 2) ($s * 0.58) $true
    $ms = New-Object System.IO.MemoryStream
    $canvas.Bitmap.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $canvas.Graphics.Dispose(); $canvas.Bitmap.Dispose()
    , $ms.ToArray()
}
$icoPath = Join-Path $OutputDirectory 'AppIcon.ico'
$fs = [System.IO.File]::Create($icoPath)
$bw = New-Object System.IO.BinaryWriter $fs
$bw.Write([UInt16]0); $bw.Write([UInt16]1); $bw.Write([UInt16]$icoSizes.Count)
$offset = 6 + 16 * $icoSizes.Count
for ($i = 0; $i -lt $icoSizes.Count; $i++) {
    $s = $icoSizes[$i]
    $dim = if ($s -ge 256) { 0 } else { $s }
    $bw.Write([byte]$dim); $bw.Write([byte]$dim); $bw.Write([byte]0); $bw.Write([byte]0)
    $bw.Write([UInt16]1); $bw.Write([UInt16]32)
    $bw.Write([UInt32]$pngs[$i].Length); $bw.Write([UInt32]$offset)
    $offset += $pngs[$i].Length
}
foreach ($png in $pngs) { $bw.Write($png) }
$bw.Dispose()

Write-Host "Assets written to $OutputDirectory"
