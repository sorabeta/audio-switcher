[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$assetsDir = Join-Path $root "assets"
if (-not (Test-Path -LiteralPath $assetsDir)) {
    New-Item -ItemType Directory -Path $assetsDir | Out-Null
}

function New-Canvas {
    param([int]$Size)
    $bmp = New-Object System.Drawing.Bitmap $Size, $Size
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    return @{ Bitmap = $bmp; Graphics = $g }
}

function Save-Png {
    param(
        [System.Drawing.Bitmap]$Bitmap,
        [string]$Path
    )
    $Bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
}

function Write-IcoFromPng {
    param(
        [byte[]]$PngBytes,
        [int]$Size,
        [string]$OutputPath
    )

    $fs = [System.IO.File]::Open($OutputPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
    try {
        $bw = New-Object System.IO.BinaryWriter($fs)
        $bw.Write([UInt16]0)
        $bw.Write([UInt16]1)
        $bw.Write([UInt16]1)
        $bw.Write([byte]0)
        $bw.Write([byte]0)
        $bw.Write([byte]0)
        $bw.Write([byte]0)
        $bw.Write([UInt16]1)
        $bw.Write([UInt16]32)
        $bw.Write([UInt32]$PngBytes.Length)
        $bw.Write([UInt32]22)
        $bw.Write($PngBytes)
        $bw.Flush()
    }
    finally {
        $fs.Dispose()
    }
}

function Get-PngBytes {
    param([System.Drawing.Bitmap]$Bitmap)
    $ms = New-Object System.IO.MemoryStream
    try {
        $Bitmap.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
        return $ms.ToArray()
    }
    finally {
        $ms.Dispose()
    }
}

function Draw-Logo {
    param(
        [System.Drawing.Graphics]$Graphics,
        [int]$Size
    )

    $Graphics.Clear([System.Drawing.Color]::Transparent)

    $rect = New-Object System.Drawing.Rectangle 0, 0, $Size, $Size
    $bg1 = [System.Drawing.Color]::FromArgb(10, 61, 98)
    $bg2 = [System.Drawing.Color]::FromArgb(24, 126, 155)
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        $rect,
        $bg1,
        $bg2,
        45
    )
    $Graphics.FillEllipse($brush, [int]($Size * 0.04), [int]($Size * 0.04), [int]($Size * 0.92), [int]($Size * 0.92))
    $brush.Dispose()

    $ringPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(235, 248, 252), [Math]::Max(4, [int]($Size * 0.055)))
    $ringPen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
    $Graphics.DrawArc(
        $ringPen,
        [int]($Size * 0.19),
        [int]($Size * 0.19),
        [int]($Size * 0.62),
        [int]($Size * 0.62),
        28,
        286
    )
    $ringPen.Dispose()

    $arrowPen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 249, 212), [Math]::Max(4, [int]($Size * 0.07)))
    $arrowPen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $arrowPen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $Graphics.DrawLine(
        $arrowPen,
        [int]($Size * 0.70),
        [int]($Size * 0.18),
        [int]($Size * 0.82),
        [int]($Size * 0.18)
    )
    $Graphics.DrawLine(
        $arrowPen,
        [int]($Size * 0.82),
        [int]($Size * 0.18),
        [int]($Size * 0.82),
        [int]($Size * 0.30)
    )
    $Graphics.DrawLine(
        $arrowPen,
        [int]($Size * 0.82),
        [int]($Size * 0.18),
        [int]($Size * 0.73),
        [int]($Size * 0.09)
    )
    $arrowPen.Dispose()

    $speakerBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(245, 250, 255))
    $speakerPoints = [System.Drawing.Point[]]@(
        (New-Object System.Drawing.Point([int]($Size * 0.26), [int]($Size * 0.55))),
        (New-Object System.Drawing.Point([int]($Size * 0.40), [int]($Size * 0.55))),
        (New-Object System.Drawing.Point([int]($Size * 0.53), [int]($Size * 0.43))),
        (New-Object System.Drawing.Point([int]($Size * 0.53), [int]($Size * 0.73))),
        (New-Object System.Drawing.Point([int]($Size * 0.40), [int]($Size * 0.61))),
        (New-Object System.Drawing.Point([int]($Size * 0.26), [int]($Size * 0.61)))
    )
    $Graphics.FillPolygon($speakerBrush, $speakerPoints)
    $speakerBrush.Dispose()

    $wavePen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 249, 212), [Math]::Max(3, [int]($Size * 0.045)))
    $wavePen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $wavePen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $Graphics.DrawArc($wavePen, [int]($Size * 0.50), [int]($Size * 0.44), [int]($Size * 0.16), [int]($Size * 0.28), -45, 90)
    $Graphics.DrawArc($wavePen, [int]($Size * 0.56), [int]($Size * 0.36), [int]($Size * 0.24), [int]($Size * 0.44), -45, 90)
    $wavePen.Dispose()
}

$preview = New-Canvas -Size 512
try {
    Draw-Logo -Graphics $preview.Graphics -Size 512
    Save-Png -Bitmap $preview.Bitmap -Path (Join-Path $assetsDir "app-logo.png")
}
finally {
    $preview.Graphics.Dispose()
    $preview.Bitmap.Dispose()
}

$iconCanvas = New-Canvas -Size 256
try {
    Draw-Logo -Graphics $iconCanvas.Graphics -Size 256
    $pngBytes = Get-PngBytes -Bitmap $iconCanvas.Bitmap
    Write-IcoFromPng -PngBytes $pngBytes -Size 256 -OutputPath (Join-Path $assetsDir "app-icon.ico")
}
finally {
    $iconCanvas.Graphics.Dispose()
    $iconCanvas.Bitmap.Dispose()
}

Write-Host "Generated:"
Write-Host (Join-Path $assetsDir "app-logo.png")
Write-Host (Join-Path $assetsDir "app-icon.ico")
