[CmdletBinding()]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Add-Type -AssemblyName System.Drawing

$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$docsDir = Join-Path $root "docs"
$logoPath = Join-Path $root "assets\app-logo.png"

if (-not (Test-Path -LiteralPath $docsDir)) {
    New-Item -ItemType Directory -Path $docsDir | Out-Null
}

if (-not (Test-Path -LiteralPath $logoPath)) {
    throw "Logo not found: $logoPath"
}

$logo = [System.Drawing.Image]::FromFile($logoPath)

function New-Canvas {
    param([int]$Width, [int]$Height)
    $bmp = New-Object System.Drawing.Bitmap $Width, $Height
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    return @{ Bitmap = $bmp; Graphics = $g }
}

function Save-Png {
    param([System.Drawing.Bitmap]$Bitmap, [string]$Path)
    $Bitmap.Save($Path, [System.Drawing.Imaging.ImageFormat]::Png)
}

function Draw-Card {
    param(
        [System.Drawing.Graphics]$Graphics,
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height,
        [System.Drawing.Color]$BackColor
    )
    $rect = New-Object System.Drawing.Rectangle $X, $Y, $Width, $Height
    $brush = New-Object System.Drawing.SolidBrush $BackColor
    $pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(225, 231, 237)), 1
    $Graphics.FillRectangle($brush, $rect)
    $Graphics.DrawRectangle($pen, $rect)
    $brush.Dispose()
    $pen.Dispose()
}

function Draw-TrailingText {
    param(
        [System.Drawing.Graphics]$Graphics,
        [string]$Text,
        [float]$FontSize,
        [int]$X,
        [int]$Y,
        [System.Drawing.Color]$Color,
        [switch]$Bold
    )
    $style = if ($Bold) { [System.Drawing.FontStyle]::Bold } else { [System.Drawing.FontStyle]::Regular }
    $font = New-Object System.Drawing.Font("Segoe UI", $FontSize, $style)
    $brush = New-Object System.Drawing.SolidBrush($Color)
    $Graphics.DrawString($Text, $font, $brush, [float]$X, [float]$Y)
    $brush.Dispose()
    $font.Dispose()
}

function New-BackgroundBrush {
    param([System.Drawing.Rectangle]$Rect)
    return New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        $Rect,
        [System.Drawing.Color]::FromArgb(243, 248, 250),
        [System.Drawing.Color]::FromArgb(226, 236, 241),
        90
    )
}

$tray = New-Canvas -Width 1200 -Height 760
try {
    $bgBrush = New-BackgroundBrush -Rect (New-Object System.Drawing.Rectangle 0, 0, 1200, 760)
    $tray.Graphics.FillRectangle($bgBrush, 0, 0, 1200, 760)
    $bgBrush.Dispose()

    Draw-TrailingText -Graphics $tray.Graphics -Text "Audio Switcher" -FontSize 34 -X 82 -Y 70 -Color ([System.Drawing.Color]::FromArgb(19, 45, 59)) -Bold
    Draw-TrailingText -Graphics $tray.Graphics -Text "Tray menu preview with branded icon and quick actions." -FontSize 14 -X 84 -Y 122 -Color ([System.Drawing.Color]::FromArgb(86, 104, 119))
    $tray.Graphics.DrawImage($logo, 82, 168, 128, 128)

    Draw-Card -Graphics $tray.Graphics -X 740 -Y 150 -Width 290 -Height 242 -BackColor ([System.Drawing.Color]::White)
    Draw-TrailingText -Graphics $tray.Graphics -Text "Switch To Next Output" -FontSize 13 -X 772 -Y 186 -Color ([System.Drawing.Color]::FromArgb(27, 36, 48))
    Draw-TrailingText -Graphics $tray.Graphics -Text "Settings" -FontSize 13 -X 772 -Y 230 -Color ([System.Drawing.Color]::FromArgb(27, 36, 48))
    Draw-TrailingText -Graphics $tray.Graphics -Text "Exit" -FontSize 13 -X 772 -Y 318 -Color ([System.Drawing.Color]::FromArgb(27, 36, 48))

    $sepPen = New-Object System.Drawing.Pen ([System.Drawing.Color]::FromArgb(230, 234, 238)), 1
    $tray.Graphics.DrawLine($sepPen, 760, 274, 1010, 274)
    $sepPen.Dispose()

    Draw-Card -Graphics $tray.Graphics -X 78 -Y 356 -Width 494 -Height 188 -BackColor ([System.Drawing.Color]::FromArgb(255, 255, 255))
    Draw-TrailingText -Graphics $tray.Graphics -Text "Audio Output Switched" -FontSize 15 -X 112 -Y 392 -Color ([System.Drawing.Color]::FromArgb(19, 45, 59)) -Bold
    Draw-TrailingText -Graphics $tray.Graphics -Text "Mi Monitor (NVIDIA High Definition Audio)" -FontSize 13 -X 112 -Y 438 -Color ([System.Drawing.Color]::FromArgb(74, 93, 109))

    Save-Png -Bitmap $tray.Bitmap -Path (Join-Path $docsDir "tray-preview.png")
}
finally {
    $tray.Graphics.Dispose()
    $tray.Bitmap.Dispose()
}

$settings = New-Canvas -Width 1200 -Height 820
try {
    $bgBrush = New-BackgroundBrush -Rect (New-Object System.Drawing.Rectangle 0, 0, 1200, 820)
    $settings.Graphics.FillRectangle($bgBrush, 0, 0, 1200, 820)
    $bgBrush.Dispose()

    Draw-Card -Graphics $settings.Graphics -X 150 -Y 72 -Width 900 -Height 650 -BackColor ([System.Drawing.Color]::White)
    Draw-TrailingText -Graphics $settings.Graphics -Text "Audio Switcher Settings" -FontSize 24 -X 198 -Y 120 -Color ([System.Drawing.Color]::FromArgb(19, 45, 59)) -Bold
    Draw-TrailingText -Graphics $settings.Graphics -Text "Choose devices, hotkeys, and notification behavior." -FontSize 13 -X 200 -Y 162 -Color ([System.Drawing.Color]::FromArgb(95, 112, 126))

    Draw-TrailingText -Graphics $settings.Graphics -Text "Output devices" -FontSize 15 -X 198 -Y 214 -Color ([System.Drawing.Color]::FromArgb(24, 47, 61)) -Bold
    Draw-Card -Graphics $settings.Graphics -X 198 -Y 248 -Width 806 -Height 180 -BackColor ([System.Drawing.Color]::FromArgb(249, 251, 252))
    Draw-TrailingText -Graphics $settings.Graphics -Text "✓ NT-USB Mini" -FontSize 13 -X 228 -Y 284 -Color ([System.Drawing.Color]::FromArgb(33, 97, 74))
    Draw-TrailingText -Graphics $settings.Graphics -Text "✓ Mi Monitor (NVIDIA High Definition Audio)" -FontSize 13 -X 228 -Y 324 -Color ([System.Drawing.Color]::FromArgb(33, 97, 74))
    Draw-TrailingText -Graphics $settings.Graphics -Text "○ Realtek Digital Output (Realtek(R) Audio)" -FontSize 13 -X 228 -Y 364 -Color ([System.Drawing.Color]::FromArgb(111, 124, 136))

    Draw-TrailingText -Graphics $settings.Graphics -Text "Hotkey" -FontSize 15 -X 198 -Y 462 -Color ([System.Drawing.Color]::FromArgb(24, 47, 61)) -Bold
    Draw-Card -Graphics $settings.Graphics -X 198 -Y 496 -Width 806 -Height 82 -BackColor ([System.Drawing.Color]::FromArgb(249, 251, 252))
    Draw-TrailingText -Graphics $settings.Graphics -Text "Ctrl    Shift    Alt    Win         F11" -FontSize 14 -X 228 -Y 526 -Color ([System.Drawing.Color]::FromArgb(42, 56, 69))

    Draw-TrailingText -Graphics $settings.Graphics -Text "Notifications" -FontSize 15 -X 198 -Y 604 -Color ([System.Drawing.Color]::FromArgb(24, 47, 61)) -Bold
    Draw-Card -Graphics $settings.Graphics -X 198 -Y 638 -Width 806 -Height 50 -BackColor ([System.Drawing.Color]::FromArgb(249, 251, 252))
    Draw-TrailingText -Graphics $settings.Graphics -Text "Tray balloon: On    Overlay: On    Duration: 1500 ms" -FontSize 13 -X 228 -Y 654 -Color ([System.Drawing.Color]::FromArgb(74, 93, 109))

    Save-Png -Bitmap $settings.Bitmap -Path (Join-Path $docsDir "settings-preview.png")
}
finally {
    $settings.Graphics.Dispose()
    $settings.Bitmap.Dispose()
    $logo.Dispose()
}

Write-Host "Generated:"
Write-Host (Join-Path $docsDir "tray-preview.png")
Write-Host (Join-Path $docsDir "settings-preview.png")
