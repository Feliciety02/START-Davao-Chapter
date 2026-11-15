# Rebuild the installable app icons from START-DOST's original PNG.
# Source: https://www.start-dost.org/_next/static/media/logo-s.7f4f7462.png
Add-Type -AssemblyName System.Drawing

$iconDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourcePath = Join-Path $iconDir 'start-dost-logo.png'
$source = [System.Drawing.Image]::FromFile($sourcePath)
$crop = [System.Drawing.RectangleF]::new(317.0, 74.0, 920.0, 1416.0)
$background = [System.Drawing.ColorTranslator]::FromHtml('#F1F1F1')
$yellow = [System.Drawing.ColorTranslator]::FromHtml('#FFDD00')
$ink = [System.Drawing.ColorTranslator]::FromHtml('#12203A')

function Save-StartIcon([string]$name, [int]$size, [bool]$maskable, [bool]$withPill) {
  $canvas = [System.Drawing.Bitmap]::new($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $g = [System.Drawing.Graphics]::FromImage($canvas)
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
  $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
  $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
  $g.Clear($background)

  $scale = $size / 512.0
  if ($withPill) {
    $markHeight = if ($maskable) { 286.0 } else { 358.0 }
    $markTop = if ($maskable) { 66.0 } else { 27.0 }
  } else {
    $markHeight = 430.0
    $markTop = 41.0
  }
  $markWidth = $markHeight * $crop.Width / $crop.Height
  $dest = [System.Drawing.RectangleF]::new([float](($size - $markWidth * $scale) / 2), [float]($markTop * $scale), [float]($markWidth * $scale), [float]($markHeight * $scale))
  $g.DrawImage($source, $dest, $crop, [System.Drawing.GraphicsUnit]::Pixel)

  if ($withPill) {
    $pillWidth = if ($maskable) { 228.0 } else { 296.0 }
    $pillHeight = if ($maskable) { 49.0 } else { 62.0 }
    $pillTop = if ($maskable) { 379.0 } else { 414.0 }
    $x = [float](($size - $pillWidth * $scale) / 2)
    $y = [float]($pillTop * $scale)
    $w = [float]($pillWidth * $scale)
    $h = [float]($pillHeight * $scale)
    $path = [System.Drawing.Drawing2D.GraphicsPath]::new()
    $path.AddArc($x, $y, $h, $h, 90, 180)
    $path.AddArc($x + $w - $h, $y, $h, $h, 270, 180)
    $path.CloseFigure()
    $pillBrush = [System.Drawing.SolidBrush]::new($yellow)
    $g.FillPath($pillBrush, $path)
    $fontBase = if ($maskable) { 27.0 } else { 35.0 }
    $fontSize = [float]($fontBase * $scale)
    $font = [System.Drawing.Font]::new('Segoe UI', $fontSize, [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
    $textBrush = [System.Drawing.SolidBrush]::new($ink)
    $format = [System.Drawing.StringFormat]::new()
    $format.Alignment = [System.Drawing.StringAlignment]::Center
    $format.LineAlignment = [System.Drawing.StringAlignment]::Center
    $g.DrawString('REGION XI', $font, $textBrush, [System.Drawing.RectangleF]::new($x, $y - 1 * $scale, $w, $h), $format)
    $format.Dispose(); $textBrush.Dispose(); $font.Dispose(); $pillBrush.Dispose(); $path.Dispose()
  }
  $canvas.Save((Join-Path $iconDir $name), [System.Drawing.Imaging.ImageFormat]::Png)
  $g.Dispose(); $canvas.Dispose()
}

# The home-screen icon follows the star used on the app's opening screen.
function Save-StarIcon([string]$name, [int]$size, [bool]$maskable) {
  $canvas = [System.Drawing.Bitmap]::new($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
  $g = [System.Drawing.Graphics]::FromImage($canvas)
  $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::HighQuality
  $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
  $g.Clear([System.Drawing.ColorTranslator]::FromHtml('#FFFFFF'))
  $scale = $size / 512.0
  $centerX = 256.0 * $scale
  $centerY = $(if ($maskable) { 237.0 * $scale } else { 225.0 * $scale })
  $outer = $(if ($maskable) { 142.0 * $scale } else { 178.0 * $scale })
  $inner = $outer * 0.63
  $points = [System.Drawing.PointF[]]::new(10)
  for ($i = 0; $i -lt 10; $i++) {
    $angle = -[Math]::PI / 2 + $i * [Math]::PI / 5
    $radius = if ($i % 2) { $inner } else { $outer }
    $points[$i] = [System.Drawing.PointF]::new([float]($centerX + $radius * [Math]::Cos($angle)), [float]($centerY + $radius * [Math]::Sin($angle)))
  }
  $shadow = [System.Drawing.PointF[]]::new(10)
  for ($i = 0; $i -lt 10; $i++) { $shadow[$i] = [System.Drawing.PointF]::new($points[$i].X, $points[$i].Y + 9 * $scale) }
  $navy = [System.Drawing.ColorTranslator]::FromHtml('#12203A')
  $yellow = [System.Drawing.ColorTranslator]::FromHtml('#FFDD00')
  $g.FillPolygon([System.Drawing.SolidBrush]::new($navy), $shadow)
  $g.FillPolygon([System.Drawing.SolidBrush]::new($yellow), $points)
  $pen = [System.Drawing.Pen]::new($navy, [float](10 * $scale))
  $pen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
  $g.DrawPolygon($pen, $points)
  $pen.Dispose()
  $eyeBrush = [System.Drawing.SolidBrush]::new($navy)
  $whiteBrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::White)
  foreach ($x in @(220.0, 292.0)) {
    $g.FillEllipse($eyeBrush, [float](($x - 9) * $scale), [float](($centerY / $scale + 8) * $scale), [float](18 * $scale), [float](24 * $scale))
    $g.FillEllipse($whiteBrush, [float](($x - 5) * $scale), [float](($centerY / $scale + 11) * $scale), [float](6 * $scale), [float](7 * $scale))
  }
  $smile = [System.Drawing.Pen]::new($navy, [float](5 * $scale))
  $g.DrawArc($smile, [float](234 * $scale), [float](($centerY / $scale + 20) * $scale), [float](44 * $scale), [float](26 * $scale), 10, 160)
  $smile.Dispose()
  $pillY = $(if ($maskable) { 382.0 } else { 411.0 })
  $pillW = $(if ($maskable) { 200.0 } else { 248.0 })
  $pillH = 48.0
  $pill = [System.Drawing.RectangleF]::new([float]((512 - $pillW) / 2 * $scale), [float]($pillY * $scale), [float]($pillW * $scale), [float]($pillH * $scale))
  $g.FillEllipse([System.Drawing.SolidBrush]::new($yellow), $pill.X, $pill.Y, $pill.Height, $pill.Height)
  $g.FillEllipse([System.Drawing.SolidBrush]::new($yellow), $pill.Right - $pill.Height, $pill.Y, $pill.Height, $pill.Height)
  $g.FillRectangle([System.Drawing.SolidBrush]::new($yellow), $pill.X + $pill.Height / 2, $pill.Y, $pill.Width - $pill.Height, $pill.Height)
  $font = [System.Drawing.Font]::new('Segoe UI', [float](26 * $scale), [System.Drawing.FontStyle]::Bold, [System.Drawing.GraphicsUnit]::Pixel)
  $format = [System.Drawing.StringFormat]::new()
  $format.Alignment = [System.Drawing.StringAlignment]::Center
  $format.LineAlignment = [System.Drawing.StringAlignment]::Center
  $g.DrawString('REGION XI', $font, $eyeBrush, $pill, $format)
  $format.Dispose(); $font.Dispose(); $whiteBrush.Dispose(); $eyeBrush.Dispose()
  $canvas.Save((Join-Path $iconDir $name), [System.Drawing.Imaging.ImageFormat]::Png)
  $g.Dispose(); $canvas.Dispose()
}

try {
  Save-StartIcon 'icon-512.png' 512 $false $true
  Save-StartIcon 'icon-192.png' 192 $false $true
  Save-StartIcon 'maskable-512.png' 512 $true $true
  Save-StartIcon 'apple-touch-icon.png' 180 $false $true
  Save-StartIcon 'mark-256.png' 256 $false $true
  Save-StartIcon 'favicon-32.png' 32 $false $false
  Save-StarIcon 'star-192.png' 192 $false
  Save-StarIcon 'star-512.png' 512 $false
  Save-StarIcon 'maskable-192.png' 192 $true
  Save-StarIcon 'maskable-512.png' 512 $true
} finally {
  $source.Dispose()
}
