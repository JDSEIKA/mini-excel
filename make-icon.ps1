# Genera assets\app.ico con varios tamaños (PNG dentro de ICO).
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

function New-RoundedPath([float]$x, [float]$y, [float]$w, [float]$h, [float]$r) {
    $p = New-Object System.Drawing.Drawing2D.GraphicsPath
    $d = $r * 2
    $p.AddArc($x, $y, $d, $d, 180, 90)
    $p.AddArc($x + $w - $d, $y, $d, $d, 270, 90)
    $p.AddArc($x + $w - $d, $y + $h - $d, $d, $d, 0, 90)
    $p.AddArc($x, $y + $h - $d, $d, $d, 90, 90)
    $p.CloseFigure()
    return $p
}

function New-IconPng([int]$size) {
    $bmp = New-Object System.Drawing.Bitmap($size, $size, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.Clear([System.Drawing.Color]::Transparent)
    $s = $size / 256.0
    $g.ScaleTransform($s, $s)

    $bgPath = New-RoundedPath 8 8 240 240 46
    $brush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        (New-Object System.Drawing.Point(0, 0)),
        (New-Object System.Drawing.Point(256, 256)),
        [System.Drawing.Color]::FromArgb(255, 46, 139, 87),
        [System.Drawing.Color]::FromArgb(255, 24, 90, 56))
    $g.FillPath($brush, $bgPath)
    $brush.Dispose()

    # Hoja blanca
    $shadow = New-RoundedPath 53 47 158 178 15
    $sb = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(55, 0, 0, 0))
    $g.FillPath($sb, $shadow)
    $sb.Dispose()
    $sheet = New-RoundedPath 49 39 158 178 14
    $wb = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 255, 255, 255))
    $g.FillPath($wb, $sheet)
    $wb.Dispose()

    # Cabecera de la hoja
    $g.SetClip($sheet)
    $hb = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 226, 244, 233))
    $g.FillRectangle($hb, 49, 39, 158, 38)
    $hb.Dispose()
    $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 33, 115, 70), 4)
    $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    for ($i = 1; $i -le 3; $i++) {
        $x = 49 + ($i * 158 / 4)
        $g.DrawLine($pen, $x, 46, $x, 210)
    }
    for ($j = 1; $j -le 5; $j++) {
        $y = 39 + ($j * 178 / 6)
        $g.DrawLine($pen, 55, $y, 201, $y)
    }
    $g.ResetClip()
    $pen.Dispose()

    # Acento amarillo: celda activa
    $accent = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 255, 196, 0))
    $g.FillRectangle($accent, 128, 155, 37, 29)
    $accent.Dispose()
    $border = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 33, 115, 70), 3)
    $g.DrawRectangle($border, 128, 155, 37, 29)
    $border.Dispose()

    # Borde exterior de la hoja
    $edge = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(255, 21, 78, 49), 2)
    $g.DrawPath($edge, $sheet)
    $edge.Dispose()

    $g.Dispose()
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    return $ms.ToArray()
}

$sizes = @(16, 20, 24, 32, 40, 48, 64, 128, 256)
$images = New-Object 'System.Collections.Generic.List[byte[]]'
foreach ($sz in $sizes) { $images.Add([byte[]](New-IconPng $sz)) }

$outPath = Join-Path $PSScriptRoot 'assets\app.ico'
$fs = [System.IO.File]::Open($outPath, [System.IO.FileMode]::Create)
$bw = New-Object System.IO.BinaryWriter($fs)
$bw.Write([UInt16]0)
$bw.Write([UInt16]1)
$bw.Write([UInt16]$sizes.Count)
$offset = 6 + 16 * $sizes.Count
for ($i = 0; $i -lt $sizes.Count; $i++) {
    $sz = $sizes[$i]
    $data = $images[$i]
    $dim = if ($sz -ge 256) { 0 } else { $sz }
    $bw.Write([Byte]$dim)
    $bw.Write([Byte]$dim)
    $bw.Write([Byte]0)
    $bw.Write([Byte]0)
    $bw.Write([UInt16]1)
    $bw.Write([UInt16]32)
    $bw.Write([UInt32]$data.Length)
    $bw.Write([UInt32]$offset)
    $offset += $data.Length
}
foreach ($data in $images) { $bw.Write($data, 0, $data.Length) }
$bw.Flush()
$bw.Close()
Write-Host ("Icono creado: {0} ({1} bytes)" -f $outPath, (Get-Item $outPath).Length)
