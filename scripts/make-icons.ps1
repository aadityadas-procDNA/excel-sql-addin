# Draws the add-in icons (database cylinder on a blue tile) into public/assets.
Add-Type -AssemblyName System.Drawing
$out = Join-Path $PSScriptRoot "..\public\assets"
New-Item -ItemType Directory -Force $out | Out-Null

foreach ($size in 16, 32, 64, 80) {
    $s = $size / 32.0
    $bmp = New-Object System.Drawing.Bitmap $size, $size
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear([System.Drawing.Color]::Transparent)

    $blue = [System.Drawing.Color]::FromArgb(15, 108, 189)
    $tile = New-Object System.Drawing.Drawing2D.GraphicsPath
    $r = 7 * $s; $d = 2 * $r; $w = $size - 1
    $tile.AddArc(0, 0, $d, $d, 180, 90); $tile.AddArc($w - $d, 0, $d, $d, 270, 90)
    $tile.AddArc($w - $d, $w - $d, $d, $d, 0, 90); $tile.AddArc(0, $w - $d, $d, $d, 90, 90)
    $tile.CloseFigure()
    $g.FillPath((New-Object System.Drawing.SolidBrush $blue), $tile)

    $white = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::White)
    $x = 8 * $s; $cw = 16 * $s; $eh = 6 * $s; $top = 5 * $s; $bottom = 27 * $s - $eh
    $g.FillRectangle($white, $x, $top + $eh / 2, $cw, $bottom - $top)
    $g.FillEllipse($white, $x, $bottom, $cw, $eh)
    $pen = New-Object System.Drawing.Pen $blue, ([Math]::Max(1, 1.6 * $s))
    foreach ($y in ($top + 7.5 * $s), ($top + 14 * $s)) { $g.DrawArc($pen, $x, $y - $eh / 2, $cw, $eh, 0, 180) }
    $g.FillEllipse($white, $x, $top, $cw, $eh)
    $g.DrawEllipse($pen, $x, $top, $cw, $eh)

    $bmp.Save((Join-Path $out "icon-$size.png"), [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()
}
Write-Host "Icons written to $out"
