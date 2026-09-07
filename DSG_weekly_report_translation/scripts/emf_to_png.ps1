# EMF -> PNG at 4x scale (GDI+). Plan-slide screenshots are EMF vector pastes;
# 4x renders crisp 8pt text for OCR / visual reading.
# Usage: powershell -File emf_to_png.ps1 <input.emf or dir> [outdir]
param(
    [Parameter(Mandatory=$true)][string]$Path,
    [string]$OutDir = ""
)
Add-Type -AssemblyName System.Drawing
if ($OutDir -ne "" -and -not (Test-Path $OutDir)) { New-Item -ItemType Directory -Path $OutDir | Out-Null }
$files = @()
if (Test-Path $Path -PathType Container) { $files = Get-ChildItem $Path -Filter *.emf } else { $files = @(Get-Item $Path) }
foreach ($f in $files) {
    $img = [System.Drawing.Image]::FromFile($f.FullName)
    $scale = 4
    $w = [int]($img.Width * $scale); $h = [int]($img.Height * $scale)
    $bmp = New-Object System.Drawing.Bitmap ($w), ($h)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.Clear([System.Drawing.Color]::White)
    $g.InterpolationMode = 'HighQualityBicubic'
    $g.TextRenderingHint = 'AntiAlias'
    $g.DrawImage($img, 0, 0, $w, $h)
    $out = if ($OutDir -ne "") { Join-Path $OutDir ($f.BaseName + ".png") } else { [System.IO.Path]::ChangeExtension($f.FullName, ".png") }
    $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose(); $img.Dispose()
    Write-Output "$($f.Name) -> $out ($w x $h)"
}
