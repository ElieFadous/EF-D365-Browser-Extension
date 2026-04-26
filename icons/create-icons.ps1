# EF Power Platform Tools — Icon Generator (Angular D + Gold Bolt)
# Run this script once to create the extension icons.
# Requires Windows with .NET Framework (pre-installed on all modern Windows).
#
# Usage: Right-click > Run with PowerShell
#        Or: powershell -ExecutionPolicy Bypass -File create-icons.ps1

Add-Type -AssemblyName System.Drawing

function New-Icon {
    param (
        [int]    $Size,
        [string] $OutputPath
    )

    $bmp = New-Object System.Drawing.Bitmap($Size, $Size)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.CompositingQuality = [System.Drawing.Drawing2D.CompositingQuality]::HighQuality

    $s = $Size / 64.0  # scale factor — paths defined on 64×64 grid

    # ── D365 blue gradient ──────────────────────────────────────────────────
    $gradBrush = New-Object System.Drawing.Drawing2D.LinearGradientBrush(
        [System.Drawing.Point]::new(0, 0),
        [System.Drawing.Point]::new($Size, $Size),
        [System.Drawing.Color]::FromArgb(255, 0, 144, 241),   # #0090F1
        [System.Drawing.Color]::FromArgb(255, 0, 78, 140)     # #004E8C
    )

    # ── Angular D shape (outer - inner = D silhouette) ──────────────────────
    # Draw outer shape, subtract inner with SrcCopy on a temp bitmap to fake evenodd
    $outer = New-Object System.Drawing.Drawing2D.GraphicsPath
    $outerPts = @(
        [System.Drawing.PointF]::new([float](10*$s), [float](6*$s)),
        [System.Drawing.PointF]::new([float](46*$s), [float](6*$s)),
        [System.Drawing.PointF]::new([float](58*$s), [float](18*$s)),
        [System.Drawing.PointF]::new([float](58*$s), [float](46*$s)),
        [System.Drawing.PointF]::new([float](46*$s), [float](58*$s)),
        [System.Drawing.PointF]::new([float](10*$s), [float](58*$s))
    )
    $outer.AddPolygon($outerPts)

    $inner = New-Object System.Drawing.Drawing2D.GraphicsPath
    $innerPts = @(
        [System.Drawing.PointF]::new([float](22*$s), [float](17*$s)),
        [System.Drawing.PointF]::new([float](40*$s), [float](17*$s)),
        [System.Drawing.PointF]::new([float](50*$s), [float](27*$s)),
        [System.Drawing.PointF]::new([float](50*$s), [float](37*$s)),
        [System.Drawing.PointF]::new([float](40*$s), [float](47*$s)),
        [System.Drawing.PointF]::new([float](22*$s), [float](47*$s))
    )
    $inner.AddPolygon($innerPts)

    # Combine outer minus inner using Exclude (evenodd equivalent)
    $dShape = New-Object System.Drawing.Drawing2D.GraphicsPath
    $dShape.FillMode = [System.Drawing.Drawing2D.FillMode]::Alternate
    $dShape.AddPath($outer, $false)
    $dShape.AddPath($inner, $false)
    $g.FillPath($gradBrush, $dShape)

    # ── Gold lightning bolt ─────────────────────────────────────────────────
    $goldBrush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::FromArgb(255, 255, 185, 0))
    $boltPts = @(
        [System.Drawing.PointF]::new([float](36*$s), [float](21*$s)),
        [System.Drawing.PointF]::new([float](28*$s), [float](34*$s)),
        [System.Drawing.PointF]::new([float](33*$s), [float](34*$s)),
        [System.Drawing.PointF]::new([float](27*$s), [float](47*$s)),
        [System.Drawing.PointF]::new([float](43*$s), [float](31*$s)),
        [System.Drawing.PointF]::new([float](37*$s), [float](31*$s))
    )
    $boltPath = New-Object System.Drawing.Drawing2D.GraphicsPath
    $boltPath.AddPolygon($boltPts)
    $g.FillPath($goldBrush, $boltPath)

    # ── Save ────────────────────────────────────────────────────────────────
    $bmp.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $g.Dispose(); $bmp.Dispose()

    Write-Host "  Created: $OutputPath" -ForegroundColor Green
}

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host ""
Write-Host "EF Power Platform Tools — Generating Icons (D365 Angular D + Bolt)" -ForegroundColor Cyan
Write-Host "─────────────────────────────────────────────────────────────────────" -ForegroundColor DarkGray

New-Icon -Size  16 -OutputPath "$scriptDir\icon16.png"
New-Icon -Size  48 -OutputPath "$scriptDir\icon48.png"
New-Icon -Size 128 -OutputPath "$scriptDir\icon128.png"

Write-Host ""
Write-Host "Done! Reload the extension in Edge/Chrome to see the new icon." -ForegroundColor Cyan
Write-Host ""
