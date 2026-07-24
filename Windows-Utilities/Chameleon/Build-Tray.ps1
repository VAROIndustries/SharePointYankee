<#
.SYNOPSIS
    Compiles Chameleon-Tray.ps1 into a standalone Windows .exe (dist\Chameleon-Tray.exe).

.DESCRIPTION
    Uses the ps2exe module. Produces a no-console (GUI) executable that runs the
    system-tray watcher with no PowerShell window. Generates a green-dot icon at
    assets\chameleon.ico on first run.

    Requires the ps2exe module:
        Install-Module ps2exe -Scope CurrentUser

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Build-Tray.ps1
#>
[CmdletBinding()]
param(
    [string]$Version = "1.0.0"
)

$ErrorActionPreference = 'Stop'
$root    = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$src     = Join-Path $root 'Chameleon-Tray.ps1'
$distDir = Join-Path $root 'dist'
$outExe  = Join-Path $distDir 'Chameleon-Tray.exe'
$iconDir = Join-Path $root 'assets'
$iconOut = Join-Path $iconDir 'chameleon.ico'

if (-not (Test-Path $src)) { throw "Source not found: $src" }
New-Item -ItemType Directory -Force -Path $distDir, $iconDir | Out-Null

# --- Generate a lizard 🦎 .ico if missing (matches the tray icon) ---
if (-not (Test-Path $iconOut)) {
    Add-Type -AssemblyName System.Drawing
    $bmp = New-Object System.Drawing.Bitmap 32, 32
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $g.Clear([System.Drawing.Color]::Transparent)
    $font  = New-Object System.Drawing.Font('Segoe UI Emoji', 24, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
    $sf    = New-Object System.Drawing.StringFormat
    $sf.Alignment = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $brush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::LimeGreen)
    $g.DrawString([System.Char]::ConvertFromUtf32(0x1F98E), $font, $brush,
        (New-Object System.Drawing.RectangleF 0, 0, 32, 32), $sf)
    $g.Dispose(); $brush.Dispose(); $font.Dispose(); $sf.Dispose()
    $icon = [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
    $fs = [System.IO.File]::Create($iconOut)
    $icon.Save($fs)
    $fs.Close(); $bmp.Dispose()
    Write-Host "Generated icon: $iconOut" -ForegroundColor DarkGray
}

# --- Ensure ps2exe is available (it may be installed under the PS7 module path
#     even when we run this under Windows PowerShell 5.1, so search both) ---
try {
    Import-Module ps2exe -ErrorAction Stop
} catch {
    $candidates = @(
        Join-Path $HOME 'OneDrive\Documents\PowerShell\Modules\ps2exe'
        Join-Path $HOME 'OneDrive\Documents\WindowsPowerShell\Modules\ps2exe'
        Join-Path $HOME 'Documents\PowerShell\Modules\ps2exe'
        Join-Path $HOME 'Documents\WindowsPowerShell\Modules\ps2exe'
    )
    $psd1 = $candidates | Where-Object { Test-Path $_ } |
        ForEach-Object { Get-ChildItem -Path $_ -Recurse -Filter 'ps2exe.psd1' -ErrorAction SilentlyContinue } |
        Select-Object -First 1
    if (-not $psd1) { throw "ps2exe module not found. Run:  Install-Module ps2exe -Scope CurrentUser" }
    Import-Module $psd1.FullName -ErrorAction Stop
}

Write-Host "Compiling $src -> $outExe" -ForegroundColor Cyan
Invoke-ps2exe `
    -inputFile  $src `
    -outputFile $outExe `
    -iconFile   $iconOut `
    -noConsole `
    -title   'Chameleon Tray' `
    -description 'Switches Wi-Fi when a target device attaches' `
    -product 'Chameleon' `
    -company 'VARO Industries' `
    -copyright '(c) 2026 VARO Industries' `
    -version $Version

if (Test-Path $outExe) {
    $size = [math]::Round((Get-Item $outExe).Length / 1KB, 1)
    Write-Host "Built: $outExe ($size KB)" -ForegroundColor Green
} else {
    throw "Build failed - $outExe was not produced."
}
