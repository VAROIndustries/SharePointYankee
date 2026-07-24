<#
.SYNOPSIS
    Chameleon (Tray) - a system-tray version of the Wi-Fi-on-dock switcher.

.DESCRIPTION
    Sits in the notification area as a lizard icon (green = target device
    present / on the "docked" network, grey = absent / on the "undocked"
    network). Right-click for a menu: toggle auto-switching, force either
    network, edit settings, view the log, About, or exit. Balloon tips
    announce each switch. Double-click the icon to open Settings.

    Detection polls every $pollSeconds via Get-PnpDevice, which plays nicely
    with the WinForms message loop. Config and log live next to the script
    (or next to the .exe when compiled with Build-Tray.ps1).

.NOTES
    - Wi-Fi profiles named below must already be saved on this machine.
    - Run Get-AttachedDevices.ps1 to find the device match string.
    - Launch silently (no console) with Chameleon-Tray.vbs.
#>

# ---- CONFIGURE THESE (defaults; overridden by Chameleon.config.json / Settings) ----
$targetDeviceMatch    = "*Your Monitor Or Dock Name*"
$wifiWhenConnected    = "OfficeWiFi"
$wifiWhenDisconnected = "HomeWiFi"
$wifiInterface        = "Wi-Fi"
$pollSeconds          = 4
# --------------------------

$ChameleonVersion = "1.0.0"
$ChameleonEmoji   = [System.Char]::ConvertFromUtf32(0x1F98E)   # lizard 🦎

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Resolve our own folder whether run as a .ps1 or compiled to a .exe (ps2exe).
$scriptDir =
    if     ($PSScriptRoot)                 { $PSScriptRoot }
    elseif ($MyInvocation.MyCommand.Path)  { Split-Path -Parent $MyInvocation.MyCommand.Path }
    else   { Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName) }
$logFile    = Join-Path $scriptDir 'Chameleon-Tray.log'
$configPath = Join-Path $scriptDir 'Chameleon.config.json'

# Resolve netsh by full path so it works even when System32 isn't on PATH.
$netsh = Join-Path $env:WINDIR 'System32\netsh.exe'
if (-not (Test-Path $netsh)) { $netsh = 'netsh' }

# Shared state
$script:autoSwitch     = $true
$script:currentProfile = $null
$script:lastPresent    = $null

function Write-Log {
    param([string]$Message)
    Add-Content -Path $logFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $Message"
}

function Import-ChameleonConfig {
    if (Test-Path $configPath) {
        try {
            $cfg = Get-Content $configPath -Raw | ConvertFrom-Json
            if ($cfg.deviceMatch)          { $script:targetDeviceMatch    = $cfg.deviceMatch }
            if ($cfg.wifiWhenConnected)    { $script:wifiWhenConnected    = $cfg.wifiWhenConnected }
            if ($cfg.wifiWhenDisconnected) { $script:wifiWhenDisconnected = $cfg.wifiWhenDisconnected }
            if ($cfg.wifiInterface)        { $script:wifiInterface        = $cfg.wifiInterface }
        } catch { Write-Log "Could not read config: $($_.Exception.Message)" }
    }
}

function Save-ChameleonConfig {
    $obj = [ordered]@{
        deviceMatch          = $targetDeviceMatch
        wifiWhenConnected    = $wifiWhenConnected
        wifiWhenDisconnected = $wifiWhenDisconnected
        wifiInterface        = $wifiInterface
    }
    ($obj | ConvertTo-Json) | Set-Content -Path $configPath -Encoding UTF8
    Write-Log "Config saved to $configPath"
}

# ---- Start-with-Windows (per-user, no admin needed) ----
$runKey  = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$runName = 'Chameleon'

function Get-SelfLaunchCommand {
    $exePath = [System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    $hostExe = [System.IO.Path]::GetFileName($exePath).ToLower()
    if ($hostExe -eq 'powershell.exe' -or $hostExe -eq 'pwsh.exe') {
        # Running as a .ps1 - prefer the silent .vbs launcher (no console window).
        $vbs = Join-Path $scriptDir 'Chameleon-Tray.vbs'
        if (Test-Path $vbs) { return "wscript.exe `"$vbs`"" }
        $ps1 = Join-Path $scriptDir 'Chameleon-Tray.ps1'
        return "`"$exePath`" -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ps1`""
    }
    # Compiled .exe - launch it directly.
    return "`"$exePath`""
}

function Test-StartupEnabled {
    $null -ne (Get-ItemProperty -Path $runKey -Name $runName -ErrorAction SilentlyContinue)
}

function Set-Startup {
    param([bool]$Enabled)
    if ($Enabled) {
        if (-not (Test-Path $runKey)) { New-Item -Path $runKey -Force | Out-Null }
        Set-ItemProperty -Path $runKey -Name $runName -Value (Get-SelfLaunchCommand)
        Write-Log "Enabled 'Start with Windows'."
    } else {
        Remove-ItemProperty -Path $runKey -Name $runName -ErrorAction SilentlyContinue
        Write-Log "Disabled 'Start with Windows'."
    }
}

function Get-CurrentWifiProfile {
    $line = (& $netsh wlan show interfaces | Select-String '^\s*Profile\s*:' | Select-Object -First 1).Line
    if ($line) { ($line -split ':', 2)[1].Trim() } else { '' }
}

function Test-DevicePresent {
    [bool](Get-PnpDevice -PresentOnly -ErrorAction SilentlyContinue |
           Where-Object { $_.Name -like $targetDeviceMatch })
}

function New-LizardIcon {
    param([System.Drawing.Color]$Tint)
    $bmp = New-Object System.Drawing.Bitmap 16, 16
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode     = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint = [System.Drawing.Text.TextRenderingHint]::AntiAliasGridFit
    $g.Clear([System.Drawing.Color]::Transparent)
    $font = New-Object System.Drawing.Font('Segoe UI Emoji', 12, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
    $sf   = New-Object System.Drawing.StringFormat
    $sf.Alignment = [System.Drawing.StringAlignment]::Center
    $sf.LineAlignment = [System.Drawing.StringAlignment]::Center
    $brush = New-Object System.Drawing.SolidBrush $Tint
    $rect  = New-Object System.Drawing.RectangleF 0, 0, 16, 16
    $g.DrawString($ChameleonEmoji, $font, $brush, $rect, $sf)
    $g.Dispose(); $brush.Dispose(); $font.Dispose(); $sf.Dispose()
    $icon = [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
    $bmp.Dispose()
    return $icon
}

$iconOn  = New-LizardIcon ([System.Drawing.Color]::LimeGreen)
$iconOff = New-LizardIcon ([System.Drawing.Color]::Gray)

function Set-WifiNetwork {
    param([string]$ProfileName, [switch]$Announce)
    if ($script:currentProfile -eq $ProfileName) { return }

    # Don't disconnect if we're already on the target network (avoids a needless blip).
    if ((Get-CurrentWifiProfile) -eq $ProfileName) {
        $script:currentProfile = $ProfileName
        Write-Log "Already on '$ProfileName' - no switch needed."
        return
    }
    $script:currentProfile = $ProfileName

    Write-Log "Switching Wi-Fi to '$ProfileName'"
    & $netsh wlan disconnect interface="$wifiInterface" | Out-Null
    Start-Sleep -Milliseconds 800
    & $netsh wlan connect name="$ProfileName" interface="$wifiInterface" | Out-Null

    if ($Announce) {
        $notify.ShowBalloonTip(3000, "Chameleon", "Switched to '$ProfileName'",
            [System.Windows.Forms.ToolTipIcon]::Info)
    }
}

# ---- Settings dialog ----
function Show-Settings {
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Chameleon Settings"
    $form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $form.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $form.MaximizeBox = $false; $form.MinimizeBox = $false; $form.TopMost = $true
    $form.ClientSize = New-Object System.Drawing.Size 440, 250
    try { $form.Icon = $iconOn } catch {}

    $labels = @('Device match (wildcard):', 'Docked Wi-Fi profile:', 'Undocked Wi-Fi profile:', 'Wi-Fi interface:')
    $vals   = @($targetDeviceMatch, $wifiWhenConnected, $wifiWhenDisconnected, $wifiInterface)
    $boxes  = @()
    for ($i = 0; $i -lt 4; $i++) {
        $lbl = New-Object System.Windows.Forms.Label
        $lbl.Text = $labels[$i]; $lbl.SetBounds(16, 22 + $i * 44, 180, 20)
        $form.Controls.Add($lbl)
        $tb = New-Object System.Windows.Forms.TextBox
        $tb.Text = $vals[$i]; $tb.SetBounds(200, 20 + $i * 44, 224, 24)
        $form.Controls.Add($tb)
        $boxes += $tb
    }

    $hint = New-Object System.Windows.Forms.Label
    $hint.Text = "Wi-Fi profiles must already be saved. Tip: run Get-AttachedDevices.ps1 for the device name."
    $hint.ForeColor = [System.Drawing.Color]::FromArgb(120, 120, 120)
    $hint.SetBounds(16, 196, 408, 16); $hint.Font = New-Object System.Drawing.Font('Segoe UI', 7.5)
    $form.Controls.Add($hint)

    $save = New-Object System.Windows.Forms.Button
    $save.Text = 'Save'; $save.SetBounds(248, 214, 80, 28)
    $save.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'; $cancel.SetBounds(340, 214, 80, 28)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.AddRange(@($save, $cancel))
    $form.AcceptButton = $save; $form.CancelButton = $cancel

    if ($form.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:targetDeviceMatch    = $boxes[0].Text.Trim()
        $script:wifiWhenConnected    = $boxes[1].Text.Trim()
        $script:wifiWhenDisconnected = $boxes[2].Text.Trim()
        $script:wifiInterface        = $boxes[3].Text.Trim()
        Save-ChameleonConfig
        Update-MenuLabels
        $script:currentProfile = $null
        $script:lastPresent    = $null
        Update-State
        Write-Log "Settings updated from the tray."
    }
    $form.Dispose()
}

# ---- About dialog (styled after VARO Industries' other tools) ----
function Show-About {
    $f = New-Object System.Windows.Forms.Form
    $f.Text = "About Chameleon"
    $f.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $f.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $f.MaximizeBox = $false; $f.MinimizeBox = $false; $f.TopMost = $true
    $f.ClientSize = New-Object System.Drawing.Size 320, 240
    try { $f.Icon = $iconOn } catch {}

    $name = New-Object System.Windows.Forms.Label
    $name.Text = "$ChameleonEmoji Chameleon"
    $name.Font = New-Object System.Drawing.Font('Segoe UI Emoji', 16, [System.Drawing.FontStyle]::Bold)
    $name.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $name.SetBounds(0, 18, 320, 34); $f.Controls.Add($name)

    $ver = New-Object System.Windows.Forms.Label
    $ver.Text = "Version $ChameleonVersion"; $ver.ForeColor = [System.Drawing.Color]::FromArgb(85, 85, 85)
    $ver.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $ver.SetBounds(0, 54, 320, 20); $f.Controls.Add($ver)

    $tag = New-Object System.Windows.Forms.Label
    $tag.Text = "Your Wi-Fi, adapted to your surroundings."
    $tag.ForeColor = [System.Drawing.Color]::FromArgb(51, 51, 51)
    $tag.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $tag.SetBounds(0, 80, 320, 20); $f.Controls.Add($tag)

    $mkLink = {
        param($text, $url, $y)
        $l = New-Object System.Windows.Forms.LinkLabel
        $l.Text = $text; $l.Tag = $url
        $l.LinkColor = [System.Drawing.Color]::FromArgb(26, 110, 200)
        $l.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
        $l.SetBounds(0, $y, 320, 20)
        $l.Add_LinkClicked({ Start-Process $this.Tag })
        $f.Controls.Add($l)
    }
    & $mkLink 'varo.industries/tools/chameleon' 'https://varo.industries/tools/chameleon' 112
    & $mkLink 'github.com/VAROIndustries/Chameleon' 'https://github.com/VAROIndustries/Chameleon' 134

    $cop = New-Object System.Windows.Forms.Label
    $cop.Text = ([System.Char]0x00A9) + " 2026 VAR" + ([System.Char]0x00D8) + " Industries"
    $cop.ForeColor = [System.Drawing.Color]::FromArgb(136, 136, 136)
    $cop.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $cop.SetBounds(0, 164, 320, 20); $f.Controls.Add($cop)

    $close = New-Object System.Windows.Forms.Button
    $close.Text = 'Close'; $close.SetBounds(120, 194, 80, 28)
    $close.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $f.Controls.Add($close); $f.AcceptButton = $close
    $f.ShowDialog() | Out-Null
    $f.Dispose()
}

# ---- Tray icon + menu ----
$notify = New-Object System.Windows.Forms.NotifyIcon
$notify.Icon    = $iconOff
$notify.Text    = "Chameleon"
$notify.Visible = $true

$menu = New-Object System.Windows.Forms.ContextMenuStrip

$statusItem = New-Object System.Windows.Forms.ToolStripMenuItem "Starting..."
$statusItem.Enabled = $false
$menu.Items.Add($statusItem) | Out-Null
$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

$autoItem = New-Object System.Windows.Forms.ToolStripMenuItem "Auto-switch"
$autoItem.CheckOnClick = $true
$autoItem.Checked = $script:autoSwitch
$autoItem.Add_Click({
    $script:autoSwitch = $autoItem.Checked
    Write-Log "Auto-switch set to $($script:autoSwitch)"
})
$menu.Items.Add($autoItem) | Out-Null

$startupItem = New-Object System.Windows.Forms.ToolStripMenuItem "Start with Windows"
$startupItem.CheckOnClick = $true
$startupItem.Checked = Test-StartupEnabled
$startupItem.Add_Click({
    try { Set-Startup -Enabled $startupItem.Checked }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Couldn't update startup setting:`n$($_.Exception.Message)", "Chameleon") | Out-Null
        $startupItem.Checked = Test-StartupEnabled
    }
})
$menu.Items.Add($startupItem) | Out-Null

$connectItem = New-Object System.Windows.Forms.ToolStripMenuItem "Force docked network"
$connectItem.Add_Click({ Set-WifiNetwork -ProfileName $wifiWhenConnected -Announce })
$menu.Items.Add($connectItem) | Out-Null

$disconnectItem = New-Object System.Windows.Forms.ToolStripMenuItem "Force undocked network"
$disconnectItem.Add_Click({ Set-WifiNetwork -ProfileName $wifiWhenDisconnected -Announce })
$menu.Items.Add($disconnectItem) | Out-Null

$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

$settingsItem = New-Object System.Windows.Forms.ToolStripMenuItem "Settings..."
$settingsItem.Add_Click({ Show-Settings })
$menu.Items.Add($settingsItem) | Out-Null

$aboutItem = New-Object System.Windows.Forms.ToolStripMenuItem "About"
$aboutItem.Add_Click({ Show-About })
$menu.Items.Add($aboutItem) | Out-Null

$logItem = New-Object System.Windows.Forms.ToolStripMenuItem "Open log"
$logItem.Add_Click({
    if (Test-Path $logFile) { Start-Process notepad.exe $logFile }
    else { [System.Windows.Forms.MessageBox]::Show("No log yet.", "Chameleon") | Out-Null }
})
$menu.Items.Add($logItem) | Out-Null

$menu.Items.Add((New-Object System.Windows.Forms.ToolStripSeparator)) | Out-Null

$exitItem = New-Object System.Windows.Forms.ToolStripMenuItem "Exit"
$exitItem.Add_Click({
    $timer.Stop()
    $notify.Visible = $false
    $notify.Dispose()
    Write-Log "Chameleon (Tray) exited."
    [System.Windows.Forms.Application]::Exit()
})
$menu.Items.Add($exitItem) | Out-Null

$notify.ContextMenuStrip = $menu
$notify.Add_MouseDoubleClick({ Show-Settings })

function Update-MenuLabels {
    $connectItem.Text    = "Force docked network: '$wifiWhenConnected'"
    $disconnectItem.Text = "Force undocked network: '$wifiWhenDisconnected'"
}

# ---- Poll loop ----
function Update-State {
    param([switch]$Startup)
    $present = Test-DevicePresent

    if ($present) {
        $notify.Icon = $iconOn
        $statusItem.Text = "Dock detected -> $wifiWhenConnected"
        $notify.Text = "Chameleon - docked ($wifiWhenConnected)"
    } else {
        $notify.Icon = $iconOff
        $statusItem.Text = "No dock -> $wifiWhenDisconnected"
        $notify.Text = "Chameleon - undocked ($wifiWhenDisconnected)"
    }

    if ($present -ne $script:lastPresent) {
        if ($script:autoSwitch) {
            $target = if ($present) { $wifiWhenConnected } else { $wifiWhenDisconnected }
            if ($Startup) { Set-WifiNetwork -ProfileName $target }
            else          { Set-WifiNetwork -ProfileName $target -Announce }
        }
        $script:lastPresent = $present
    }
}

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = [int]($pollSeconds * 1000)
$timer.Add_Tick({ Update-State })

Import-ChameleonConfig
Update-MenuLabels
Write-Log "Chameleon (Tray) started. Target device match: '$targetDeviceMatch'"
Update-State -Startup
$timer.Start()

$notify.ShowBalloonTip(2000, "Chameleon", "Watching for '$targetDeviceMatch'",
    [System.Windows.Forms.ToolTipIcon]::Info)

$appContext = New-Object System.Windows.Forms.ApplicationContext
[System.Windows.Forms.Application]::Run($appContext)
