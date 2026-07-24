<#
.SYNOPSIS
    Chameleon - watches for a specific PnP device (e.g. an external monitor or
    dock) to connect/disconnect and switches the Wi-Fi network to match.

.DESCRIPTION
    When the target device appears, Chameleon connects to $wifiWhenConnected.
    When it disappears, it connects to $wifiWhenDisconnected. Detection uses a
    WMI/CIM indication subscription so switches happen within a couple seconds
    of docking/undocking.

.NOTES
    - The target Wi-Fi profiles must already be saved on this machine (i.e.
      you've connected to them manually at least once). netsh switches between
      known profiles; it does not add new networks.
    - Run Get-AttachedDevices.ps1 to find the right $targetDeviceMatch string.
    - Confirm your interface name with: netsh wlan show interfaces
    - Install as a Scheduled Task at logon (see README.md).
#>

# ---- CONFIGURE THESE (defaults; overridden by Chameleon.config.json if present) ----
$targetDeviceMatch    = "*Your Monitor Or Dock Name*"  # wildcard match against device Name
$wifiWhenConnected    = "OfficeWiFi"                    # profile to use when device is present
$wifiWhenDisconnected = "HomeWiFi"                      # profile to use when device is absent
$wifiInterface        = "Wi-Fi"                         # netsh wlan show interfaces to confirm
$logFile              = "$env:LOCALAPPDATA\Chameleon.log"
# --------------------------

# Single source of truth: if Chameleon.config.json sits next to this script, it wins.
$scriptDir  = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$configPath = Join-Path $scriptDir 'Chameleon.config.json'
if (Test-Path $configPath) {
    $cfg = Get-Content $configPath -Raw | ConvertFrom-Json
    if ($cfg.deviceMatch)          { $targetDeviceMatch    = $cfg.deviceMatch }
    if ($cfg.wifiWhenConnected)    { $wifiWhenConnected    = $cfg.wifiWhenConnected }
    if ($cfg.wifiWhenDisconnected) { $wifiWhenDisconnected = $cfg.wifiWhenDisconnected }
    if ($cfg.wifiInterface)        { $wifiInterface        = $cfg.wifiInterface }
}

function Write-Log {
    param([string]$Message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $Message"
    Add-Content -Path $logFile -Value $line
}

# Resolve netsh by full path so it works even when System32 isn't on PATH.
$netsh = Join-Path $env:WINDIR 'System32\netsh.exe'
if (-not (Test-Path $netsh)) { $netsh = 'netsh' }

function Get-CurrentWifiProfile {
    $line = (& $netsh wlan show interfaces | Select-String '^\s*Profile\s*:' | Select-Object -First 1).Line
    if ($line) { ($line -split ':', 2)[1].Trim() } else { '' }
}

function Set-WifiNetwork {
    param([string]$ProfileName)

    # Skip if we already switched to this profile (docking fires many PnP events).
    if ($Global:ChameleonCurrentProfile -eq $ProfileName) { return }
    # Don't disconnect if we're already on the target network.
    if ((Get-CurrentWifiProfile) -eq $ProfileName) { $Global:ChameleonCurrentProfile = $ProfileName; return }
    $Global:ChameleonCurrentProfile = $ProfileName

    Write-Log "Switching Wi-Fi to '$ProfileName'"
    & $netsh wlan disconnect interface="$wifiInterface" | Out-Null
    Start-Sleep -Seconds 2
    & $netsh wlan connect name="$ProfileName" interface="$wifiInterface" | Out-Null
}

Write-Log "Chameleon started. Target device match: '$targetDeviceMatch'"

# Fire once on startup based on current state, so it's correct even if the
# device was already connected before this script started.
$alreadyPresent = Get-PnpDevice -PresentOnly | Where-Object { $_.Name -like $targetDeviceMatch }
if ($alreadyPresent) {
    Set-WifiNetwork -ProfileName $wifiWhenConnected
} else {
    Set-WifiNetwork -ProfileName $wifiWhenDisconnected
}

# CIM indication: fires on any PnP device creation/deletion (2s polling window).
$query = "SELECT * FROM __InstanceOperationEvent WITHIN 2 WHERE TargetInstance ISA 'Win32_PnPEntity'"

Register-CimIndicationEvent -Query $query -SourceIdentifier "ChameleonWatcher" -Action {
    $instance  = $Event.SourceEventArgs.NewEvent.TargetInstance
    $eventType = $Event.SourceEventArgs.NewEvent.__CLASS

    if ($instance.Name -like $using:targetDeviceMatch) {
        if ($eventType -eq "__InstanceCreationEvent") {
            Set-WifiNetwork -ProfileName $using:wifiWhenConnected
        }
        elseif ($eventType -eq "__InstanceDeletionEvent") {
            Set-WifiNetwork -ProfileName $using:wifiWhenDisconnected
        }
    }
} | Out-Null

# Keep the script alive so the event subscription stays active.
try {
    while ($true) { Start-Sleep -Seconds 3600 }
}
finally {
    Unregister-Event -SourceIdentifier "ChameleonWatcher" -ErrorAction SilentlyContinue
    Write-Log "Chameleon stopped."
}
