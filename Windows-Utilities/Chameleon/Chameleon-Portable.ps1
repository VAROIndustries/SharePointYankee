<#
.SYNOPSIS
    Chameleon (Portable) - runs entirely from a USB stick / any folder with no
    installation. Reads its settings from Chameleon.config.json sitting next to
    this script, and logs next to itself.

.DESCRIPTION
    Same behaviour as Chameleon.ps1 but nothing is written to the host machine:
    config, and log all live in the script's own folder ($PSScriptRoot), so you
    can carry it on a USB drive and run it anywhere via Chameleon-Portable.bat.

    On first run, if Chameleon.config.json is missing it is created from the
    example and the script exits so you can fill it in.

.NOTES
    - The Wi-Fi profiles named in the config must already be saved on whatever
      machine you run this on.
    - Run Get-AttachedDevices.ps1 to find the device match string.
#>

$scriptDir  = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$configPath = Join-Path $scriptDir 'Chameleon.config.json'
$logFile    = Join-Path $scriptDir 'Chameleon-Portable.log'

function Write-Log {
    param([string]$Message)
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $Message"
    Add-Content -Path $logFile -Value $line
}

# ---- Load config (portable: lives next to the script) ----
if (-not (Test-Path $configPath)) {
    $example = Join-Path $scriptDir 'Chameleon.config.example.json'
    if (Test-Path $example) { Copy-Item $example $configPath }
    else {
@'
{
  "deviceMatch": "*Your Monitor Or Dock Name*",
  "wifiWhenConnected": "OfficeWiFi",
  "wifiWhenDisconnected": "HomeWiFi",
  "wifiInterface": "Wi-Fi"
}
'@ | Set-Content -Path $configPath -Encoding UTF8
    }
    Write-Host "Created $configPath - edit it with your device and Wi-Fi names, then re-run." -ForegroundColor Yellow
    return
}

$cfg = Get-Content $configPath -Raw | ConvertFrom-Json
$targetDeviceMatch    = $cfg.deviceMatch
$wifiWhenConnected    = $cfg.wifiWhenConnected
$wifiWhenDisconnected = $cfg.wifiWhenDisconnected
$wifiInterface        = if ($cfg.wifiInterface) { $cfg.wifiInterface } else { "Wi-Fi" }

# Resolve netsh by full path so it works even when System32 isn't on PATH.
$netsh = Join-Path $env:WINDIR 'System32\netsh.exe'
if (-not (Test-Path $netsh)) { $netsh = 'netsh' }

function Get-CurrentWifiProfile {
    $line = (& $netsh wlan show interfaces | Select-String '^\s*Profile\s*:' | Select-Object -First 1).Line
    if ($line) { ($line -split ':', 2)[1].Trim() } else { '' }
}

function Set-WifiNetwork {
    param([string]$ProfileName)
    if ($Global:ChameleonCurrentProfile -eq $ProfileName) { return }
    # Don't disconnect if we're already on the target network.
    if ((Get-CurrentWifiProfile) -eq $ProfileName) { $Global:ChameleonCurrentProfile = $ProfileName; return }
    $Global:ChameleonCurrentProfile = $ProfileName

    Write-Log "Switching Wi-Fi to '$ProfileName'"
    & $netsh wlan disconnect interface="$wifiInterface" | Out-Null
    Start-Sleep -Seconds 2
    & $netsh wlan connect name="$ProfileName" interface="$wifiInterface" | Out-Null
}

Write-Log "Chameleon (Portable) started. Target device match: '$targetDeviceMatch'"

$alreadyPresent = Get-PnpDevice -PresentOnly | Where-Object { $_.Name -like $targetDeviceMatch }
if ($alreadyPresent) {
    Set-WifiNetwork -ProfileName $wifiWhenConnected
} else {
    Set-WifiNetwork -ProfileName $wifiWhenDisconnected
}

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

try {
    while ($true) { Start-Sleep -Seconds 3600 }
}
finally {
    Unregister-Event -SourceIdentifier "ChameleonWatcher" -ErrorAction SilentlyContinue
    Write-Log "Chameleon (Portable) stopped."
}
