<#
.SYNOPSIS
    Checks the health of a SharePoint On-Premises farm across one or more servers.

.DESCRIPTION
    Test-SPFarmHealth performs a structured health check against each server in a
    SharePoint farm. Checks include disk free space, CPU utilization, memory utilization,
    SharePoint Timer Service status, SharePoint Search Host Controller status, and IIS
    (W3SVC) service status.

    Results are emitted as PSCustomObjects with consistent Server / Check / Status / Value /
    Threshold properties, making them easy to pipe, filter, or export. Pass -OutputPath to
    write a CSV report automatically.

    All WMI calls use Get-CimInstance (PS 5.1+ / PS 7 compatible). The SharePoint snap-in
    is loaded only when -FarmUrl is supplied to enumerate servers automatically; if you pass
    -ServerName directly the snap-in is not required on the machine running the script
    (though it must be present on each target server).

.PARAMETER FarmUrl
    URL of any SharePoint site in the farm. When supplied, the script loads the SharePoint
    PowerShell snap-in and enumerates all farm servers automatically. Cannot be combined
    with -ServerName.

.PARAMETER ServerName
    One or more server hostnames or IP addresses to check. Use this when you want to target
    specific servers without loading the SharePoint snap-in locally. Cannot be combined with
    -FarmUrl.

.PARAMETER DiskWarningThresholdGB
    Free disk space threshold in GB. Drives with less free space than this value are flagged
    as Warning or Critical. Default: 10 GB.

.PARAMETER CpuWarningThreshold
    CPU utilization percentage at or above which a Warning is raised. Default: 85 %.

.PARAMETER MemoryWarningThreshold
    Memory utilization percentage at or above which a Warning is raised. Default: 90 %.

.PARAMETER OutputPath
    Optional. Full path to a CSV file where results will be written. The file is created or
    overwritten. Example: C:\Logs\SPFarmHealth_2026-06-23.csv

.PARAMETER Credential
    Optional PSCredential used for remote CIM sessions. If omitted, the current user context
    is used (Kerberos / NTLM passthrough). Required when the running account does not have
    administrative rights on remote servers.

.EXAMPLE
    # Check farm servers via SharePoint snap-in, output to console
    .\Test-SPFarmHealth.ps1 -FarmUrl "https://sharepoint.contoso.com"

.EXAMPLE
    # Check specific servers with custom thresholds, export to CSV
    .\Test-SPFarmHealth.ps1 -ServerName "SP-WFE01","SP-APP01","SP-SRCH01" `
        -DiskWarningThresholdGB 20 -CpuWarningThreshold 80 -MemoryWarningThreshold 85 `
        -OutputPath "C:\Logs\SPFarmHealth.csv"

.EXAMPLE
    # Pipe results to filter for non-OK items only
    .\Test-SPFarmHealth.ps1 -ServerName "SP-WFE01","SP-APP01" |
        Where-Object { $_.Status -ne 'OK' }

.NOTES
    Author  : Geoff Varosky
    Website : https://sharepointyankee.com
    Version : 1.0
    Requires: PowerShell 5.1 or later (PowerShell 7 recommended for cross-platform use)
              Microsoft.SharePoint.PowerShell snap-in (only when using -FarmUrl)
              CIM/WMI access (TCP 5985/5986 or DCOM) to each target server
              Local admin rights on each target server

    The -FarmUrl parameter requires this script to run on a machine with the SharePoint
    management tools installed (e.g., an application server or a machine with the SP
    Management Shell). If running remotely without the snap-in, use -ServerName instead.

    Disk checks iterate all fixed logical disks (DriveType = 3). A drive below
    DiskWarningThresholdGB / 2 is escalated to Critical; below the threshold is Warning.
#>

#Requires -Version 5.1

[CmdletBinding(DefaultParameterSetName = 'ByServerName')]
param (
    [Parameter(Mandatory, ParameterSetName = 'ByFarmUrl', HelpMessage = 'URL of any site in the farm.')]
    [ValidateNotNullOrEmpty()]
    [string] $FarmUrl,

    [Parameter(Mandatory, ParameterSetName = 'ByServerName', HelpMessage = 'Array of server hostnames to check.')]
    [ValidateNotNullOrEmpty()]
    [string[]] $ServerName,

    [Parameter()]
    [ValidateRange(1, 9999)]
    [double] $DiskWarningThresholdGB = 10,

    [Parameter()]
    [ValidateRange(1, 100)]
    [int] $CpuWarningThreshold = 85,

    [Parameter()]
    [ValidateRange(1, 100)]
    [int] $MemoryWarningThreshold = 90,

    [Parameter()]
    [string] $OutputPath,

    [Parameter()]
    [PSCredential] $Credential
)

#region --- Initialization ---

$ErrorActionPreference = 'Stop'
$results = [System.Collections.Generic.List[PSCustomObject]]::new()

# Helper: emit a structured result object and add it to $results
function New-HealthResult {
    [OutputType([PSCustomObject])]
    param (
        [string] $Server,
        [string] $Check,
        [ValidateSet('OK', 'Warning', 'Critical', 'Error')]
        [string] $Status,
        [string] $Value,
        [string] $Threshold,
        [string] $Detail
    )
    $obj = [PSCustomObject]@{
        Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Server    = $Server
        Check     = $Check
        Status    = $Status
        Value     = $Value
        Threshold = $Threshold
        Detail    = $Detail
    }
    $results.Add($obj)
    return $obj
}

#endregion

#region --- Server Enumeration ---

if ($PSCmdlet.ParameterSetName -eq 'ByFarmUrl') {
    Write-Verbose "Loading SharePoint PowerShell snap-in to enumerate farm servers..."

    if ((Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null) {
        try {
            Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction Stop
        } catch {
            throw "Failed to load Microsoft.SharePoint.PowerShell snap-in. " +
                  "Run this script from a machine with SharePoint management tools installed. Error: $_"
        }
    }

    Write-Verbose "Connecting to farm via: $FarmUrl"
    try {
        $spFarm = Get-SPFarm -ErrorAction Stop
        $servers = $spFarm.Servers |
            Where-Object { $_.Role -ne 'Invalid' } |
            Select-Object -ExpandProperty Address
        Write-Verbose "Discovered $($servers.Count) farm server(s): $($servers -join ', ')"
    } catch {
        throw "Unable to retrieve farm server list. Ensure you are running from a SharePoint server and the farm is accessible. Error: $_"
    }
} else {
    $servers = $ServerName
}

if (-not $servers -or $servers.Count -eq 0) {
    throw "No servers to check. Verify the farm is reachable or supply server names via -ServerName."
}

#endregion

#region --- Per-Server Health Checks ---

$totalServers = $servers.Count
$currentServer = 0

foreach ($server in $servers) {
    $currentServer++
    Write-Progress -Activity "SharePoint Farm Health Check" `
                   -Status "Checking $server ($currentServer of $totalServers)" `
                   -PercentComplete (($currentServer / $totalServers) * 100)

    Write-Verbose "--- Checking server: $server ---"

    # Build CIM session options — reuse a single session per server for all checks
    $cimSession = $null
    $sessionError = $null

    try {
        $cimParams = @{ ComputerName = $server; ErrorAction = 'Stop' }
        if ($Credential) { $cimParams['Credential'] = $Credential }

        # Prefer WSMan (PS Remoting / WinRM); fall back to DCOM for older servers
        $sessionOption = New-CimSessionOption -Protocol Wsman
        try {
            $cimSession = New-CimSession @cimParams -SessionOption $sessionOption
        } catch {
            Write-Verbose "WSMan failed for $server, falling back to DCOM."
            $sessionOption = New-CimSessionOption -Protocol Dcom
            $cimSession = New-CimSession @cimParams -SessionOption $sessionOption
        }
    } catch {
        $sessionError = $_.Exception.Message
        New-HealthResult -Server $server -Check 'CIM Connection' -Status 'Error' `
            -Value 'Failed' -Threshold 'N/A' `
            -Detail "Could not establish CIM session: $sessionError" | Write-Output
        continue  # Skip remaining checks for this server — no session available
    }

    try {

        #-- 1. Disk Free Space ---------------------------------------------------
        Write-Verbose "$server: Checking disk free space..."
        try {
            $disks = Get-CimInstance -CimSession $cimSession -ClassName Win32_LogicalDisk `
                        -Filter "DriveType = 3" -ErrorAction Stop |
                     Select-Object DeviceID, Size, FreeSpace

            foreach ($disk in $disks) {
                $freeGB   = [math]::Round($disk.FreeSpace / 1GB, 2)
                $totalGB  = [math]::Round($disk.Size / 1GB, 2)
                $pctFree  = if ($disk.Size -gt 0) { [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 1) } else { 0 }
                $critical = [math]::Round($DiskWarningThresholdGB / 2, 2)

                $diskStatus = if ($freeGB -le $critical) {
                    'Critical'
                } elseif ($freeGB -le $DiskWarningThresholdGB) {
                    'Warning'
                } else {
                    'OK'
                }

                New-HealthResult -Server $server `
                    -Check "Disk Free ($($disk.DeviceID))" `
                    -Status $diskStatus `
                    -Value "$freeGB GB free of $totalGB GB ($pctFree% free)" `
                    -Threshold "Warning < $DiskWarningThresholdGB GB, Critical < $critical GB" `
                    -Detail '' | Write-Output
            }
        } catch {
            New-HealthResult -Server $server -Check 'Disk Free' -Status 'Error' `
                -Value 'Query Failed' -Threshold "$DiskWarningThresholdGB GB" `
                -Detail $_.Exception.Message | Write-Output
        }

        #-- 2. CPU Utilization ---------------------------------------------------
        Write-Verbose "$server: Checking CPU utilization..."
        try {
            # Sample LoadPercentage from Win32_Processor — average across all sockets
            $cpuInstances = Get-CimInstance -CimSession $cimSession -ClassName Win32_Processor `
                                -ErrorAction Stop
            $cpuAvg = [math]::Round(($cpuInstances | Measure-Object -Property LoadPercentage -Average).Average, 1)

            $cpuStatus = if ($cpuAvg -ge $CpuWarningThreshold) { 'Warning' } else { 'OK' }

            New-HealthResult -Server $server -Check 'CPU Utilization' -Status $cpuStatus `
                -Value "$cpuAvg %" -Threshold "Warning >= $CpuWarningThreshold %" `
                -Detail '' | Write-Output
        } catch {
            New-HealthResult -Server $server -Check 'CPU Utilization' -Status 'Error' `
                -Value 'Query Failed' -Threshold "$CpuWarningThreshold %" `
                -Detail $_.Exception.Message | Write-Output
        }

        #-- 3. Memory Utilization ------------------------------------------------
        Write-Verbose "$server: Checking memory utilization..."
        try {
            $os = Get-CimInstance -CimSession $cimSession -ClassName Win32_OperatingSystem `
                    -ErrorAction Stop |
                  Select-Object TotalVisibleMemorySize, FreePhysicalMemory

            $totalMB = [math]::Round($os.TotalVisibleMemorySize / 1KB, 0)
            $freeMB  = [math]::Round($os.FreePhysicalMemory / 1KB, 0)
            $usedPct = [math]::Round((($totalMB - $freeMB) / $totalMB) * 100, 1)

            $memStatus = if ($usedPct -ge $MemoryWarningThreshold) { 'Warning' } else { 'OK' }

            New-HealthResult -Server $server -Check 'Memory Utilization' -Status $memStatus `
                -Value "$usedPct % used ($freeMB MB free of $totalMB MB)" `
                -Threshold "Warning >= $MemoryWarningThreshold %" `
                -Detail '' | Write-Output
        } catch {
            New-HealthResult -Server $server -Check 'Memory Utilization' -Status 'Error' `
                -Value 'Query Failed' -Threshold "$MemoryWarningThreshold %" `
                -Detail $_.Exception.Message | Write-Output
        }

        #-- 4. SharePoint Timer Service (SPTimerV4) ------------------------------
        Write-Verbose "$server: Checking SharePoint Timer Service..."
        try {
            $timerSvc = Get-CimInstance -CimSession $cimSession -ClassName Win32_Service `
                            -Filter "Name = 'SPTimerV4'" -ErrorAction Stop

            if (-not $timerSvc) {
                New-HealthResult -Server $server -Check 'SP Timer Service' -Status 'Warning' `
                    -Value 'Not Found' -Threshold 'Running' `
                    -Detail 'SPTimerV4 service was not found on this server. It may not be a SharePoint server.' | Write-Output
            } else {
                $timerStatus = if ($timerSvc.State -eq 'Running') { 'OK' } else { 'Critical' }
                New-HealthResult -Server $server -Check 'SP Timer Service' -Status $timerStatus `
                    -Value $timerSvc.State -Threshold 'Running' `
                    -Detail "StartMode: $($timerSvc.StartMode)" | Write-Output
            }
        } catch {
            New-HealthResult -Server $server -Check 'SP Timer Service' -Status 'Error' `
                -Value 'Query Failed' -Threshold 'Running' `
                -Detail $_.Exception.Message | Write-Output
        }

        #-- 5. SharePoint Search Host Controller (OSearch16 / OSearch15) ---------
        Write-Verbose "$server: Checking SharePoint Search service..."
        try {
            # Try SP 2016/2019 name first, fall back to SP 2013 name
            $searchSvc = Get-CimInstance -CimSession $cimSession -ClassName Win32_Service `
                            -Filter "Name = 'OSearch16' OR Name = 'OSearch15'" -ErrorAction Stop |
                         Select-Object -First 1

            if (-not $searchSvc) {
                # Search Host Controller is only present on servers running the Search role
                New-HealthResult -Server $server -Check 'SP Search Service' -Status 'OK' `
                    -Value 'Not Present' -Threshold 'N/A' `
                    -Detail 'OSearch15/OSearch16 not found — this server may not host the Search role.' | Write-Output
            } else {
                $searchStatus = if ($searchSvc.State -eq 'Running') { 'OK' } else { 'Critical' }
                New-HealthResult -Server $server -Check 'SP Search Service' -Status $searchStatus `
                    -Value "$($searchSvc.Name): $($searchSvc.State)" -Threshold 'Running' `
                    -Detail "StartMode: $($searchSvc.StartMode)" | Write-Output
            }
        } catch {
            New-HealthResult -Server $server -Check 'SP Search Service' -Status 'Error' `
                -Value 'Query Failed' -Threshold 'Running' `
                -Detail $_.Exception.Message | Write-Output
        }

        #-- 6. IIS / World Wide Web Publishing Service (W3SVC) -------------------
        Write-Verbose "$server: Checking IIS (W3SVC)..."
        try {
            $w3svc = Get-CimInstance -CimSession $cimSession -ClassName Win32_Service `
                        -Filter "Name = 'W3SVC'" -ErrorAction Stop

            if (-not $w3svc) {
                New-HealthResult -Server $server -Check 'IIS (W3SVC)' -Status 'Warning' `
                    -Value 'Not Found' -Threshold 'Running' `
                    -Detail 'W3SVC not found — IIS may not be installed on this server.' | Write-Output
            } else {
                $iisStatus = if ($w3svc.State -eq 'Running') { 'OK' } else { 'Critical' }
                New-HealthResult -Server $server -Check 'IIS (W3SVC)' -Status $iisStatus `
                    -Value $w3svc.State -Threshold 'Running' `
                    -Detail "StartMode: $($w3svc.StartMode)" | Write-Output
            }
        } catch {
            New-HealthResult -Server $server -Check 'IIS (W3SVC)' -Status 'Error' `
                -Value 'Query Failed' -Threshold 'Running' `
                -Detail $_.Exception.Message | Write-Output
        }

    } finally {
        # Always clean up the CIM session regardless of check outcome
        if ($cimSession) {
            Remove-CimSession -CimSession $cimSession -ErrorAction SilentlyContinue
        }
    }
}

Write-Progress -Activity "SharePoint Farm Health Check" -Completed

#endregion

#region --- Output / Export ---

if ($OutputPath) {
    try {
        $results | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Results exported to: $OutputPath" -ForegroundColor Cyan
    } catch {
        Write-Warning "Failed to write CSV to '$OutputPath': $_"
    }
}

# Summary to console
$warnCount     = ($results | Where-Object { $_.Status -eq 'Warning'  }).Count
$criticalCount = ($results | Where-Object { $_.Status -eq 'Critical' }).Count
$errorCount    = ($results | Where-Object { $_.Status -eq 'Error'    }).Count

Write-Host ""
Write-Host "Farm Health Check Summary" -ForegroundColor White
Write-Host "--------------------------"
Write-Host "Servers checked : $totalServers"
Write-Host "Total checks    : $($results.Count)"
Write-Host "OK              : $(($results | Where-Object { $_.Status -eq 'OK' }).Count)" -ForegroundColor Green

if ($warnCount -gt 0) {
    Write-Host "Warning         : $warnCount" -ForegroundColor Yellow
}
if ($criticalCount -gt 0) {
    Write-Host "Critical        : $criticalCount" -ForegroundColor Red
}
if ($errorCount -gt 0) {
    Write-Host "Error           : $errorCount" -ForegroundColor DarkRed
}
Write-Host ""

# Return all result objects to the pipeline
$results

#endregion
