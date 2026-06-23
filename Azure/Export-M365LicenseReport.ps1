<#
.SYNOPSIS
    Exports a comprehensive Microsoft 365 license usage report to CSV.

.DESCRIPTION
    Connects to Microsoft Graph and retrieves all subscribed SKUs in the tenant,
    then produces a report showing each license SKU with its friendly name, total
    license count, assigned count, and available count.

    Friendly names are resolved from a built-in mapping table covering the most
    common commercial, business, and F-tier SKUs.  Any SKU not in the mapping
    table falls back to the raw SKU part name returned by the Graph API.

.PARAMETER OutputPath
    Full path to the output CSV file.  Defaults to
    "M365LicenseReport_<timestamp>.csv" in the current directory.

.PARAMETER IncludeDisabled
    When specified, SKUs whose CapabilityStatus is "Suspended" or "Warning"
    (i.e., the subscription is not fully active) are included in the report.
    By default only "Enabled" SKUs are returned.

.EXAMPLE
    .\Export-M365LicenseReport.ps1 -OutputPath "C:\Reports\licenses.csv"

    Connects interactively and writes the report to the specified path.

.EXAMPLE
    .\Export-M365LicenseReport.ps1 -IncludeDisabled

    Includes suspended/warning SKUs in the output.  The file is written to the
    current directory with an auto-generated timestamp name.

.NOTES
    Author  : Geoff Varosky
    Module  : Microsoft.Graph.Identity.DirectoryManagement (part of Microsoft.Graph)
    Version : 1.0.0
    Requires: PowerShell 5.1 or 7+, Microsoft.Graph module v2+
    GitHub  : https://github.com/VAROIndustries/SharePointYankee

    Required Graph scopes:
        Organization.Read.All
        LicenseAssignment.ReadWrite.All   (read-only; ReadWrite grants Read)
        -- OR --
        Directory.Read.All

    Authentication: Uses Connect-MgGraph with delegated (interactive) auth by
    default.  For unattended/scheduled use, pre-authenticate with a certificate:
        Connect-MgGraph -TenantId <tid> -ClientId <appid> -CertificateThumbprint <thumb>

    The SKU ID mapping table is current as of 2025.  Microsoft periodically adds
    new product SKUs.  For an authoritative list see:
    https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference
#>

#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'Microsoft.Graph.Identity.DirectoryManagement'; ModuleVersion = '2.0.0' }

[CmdletBinding()]
param (
    [Parameter(
        HelpMessage = 'Full path for the output CSV file.  Defaults to current directory with timestamp.'
    )]
    [string]$OutputPath = (Join-Path -Path (Get-Location) -ChildPath ("M365LicenseReport_{0}.csv" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))),

    [Parameter(
        HelpMessage = 'Include SKUs with a CapabilityStatus other than Enabled (e.g., Suspended, Warning).'
    )]
    [switch]$IncludeDisabled
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region --- Friendly-name lookup table ---
# Keys are the GUID-format SkuId values returned by the Graph API.
# Values are human-readable product names.
# Source: https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference
$SkuFriendlyNames = @{
    # Microsoft 365 E-tier
    '06ebc4ee-1bb5-47dd-8120-11324bc54e06' = 'Microsoft 365 E5'
    '05e9a617-0261-4cee-bb44-138d3ef5d965' = 'Microsoft 365 E3'
    'd61d61cc-f992-433f-a577-5bd016037eeb' = 'Microsoft 365 E3 (No Teams)'
    'cd2925a3-5076-4233-8931-638a8c94f773' = 'Microsoft 365 E5 (No Teams)'
    '18181a46-0d4e-45cd-891e-60aabd171b4e' = 'Office 365 E1'
    '6fd2c87f-b296-42f0-b197-1e91e994b900' = 'Office 365 E3'
    'c7df2760-2c81-4ef7-b578-5b5392b571df' = 'Office 365 E5'

    # Microsoft 365 Business
    'cbdc14ab-d96c-4c30-b9f4-6ada7cdc1d46' = 'Microsoft 365 Business Premium'
    'f245ecc8-75af-4f8e-b61f-27d8114de5f3' = 'Microsoft 365 Business Standard'
    'o365-smb'                              = 'Microsoft 365 Business Basic'
    'b214fe43-f9a3-4aab-8f5d-40c7b29c73c3' = 'Microsoft 365 Business Basic'

    # F-tier (Frontline)
    '66b55226-6b4f-492c-910c-a3b7a3c9d993' = 'Microsoft 365 F1'
    'dcb1a3ae-b33f-4487-846a-a640262fadf4' = 'Microsoft 365 F3'
    '4b585984-651b-448a-9e53-3b10f069cf7f' = 'Office 365 F3'

    # Azure Active Directory / Entra
    '078d2b04-f1bd-4111-bbd4-b4b1b354cef4' = 'Azure AD Premium P1'
    '84a661c4-e949-4bd2-a560-ed7766fcaf2b' = 'Azure AD Premium P2'
    'b05e124f-c7cc-45a0-a6aa-8cf78c946968' = 'Enterprise Mobility + Security E5'
    'efccb6f7-5641-4e0e-bd10-b4976e1bf68e' = 'Enterprise Mobility + Security E3'

    # Microsoft 365 Apps
    'c2273bd0-dff7-4215-9ef5-2c7bcfb06425' = 'Microsoft 365 Apps for Enterprise'
    'b8facc9b-4e7c-4f56-b8e3-6a8b4a8abe70' = 'Microsoft 365 Apps for Business'
    '3b555118-da6a-4418-894f-7df1e2096870' = 'Office 365 Business Essentials'

    # Standalone / Add-ons
    'a403ebcc-fae0-4ca2-8c8c-7a907fd6c235' = 'Power BI Pro'
    'f8a1db68-be16-40ed-86d5-cb42ce701560' = 'Power BI Premium Per User'
    'b30411f5-fea1-4a59-9ad9-3db7c7ead579' = 'Power Apps per User Plan'
    '55c46414-f4de-488e-9b30-79d27b9f0e53' = 'Microsoft Teams Rooms Standard'
    '4fb214cb-a430-4a91-9c91-4976763aa272' = 'Microsoft Teams Rooms Pro'
    '26124093-3d78-432b-b5dc-48bf992543d5' = 'Microsoft Defender for Office 365 Plan 1'
    '47794cd0-f0e5-45c5-9033-2eb6b5fc84e0' = 'Microsoft Defender for Office 365 Plan 2'
    '111046dd-295b-4d6d-9724-d52ac90bd1f2' = 'Microsoft Defender for Endpoint Plan 2'
    '726a0894-2c77-4d65-99da-9775ef05aad1' = 'Microsoft Defender for Business'

    # Exchange / Email
    '19ec0d23-8335-4cbd-94ac-6050e30712fa' = 'Exchange Online Plan 2'
    'efeefc69-f64b-4923-b4f7-98d3b5b7b2a7' = 'Exchange Online Plan 1'
    '3b4fe198-2b5b-4dda-88a5-d5b2edf200dc' = 'Exchange Online Kiosk'

    # Intune
    'c1ec4a95-1f05-45b3-a911-aa3fa01094f5' = 'Microsoft Intune Plan 1'
    'b17653a4-2443-4e8c-a550-18249dda78bb' = 'Microsoft Intune Suite'

    # Visio / Project
    '38b434d2-a15e-4cde-9a98-e737c75623c1' = 'Visio Plan 2'
    'b2c81ec7-4c3c-4c3e-b279-5b049f56a37c' = 'Visio Plan 1'
    '09015f9f-377f-4538-bbb5-f75ceb09739a' = 'Project Plan 5'
    '53818b1b-4a27-454b-8896-0dba576410e6' = 'Project Plan 3'
    'beb6439c-caad-48d3-bf46-0c82871e12be' = 'Project Plan 1'

    # Developer / Educational
    'c42b9cae-ea4f-4ab7-9717-81576235ccac' = 'Microsoft 365 E5 Developer (without Windows and Audio Conferencing)'
    'aceabfd8-3b98-43a4-84d9-494a9b7e3f8e' = 'Microsoft 365 A5 for Faculty'
    'e97c048c-37a4-45a3-beff-4b10be3b4c92' = 'Microsoft 365 A3 for Faculty'
}
#endregion

#region --- Helper: resolve friendly name ---
function Resolve-SkuFriendlyName {
    [CmdletBinding()]
    param (
        [string]$SkuId,
        [string]$SkuPartNumber
    )

    if ($SkuFriendlyNames.ContainsKey($SkuId)) {
        return $SkuFriendlyNames[$SkuId]
    }

    # Fall back to the raw SKU part number with minimal prettification
    return $SkuPartNumber
}
#endregion

#region --- Logging helper ---
function Write-Log {
    [CmdletBinding()]
    param (
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )
    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $entry     = "[$timestamp][$Level] $Message"

    switch ($Level) {
        'WARN'  { Write-Warning $Message }
        'ERROR' { Write-Error   $Message -ErrorAction Continue }
        default { Write-Verbose $entry }
    }
}
#endregion

try {
    #region --- Connect to Microsoft Graph ---
    Write-Log 'Connecting to Microsoft Graph...'

    # Check for an existing session to avoid re-prompting when running iteratively.
    $context = Get-MgContext -ErrorAction SilentlyContinue
    if (-not $context) {
        Connect-MgGraph -Scopes 'Organization.Read.All', 'Directory.Read.All' -NoWelcome
        $context = Get-MgContext
    }

    Write-Log "Connected as: $($context.Account) | Tenant: $($context.TenantId)"
    #endregion

    #region --- Retrieve subscribed SKUs ---
    Write-Log 'Retrieving subscribed SKUs from Microsoft Graph...'

    $skus = Get-MgSubscribedSku -All

    if (-not $IncludeDisabled) {
        $filteredSkus = $skus | Where-Object { $_.CapabilityStatus -eq 'Enabled' }
        $excludedCount = ($skus | Measure-Object).Count - ($filteredSkus | Measure-Object).Count
        if ($excludedCount -gt 0) {
            Write-Log "$excludedCount SKU(s) excluded because CapabilityStatus is not 'Enabled'.  Use -IncludeDisabled to include them." -Level WARN
        }
        $skus = $filteredSkus
    }

    if (($skus | Measure-Object).Count -eq 0) {
        Write-Log 'No licensed SKUs found in this tenant.' -Level WARN
    }
    #endregion

    #region --- Build report rows ---
    Write-Log 'Building license report...'
    $report = [System.Collections.Generic.List[PSCustomObject]]::new()
    $i      = 0

    foreach ($sku in $skus) {
        $i++
        $total     = $sku.PrepaidUnits.Enabled + $sku.PrepaidUnits.Warning
        $assigned  = $sku.ConsumedUnits
        $available = [Math]::Max(0, $total - $assigned)

        Write-Progress -Activity 'Processing license SKUs' `
                       -Status  "SKU $i of $(($skus | Measure-Object).Count): $($sku.SkuPartNumber)" `
                       -PercentComplete ([int](($i / [Math]::Max(1, ($skus | Measure-Object).Count)) * 100))

        $row = [PSCustomObject]@{
            SkuId            = $sku.SkuId
            SkuPartNumber    = $sku.SkuPartNumber
            FriendlyName     = Resolve-SkuFriendlyName -SkuId $sku.SkuId -SkuPartNumber $sku.SkuPartNumber
            CapabilityStatus = $sku.CapabilityStatus
            TotalLicenses    = $total
            AssignedLicenses = $assigned
            AvailableLicenses = $available
            PercentUsed      = if ($total -gt 0) { [Math]::Round(($assigned / $total) * 100, 1) } else { 0 }
        }

        $report.Add($row)
    }

    Write-Progress -Activity 'Processing license SKUs' -Completed
    #endregion

    #region --- Export to CSV ---
    # Ensure the output directory exists before writing.
    $outputDir = Split-Path -Path $OutputPath -Parent
    if ($outputDir -and -not (Test-Path -Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        Write-Log "Created output directory: $outputDir"
    }

    $report | Sort-Object FriendlyName |
              Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

    Write-Log "Report exported to: $OutputPath"
    #endregion

    #region --- Console summary ---
    $totalSKUs      = ($report | Measure-Object).Count
    $totalLicenses  = ($report | Measure-Object -Property TotalLicenses    -Sum).Sum
    $totalAssigned  = ($report | Measure-Object -Property AssignedLicenses -Sum).Sum
    $totalAvailable = ($report | Measure-Object -Property AvailableLicenses -Sum).Sum

    Write-Host ''
    Write-Host '---- Microsoft 365 License Summary ----' -ForegroundColor Cyan
    Write-Host ("SKUs reported  : {0}"   -f $totalSKUs)     -ForegroundColor White
    Write-Host ("Total licenses : {0}"   -f $totalLicenses)  -ForegroundColor White
    Write-Host ("Assigned       : {0}"   -f $totalAssigned)  -ForegroundColor White
    Write-Host ("Available      : {0}"   -f $totalAvailable) -ForegroundColor White
    Write-Host ''
    Write-Host "Output file    : $OutputPath" -ForegroundColor Green
    #endregion
}
catch {
    Write-Error "Script failed: $_"
    exit 1
}
finally {
    Write-Progress -Activity 'Processing license SKUs' -Completed -ErrorAction SilentlyContinue

    # Disconnect only if the script itself established the connection.
    # If the caller was already connected, leave the session open.
    if ($context -and -not (Get-MgContext -ErrorAction SilentlyContinue)) {
        Disconnect-MgGraph -ErrorAction SilentlyContinue | Out-Null
    }
}
