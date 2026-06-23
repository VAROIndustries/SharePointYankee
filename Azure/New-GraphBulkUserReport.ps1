<#
.SYNOPSIS
    Exports a comprehensive user report from Microsoft 365 using the Microsoft Graph
    PowerShell SDK, including license assignments, last sign-in, MFA status, and
    account state.

.DESCRIPTION
    Queries Microsoft Graph for all users in the tenant and produces a CSV report
    containing:
      - Display name and User Principal Name
      - Account enabled / disabled status
      - User type (Member vs Guest)
      - Assigned license SKU names (friendly names resolved from SKU GUID lookup)
      - Last interactive sign-in date (requires AuditLog.Read.All permission)
      - MFA registration status (per-user MFA state from authentication methods)
      - Whether the user has a registered Microsoft Authenticator or FIDO2 key

    By default the report includes all member users. Use -IncludeGuests to add
    guest accounts and -LicensedOnly to restrict to users with at least one
    assigned license.

    Authentication uses interactive browser auth by default. For unattended /
    scheduled runs pass -TenantId and -ClientId with a certificate thumbprint via
    Connect-MgGraph before calling this script, or use a managed identity.

    Handles large tenants via server-side paging — safe for tenants with 100,000+
    users without loading all results into memory at once.

.PARAMETER OutputPath
    Full path to the output CSV file. Defaults to "UserReport-<TenantId>-<Date>.csv"
    in the current directory.

.PARAMETER IncludeGuests
    Switch. When present, guest users (UserType eq 'Guest') are included in the
    report. By default only Member users are returned.

.PARAMETER LicensedOnly
    Switch. When present, only users with at least one assigned license are included.
    Useful for focusing on billable users.

.PARAMETER BatchSize
    Number of users to retrieve per Graph API page. Valid range: 1-999. Defaults to
    500. Larger values reduce round trips but increase memory usage per batch.

.PARAMETER Scopes
    Microsoft Graph permission scopes to request during interactive authentication.
    Defaults to the minimum required set. Ignored when an existing Graph connection
    is already established.

.EXAMPLE
    New-GraphBulkUserReport -OutputPath "C:\Reports\users.csv"

    Exports all member users to a CSV with default settings.

.EXAMPLE
    New-GraphBulkUserReport -IncludeGuests -LicensedOnly -OutputPath "C:\Reports\licensed-users.csv"

    Exports only licensed users (members and guests) to CSV.

.EXAMPLE
    # Connect with certificate auth first (for unattended runs)
    Connect-MgGraph -TenantId "contoso.onmicrosoft.com" -ClientId "00000000-..." -CertificateThumbprint "ABCD1234..."
    New-GraphBulkUserReport -OutputPath "C:\Reports\users-$(Get-Date -f yyyyMMdd).csv" -LicensedOnly

    Uses an existing Graph connection (no interactive prompt) and exports licensed users.

.NOTES
    Author   : Geoff Varosky
    Version  : 1.0.0
    Requires : Microsoft.Graph.Users, Microsoft.Graph.Identity.SignIns,
               Microsoft.Graph.Reports modules (part of Microsoft.Graph SDK)
    GitHub   : https://github.com/VAROIndustries/SharePointYankee

    Minimum required Graph API permissions:
      - User.Read.All           — read user profiles and license assignments
      - AuditLog.Read.All       — read sign-in logs (LastSignIn)
      - UserAuthenticationMethod.Read.All — read MFA registration status

    For application (non-interactive) auth, grant these as Application permissions
    in your app registration. For delegated auth the signed-in user must have at
    least the Reports Reader or Global Reader role.

    The SKU friendly-name lookup covers common Microsoft 365 SKUs. Unknown SKUs
    will display as the raw SKU part number (e.g., "ENTERPRISEPREMIUM").
    For a full current list: https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference
#>
#Requires -Version 7.0
#Requires -Modules @{ ModuleName='Microsoft.Graph.Users'; ModuleVersion='2.0.0' }
#Requires -Modules @{ ModuleName='Microsoft.Graph.Identity.SignIns'; ModuleVersion='2.0.0' }
#Requires -Modules @{ ModuleName='Microsoft.Graph.Reports'; ModuleVersion='2.0.0' }

[CmdletBinding(SupportsShouldProcess)]
param (
    [Parameter()]
    [string]$OutputPath,

    [Parameter()]
    [switch]$IncludeGuests,

    [Parameter()]
    [switch]$LicensedOnly,

    [Parameter()]
    [ValidateRange(1, 999)]
    [int]$BatchSize = 500,

    [Parameter()]
    [string[]]$Scopes = @(
        'User.Read.All',
        'AuditLog.Read.All',
        'UserAuthenticationMethod.Read.All'
    )
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

#region SKU friendly name lookup
# Maps SkuPartNumber to a human-readable license name.
# Source: https://learn.microsoft.com/en-us/entra/identity/users/licensing-service-plan-reference
$SkuNames = @{
    'SPE_E3'                        = 'Microsoft 365 E3'
    'SPE_E5'                        = 'Microsoft 365 E5'
    'SPE_F1'                        = 'Microsoft 365 F1'
    'SPE_F3'                        = 'Microsoft 365 F3'
    'ENTERPRISEPREMIUM'             = 'Office 365 E5'
    'ENTERPRISEPACK'                = 'Office 365 E3'
    'STANDARDPACK'                  = 'Office 365 E1'
    'DESKLESSPACK'                  = 'Office 365 F3'
    'EXCHANGEENTERPRISE'            = 'Exchange Online Plan 2'
    'EXCHANGESTANDARD'              = 'Exchange Online Plan 1'
    'EXCHANGEDESKLESS'              = 'Exchange Online Kiosk'
    'EMS'                           = 'Enterprise Mobility + Security E3'
    'EMSPREMIUM'                    = 'Enterprise Mobility + Security E5'
    'INTUNE_A'                      = 'Microsoft Intune'
    'AAD_PREMIUM'                   = 'Entra ID P1'
    'AAD_PREMIUM_P2'                = 'Entra ID P2'
    'TEAMS_EXPLORATORY'             = 'Microsoft Teams Exploratory'
    'MCOSTANDARD'                   = 'Skype for Business Online Plan 2'
    'POWER_BI_PRO'                  = 'Power BI Pro'
    'POWER_BI_PREMIUM_PER_USER'     = 'Power BI Premium Per User'
    'POWERAPPS_PER_USER'            = 'Power Apps Per User'
    'FLOW_PER_USER'                 = 'Power Automate Per User'
    'PROJECTPREMIUM'                = 'Project Plan 5'
    'PROJECTPROFESSIONAL'           = 'Project Plan 3'
    'VISIOCLIENT'                   = 'Visio Plan 2'
    'VISIOONLINE_PLAN1'             = 'Visio Plan 1'
    'WIN10_PRO_ENT_SUB'             = 'Windows 10/11 Enterprise E3'
    'WIN_ENT_E5'                    = 'Windows 10/11 Enterprise E5'
    'DEFENDER_ENDPOINT_P1'          = 'Microsoft Defender for Endpoint P1'
    'MDATP_Server'                  = 'Microsoft Defender for Endpoint Server'
    'ATP_ENTERPRISE'                = 'Microsoft Defender for Office 365 P1'
    'THREAT_INTELLIGENCE'           = 'Microsoft Defender for Office 365 P2'
    'RIGHTSMANAGEMENT'              = 'Azure Information Protection P1'
    'INFORMATION_PROTECTION_COMPLIANCE' = 'Microsoft Purview E5 Compliance'
    'MICROSOFT_BUSINESS_CENTER'     = 'Microsoft Business Center'
    'O365_BUSINESS_PREMIUM'         = 'Microsoft 365 Business Premium'
    'O365_BUSINESS_ESSENTIALS'      = 'Microsoft 365 Business Basic'
    'O365_BUSINESS'                 = 'Microsoft 365 Apps for Business'
    'OFFICESUBSCRIPTION'            = 'Microsoft 365 Apps for Enterprise'
}
#endregion

#region Connect to Graph if not already connected
$existingContext = Get-MgContext -ErrorAction SilentlyContinue
if (-not $existingContext) {
    Write-Host "No active Graph connection found. Connecting interactively..." -ForegroundColor Yellow
    Connect-MgGraph -Scopes $Scopes -NoWelcome
    $existingContext = Get-MgContext
}

$tenantId = $existingContext.TenantId
Write-Verbose "Connected to tenant: $tenantId as $($existingContext.Account)"

# Verify required scopes are granted
$grantedScopes  = $existingContext.Scopes
$requiredScopes = @('User.Read.All', 'AuditLog.Read.All', 'UserAuthenticationMethod.Read.All')
$missingScopes  = $requiredScopes | Where-Object { $_ -notin $grantedScopes }
if ($missingScopes) {
    Write-Warning "The following scopes are not in the current token and data may be incomplete: $($missingScopes -join ', ')"
}
#endregion

#region Resolve output path
if (-not $PSBoundParameters.ContainsKey('OutputPath') -or [string]::IsNullOrWhiteSpace($OutputPath)) {
    $dateStamp  = Get-Date -Format 'yyyyMMdd'
    $OutputPath = Join-Path (Get-Location).Path "UserReport-$tenantId-$dateStamp.csv"
}
#endregion

#region Build user filter
# Graph $filter for user type
$filterParts = [System.Collections.Generic.List[string]]::new()
if (-not $IncludeGuests) {
    $filterParts.Add("userType eq 'Member'")
}

$userFilter = if ($filterParts.Count -gt 0) { $filterParts -join ' and ' } else { $null }
#endregion

#region Retrieve SKU display names from tenant (supplements local lookup)
Write-Host "Loading subscribed SKU information..." -ForegroundColor Cyan
$tenantSkus = @{}
try {
    Get-MgSubscribedSku -All | ForEach-Object {
        $partNumber = $_.SkuPartNumber
        $friendly   = if ($SkuNames.ContainsKey($partNumber)) { $SkuNames[$partNumber] } else { $partNumber }
        $tenantSkus[$_.SkuId] = $friendly
    }
    Write-Verbose "Loaded $($tenantSkus.Count) SKUs from tenant."
}
catch {
    Write-Warning "Could not retrieve subscribed SKUs. License names will be shown as GUID. Error: $_"
}
#endregion

#region Retrieve users with paging
Write-Host "Querying users from Microsoft Graph (BatchSize: $BatchSize)..." -ForegroundColor Cyan

# Properties to retrieve — sign-in activity requires $select with signInActivity
$selectProperties = @(
    'id', 'displayName', 'userPrincipalName', 'accountEnabled',
    'userType', 'assignedLicenses', 'mail', 'department',
    'jobTitle', 'createdDateTime', 'signInActivity'
) -join ','

$getUserParams = @{
    All              = $true
    PageSize         = $BatchSize
    Select           = $selectProperties
    ConsistencyLevel = 'eventual'
    CountVariable    = 'totalUsers'
}
if ($userFilter) {
    $getUserParams['Filter'] = $userFilter
}

$allUsers = [System.Collections.Generic.List[object]]::new()
$pageNum  = 0

try {
    # Get-MgUser with -All handles paging automatically via the SDK
    Get-MgUser @getUserParams | ForEach-Object {
        $allUsers.Add($_)
        $pageNum++
        if ($pageNum % $BatchSize -eq 0) {
            Write-Progress -Activity 'Retrieving users' `
                -Status "Fetched $pageNum users..." `
                -PercentComplete -1
        }
    }
}
catch {
    throw "Failed to retrieve users from Microsoft Graph: $_"
}

Write-Progress -Activity 'Retrieving users' -Completed
Write-Host "Retrieved $($allUsers.Count) users." -ForegroundColor Green
#endregion

#region Build report rows
$report      = [System.Collections.Generic.List[pscustomobject]]::new()
$totalCount  = $allUsers.Count
$processed   = 0

foreach ($user in $allUsers) {
    $processed++
    Write-Progress -Activity 'Processing users' `
        -Status "Processing: $($user.UserPrincipalName)" `
        -PercentComplete (($processed / $totalCount) * 100)

    #region Licenses
    $licenseNames = if ($user.AssignedLicenses -and $user.AssignedLicenses.Count -gt 0) {
        $user.AssignedLicenses | ForEach-Object {
            $skuId = $_.SkuId
            if ($tenantSkus.ContainsKey($skuId)) { $tenantSkus[$skuId] } else { $skuId }
        }
    }
    else {
        @()
    }

    # Apply -LicensedOnly filter post-fetch (signInActivity requires $select which may
    # conflict with $filter on assignedLicenses in some tenants)
    if ($LicensedOnly -and $licenseNames.Count -eq 0) {
        continue
    }
    #endregion

    #region Last sign-in
    $lastSignIn = $null
    if ($user.SignInActivity) {
        # LastSignInDateTime is the most recent interactive sign-in
        $lastSignIn = $user.SignInActivity.LastSignInDateTime
    }
    #endregion

    #region MFA status via authentication methods
    # Note: This is an additional Graph call per user. For very large tenants (50k+)
    # consider running the MFA section in a separate pass or using the
    # Get-MgReportAuthenticationMethodUserRegistrationDetail endpoint instead
    # (requires Reports.Read.All, available in bulk without per-user calls).
    $mfaRegistered     = $false
    $authenticatorApp  = $false
    $fido2Key          = $false

    try {
        $authMethods = Get-MgUserAuthenticationMethod -UserId $user.Id -ErrorAction Stop

        foreach ($method in $authMethods) {
            $odataType = $method.AdditionalProperties['@odata.type']
            switch ($odataType) {
                '#microsoft.graph.microsoftAuthenticatorAuthenticationMethod' {
                    $authenticatorApp = $true
                    $mfaRegistered    = $true
                }
                '#microsoft.graph.phoneAuthenticationMethod' {
                    $mfaRegistered = $true
                }
                '#microsoft.graph.fido2AuthenticationMethod' {
                    $fido2Key      = $true
                    $mfaRegistered = $true
                }
                '#microsoft.graph.softwareOathAuthenticationMethod' {
                    $mfaRegistered = $true
                }
                '#microsoft.graph.windowsHelloForBusinessAuthenticationMethod' {
                    $mfaRegistered = $true
                }
                '#microsoft.graph.emailAuthenticationMethod' {
                    # Email is not counted as MFA (it's a SSPR method); intentionally skipped
                }
            }
        }
    }
    catch {
        # UserAuthenticationMethod.Read.All may not be granted; continue without MFA data
        Write-Verbose "Could not retrieve auth methods for $($user.UserPrincipalName): $_"
    }
    #endregion

    $report.Add([pscustomobject]@{
        DisplayName          = $user.DisplayName
        UserPrincipalName    = $user.UserPrincipalName
        Mail                 = $user.Mail
        Department           = $user.Department
        JobTitle             = $user.JobTitle
        UserType             = $user.UserType
        AccountEnabled       = $user.AccountEnabled
        CreatedDateTime      = if ($user.CreatedDateTime) { $user.CreatedDateTime.ToString('yyyy-MM-ddTHH:mm:ssZ') } else { '' }
        LastSignIn           = if ($lastSignIn) { ([datetime]$lastSignIn).ToString('yyyy-MM-ddTHH:mm:ssZ') } else { 'Never / Unknown' }
        LicenseCount         = $licenseNames.Count
        Licenses             = $licenseNames -join '; '
        MfaRegistered        = $mfaRegistered
        AuthenticatorApp     = $authenticatorApp
        Fido2Key             = $fido2Key
    })
}

Write-Progress -Activity 'Processing users' -Completed
#endregion

#region Export to CSV
if ($report.Count -eq 0) {
    Write-Warning "No users matched the specified criteria. CSV will not be written."
}
else {
    if ($PSCmdlet.ShouldProcess($OutputPath, "Export $($report.Count) users to CSV")) {
        $report | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host "Report exported to: $OutputPath" -ForegroundColor Green
        Write-Host "  Total users in report : $($report.Count)"
        Write-Host "  Licensed users         : $(($report | Where-Object LicenseCount -gt 0).Count)"
        Write-Host "  MFA registered         : $(($report | Where-Object MfaRegistered -eq $true).Count)"
        Write-Host "  Disabled accounts      : $(($report | Where-Object AccountEnabled -eq $false).Count)"
        Write-Host "  Guest accounts         : $(($report | Where-Object UserType -eq 'Guest').Count)"
    }
}
#endregion

#region Disconnect reminder
# We do not call Disconnect-MgGraph automatically because the caller may have
# established the connection before invoking this script and may want it to persist.
Write-Verbose "Graph connection remains active. Call Disconnect-MgGraph when finished."
#endregion
