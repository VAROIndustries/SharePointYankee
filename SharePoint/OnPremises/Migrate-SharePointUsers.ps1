#Requires -Version 5.0
#Requires -RunAsAdministrator

<#
.SYNOPSIS
    Migrates users from one domain to another in SharePoint 2019 on-premises.

.DESCRIPTION
    This script connects to the SharePoint farm, identifies all users from the old domain
    across all site collections, migrates their identities to the new domain using
    Move-SPUser, handles permissions and user profiles, and includes error handling
    and logging throughout.

    The script requires the SharePoint 2019 PowerShell snap-in (Microsoft.SharePoint.PowerShell)
    which is only available on SharePoint servers. Run this script directly on a SharePoint
    server or via a remote session with the snap-in loaded.

    KEY FEATURES:
    - Enumerates all site collections in the farm
    - Identifies users whose login name matches the old domain pattern
    - Migrates each user using Move-SPUser, which preserves permission assignments
    - Optionally updates User Profile Service Application account names
    - Comprehensive logging with timestamps and severity levels
    - Full PowerShell transcript for post-run analysis
    - Graceful error handling: per-user failures are caught and logged without
      stopping the entire migration

    IMPORTANT:
    - Back up your SharePoint farm before running this script.
    - Test in a development environment before running in production.
    - Move-SPUser preserves direct permissions but may not update all profile
      data; review user profiles after the migration completes.
    - The User Profile Synchronization service may need to run after migration
      to fully reconcile profile data.

.PARAMETER OldDomain
    The old Windows domain name (NetBIOS or FQDN portion used in the login name).
    Example: 'contoso'

    The script matches users whose UserLogin contains this string and replaces it
    with the value supplied in -NewDomain.

.PARAMETER NewDomain
    The new Windows domain name (NetBIOS or FQDN portion).
    Example: 'fabrikam'

.PARAMETER CentralAdminURL
    The full URL of the SharePoint Central Administration site.
    Example: 'http://sp-server:2013'

    Used to validate the farm connection and as a reference in log messages.

.PARAMETER LogFilePath
    Full path to the log file where all messages will be written.
    Defaults to 'C:\SharePointUserMigration.log'.
    The directory must exist; the file will be created or appended to if it already exists.

.EXAMPLE
    .\Migrate-SharePointUsers.ps1 `
        -OldDomain 'contoso' `
        -NewDomain 'fabrikam' `
        -CentralAdminURL 'http://sp-server:2013'

    Description:
        Migrates all users whose login name contains 'contoso' to equivalent
        accounts in the 'fabrikam' domain. Logs to the default path
        (C:\SharePointUserMigration.log).

.EXAMPLE
    .\Migrate-SharePointUsers.ps1 `
        -OldDomain 'contoso' `
        -NewDomain 'fabrikam' `
        -CentralAdminURL 'http://sp-server:2013' `
        -LogFilePath 'D:\Logs\SPMigration.log'

    Description:
        Same migration with a custom log file path.

.NOTES
    AUTHOR:
        Geoff Varosky

    VERSION:
        1.0.0

    LAST UPDATED:
        2026-03-23

    REQUIREMENTS:
        - PowerShell 5.1 on Windows Server
        - Must run on a SharePoint server with the Microsoft.SharePoint.PowerShell snap-in
        - SharePoint Farm Administrator rights
        - #Requires -RunAsAdministrator: must run in an elevated PowerShell session

    SUPPORTED PLATFORMS:
        - SharePoint 2013 / 2016 / 2019 on-premises
        - Does NOT apply to SharePoint Online (use the Microsoft 365 admin center
          or Azure AD tooling for UPN changes in SPO)

    LIMITATIONS:
        - Move-SPUser updates direct permissions but does not guarantee full profile sync.
          Run a User Profile Sync after migration.
        - If a user exists in multiple site collections, Move-SPUser is called once per
          site collection — this is the expected behavior for on-premises environments.
        - The script does not handle AD group memberships or claims-based auth tokens
          beyond what Move-SPUser natively updates.

    GitHub: https://github.com/VAROIndustries/SharePointYankee
#>

param (
    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$OldDomain,

    [Parameter(Mandatory = $true)]
    [ValidateNotNullOrEmpty()]
    [string]$NewDomain,

    [Parameter(Mandatory = $true)]
    [ValidatePattern('^https?://')]
    [string]$CentralAdminURL,

    [Parameter(Mandatory = $false)]
    [string]$LogFilePath = 'C:\SharePointUserMigration.log'
)

#region Logging

function Write-Log {
    param (
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter(Mandatory = $false)]
        [ValidateSet('INFO', 'WARN', 'ERROR')]
        [string]$Level = 'INFO'
    )

    $timestamp  = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logMessage = "[$timestamp] [$Level] $Message"

    Write-Host $logMessage
    Add-Content -Path $LogFilePath -Value $logMessage
}

#endregion Logging

# Start a full transcript alongside the structured log.
Start-Transcript -Path $LogFilePath -Append

try {
    Write-Log "Starting SharePoint user migration from domain '$OldDomain' to '$NewDomain'."
    Write-Log "Central Admin URL: $CentralAdminURL"
    Write-Log "Log file: $LogFilePath"

    #region Load SharePoint Snap-In

    # The Microsoft.SharePoint.PowerShell snap-in is required for on-premises
    # SharePoint cmdlets (Get-SPSite, Move-SPUser, etc.). It is only available
    # on servers that have SharePoint installed.
    if ((Get-PSSnapin -Name Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue) -eq $null) {
        Add-PSSnapin Microsoft.SharePoint.PowerShell
        Write-Log "Added Microsoft.SharePoint.PowerShell snap-in."
    } else {
        Write-Log "Microsoft.SharePoint.PowerShell snap-in already loaded."
    }

    #endregion Load SharePoint Snap-In

    #region Enumerate Site Collections

    Write-Log "Retrieving all site collections from the farm..."
    $sites = Get-SPSite -Limit All
    Write-Log "Retrieved $($sites.Count) site collection(s)."

    #endregion Enumerate Site Collections

    #region Migrate Users Across Site Collections

    $totalMigrated = 0
    $totalErrors   = 0

    foreach ($site in $sites) {
        Write-Log "Processing site collection: $($site.Url)"

        try {
            # Retrieve all users in the site collection's root web.
            # -Limit All ensures we do not hit the default 200-user cap.
            $users = Get-SPUser -Web $site.RootWeb -Limit All

            foreach ($user in $users) {
                # Match users whose login name contains the old domain string.
                # Adjust the -like pattern if your login format differs
                # (e.g., "i:0#.w|contoso\user" for Windows claims).
                if ($user.UserLogin -like "*$OldDomain*") {
                    $oldLogin = $user.UserLogin
                    $newLogin = $oldLogin -replace [regex]::Escape($OldDomain), $NewDomain

                    Write-Log "Migrating '$oldLogin' -> '$newLogin' in site '$($site.Url)'."

                    try {
                        # Move-SPUser updates the user's login name and preserves
                        # all direct permission assignments for that user.
                        # -IgnoreSid: do not validate the new account SID against AD
                        #   (useful when migrating between domains where the SIDs differ).
                        Move-SPUser -Identity $oldLogin `
                                    -NewAlias $newLogin `
                                    -IgnoreSid `
                                    -Confirm:$false

                        Write-Log "Successfully migrated '$oldLogin' to '$newLogin'."
                        $totalMigrated++
                    } catch {
                        Write-Log "Error migrating '$oldLogin' to '$newLogin': $($_.Exception.Message)" 'ERROR'
                        $totalErrors++
                    }
                }
            }
        } catch {
            Write-Log "Error processing site collection '$($site.Url)': $($_.Exception.Message)" 'ERROR'
            $totalErrors++
        } finally {
            # Always dispose of SPSite objects to prevent memory leaks in the
            # SharePoint object model. This is mandatory for on-prem scripting.
            if ($site -ne $null) {
                $site.Dispose()
            }
        }
    }

    Write-Log "Site collection pass complete. Migrated: $totalMigrated. Errors: $totalErrors."

    #endregion Migrate Users Across Site Collections

    #region Update User Profile Service

    # Move-SPUser updates direct permissions, but the User Profile Service Application
    # (UPSA) stores account names independently. Update profiles here for users whose
    # AccountName still reflects the old domain.
    Write-Log "Checking User Profile Service Application for additional account updates..."

    $upsa = Get-SPServiceApplication | Where-Object { $_.TypeName -eq 'User Profile Service Application' }

    if ($upsa) {
        Write-Log "Found User Profile Service Application: $($upsa.DisplayName)"

        try {
            $upm      = New-Object Microsoft.Office.Server.UserProfiles.UserProfileManager($upsa)
            $profiles = $upm.GetEnumerator()

            $profilesMigrated = 0
            $profileErrors    = 0

            foreach ($profile in $profiles) {
                if ($profile.AccountName -like "*$OldDomain*") {
                    $oldAccount = $profile.AccountName
                    $newAccount = $oldAccount -replace [regex]::Escape($OldDomain), $NewDomain

                    Write-Log "Updating user profile: '$oldAccount' -> '$newAccount'."

                    try {
                        # Update the AccountName property on the profile and commit the change.
                        # Note: In most cases Move-SPUser handles this, but profiles may need
                        # explicit updates when the UPS sync is not configured or has run recently.
                        $profile.AccountName = $newAccount
                        $profile.Commit()
                        Write-Log "Successfully updated profile for '$newAccount'."
                        $profilesMigrated++
                    } catch {
                        Write-Log "Error updating profile for '$oldAccount': $($_.Exception.Message)" 'ERROR'
                        $profileErrors++
                    }
                }
            }

            Write-Log "User Profile pass complete. Updated: $profilesMigrated. Errors: $profileErrors."
        } catch {
            Write-Log "Error accessing User Profile Manager: $($_.Exception.Message)" 'ERROR'
        }
    } else {
        Write-Log "No User Profile Service Application found. Skipping profile updates."
    }

    #endregion Update User Profile Service

    Write-Log "User migration completed. Total migrated: $totalMigrated. Total errors: $totalErrors."

} catch {
    Write-Log "A fatal error occurred during migration: $($_.Exception.Message)" 'ERROR'
} finally {
    Stop-Transcript
}

Write-Log "Script execution finished."
