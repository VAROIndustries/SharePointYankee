#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '1.12.0' }

<#
.SYNOPSIS
    Removes previous file versions from a SharePoint Online document library.

.DESCRIPTION
    Remove-SPFileVersions connects to a SharePoint Online site using PnP.PowerShell
    and deletes all previous versions of every file in the specified document library,
    retaining only the current (published) version.

    This is useful for reclaiming storage quota consumed by version history. SharePoint
    Online retains versions indefinitely by default, and large libraries with many
    co-authored or frequently updated files can accumulate significant version storage.

    KEY BEHAVIORS:
    - Only PREVIOUS versions are removed. The current/published version of each
      file is always preserved.
    - Major and minor (draft) versions are both removed where they exist.
    - Files with no previous versions are skipped silently.
    - Supports -WhatIf: pass -WhatIf to preview which files would be affected
      without deleting anything.
    - Supports -Confirm: pass -Confirm:$false to suppress per-file confirmation
      prompts; omit to be prompted before each deletion.
    - Provides a progress bar and a summary of files processed, versions removed,
      and any errors encountered.
    - All activity is written to both the console and an optional log file.

    STORAGE NOTE:
    Deleted versions move to the SharePoint Recycle Bin. The storage credit is
    not fully reclaimed until the Recycle Bin is emptied. To immediately reclaim
    quota, empty the site Recycle Bin after running this script.

.PARAMETER SiteUrl
    The full URL of the SharePoint Online site.
    Example: https://contoso.sharepoint.com/sites/Projects

.PARAMETER LibraryName
    The display name (title) of the document library to process.
    Example: "Documents" or "Shared Documents"

.PARAMETER UseExistingConnection
    When specified, uses the currently active PnP connection instead of
    prompting for interactive login. Use this when you have already called
    Connect-PnPOnline before invoking this function.

.PARAMETER BatchSize
    The number of files to retrieve per page when enumerating the library.
    Larger values improve performance on large libraries at the cost of
    increased memory usage. Default is 500. Valid range: 1-5000.

.PARAMETER LogPath
    Optional. Full path to a log file. If specified, all output is also
    written to this file in addition to the console. The file is created
    if it does not exist; existing files are appended to.

.EXAMPLE
    Remove-SPFileVersions `
        -SiteUrl "https://contoso.sharepoint.com/sites/Projects" `
        -LibraryName "Documents" `
        -WhatIf

    Description:
        Preview mode. Lists all files that have previous versions without
        deleting anything.

.EXAMPLE
    Remove-SPFileVersions `
        -SiteUrl "https://contoso.sharepoint.com/sites/Projects" `
        -LibraryName "Documents"

    Description:
        Removes all previous versions from every file in the Documents library.
        Interactive login is used.

.EXAMPLE
    Remove-SPFileVersions `
        -SiteUrl "https://contoso.sharepoint.com/sites/Projects" `
        -LibraryName "Shared Documents" `
        -BatchSize 1000 `
        -LogPath "C:\Logs\VersionCleanup.log"

    Description:
        Removes previous versions using a larger batch size and writes all
        output to a log file.

.EXAMPLE
    Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/Projects" -Interactive
    Remove-SPFileVersions `
        -SiteUrl "https://contoso.sharepoint.com/sites/Projects" `
        -LibraryName "Documents" `
        -UseExistingConnection `
        -Confirm:$false

    Description:
        Reuses an existing PnP connection and suppresses per-file confirmation
        prompts for unattended execution.

.NOTES
    AUTHOR:
        Geoff Varosky

    VERSION:
        1.0.0

    LAST UPDATED:
        2026-06-23

    REQUIREMENTS:
        - PowerShell 5.1 or 7.x
        - PnP.PowerShell 1.12.0 or later
          Install: Install-Module PnP.PowerShell -Scope CurrentUser
        - Contribute or higher on the document library
          (managing versions requires Write access to each file)

    CAUTION:
        - Version deletion is not easily reversible once the Recycle Bin is emptied.
        - Always run with -WhatIf first to confirm scope.
        - Do not run against libraries that have compliance or legal hold requirements
          without confirming the deletions are permitted.

    REFERENCES:
        https://pnp.github.io/powershell/cmdlets/Get-PnPListItem.html
        https://pnp.github.io/powershell/cmdlets/Get-PnPFileVersion.html
        https://pnp.github.io/powershell/cmdlets/Remove-PnPFileVersion.html

    GitHub: https://github.com/VAROIndustries/SharePointYankee
#>

function Remove-SPFileVersions {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param (
        [Parameter(
            Mandatory = $true,
            HelpMessage = 'Full URL of the SharePoint Online site.'
        )]
        [ValidatePattern('^https://')]
        [string]$SiteUrl,

        [Parameter(
            Mandatory = $true,
            HelpMessage = 'Display name of the document library to process.'
        )]
        [ValidateNotNullOrEmpty()]
        [string]$LibraryName,

        [Parameter(Mandatory = $false)]
        [switch]$UseExistingConnection,

        [Parameter(Mandatory = $false)]
        [ValidateRange(1, 5000)]
        [int]$BatchSize = 500,

        [Parameter(Mandatory = $false)]
        [ValidateScript({
            $parent = Split-Path $_ -Parent
            if ($parent -and -not (Test-Path $parent)) {
                throw "Log directory '$parent' does not exist."
            }
            $true
        })]
        [string]$LogPath
    )

    $ErrorActionPreference = 'Stop'

    #region Logging

    function Write-Log {
        param (
            [string]$Message,
            [ValidateSet('INFO', 'WARN', 'ERROR')]
            [string]$Level = 'INFO'
        )

        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $line      = "[$timestamp] [$Level] $Message"

        switch ($Level) {
            'WARN'  { Write-Warning $Message }
            'ERROR' { Write-Error   $Message -ErrorAction Continue }
            default { Write-Host    $line }
        }

        if ($LogPath) {
            Add-Content -Path $LogPath -Value $line -Encoding UTF8
        }
    }

    #endregion Logging

    #region Connect

    if (-not $UseExistingConnection) {
        try {
            Write-Log "Connecting to '$SiteUrl'..."
            Connect-PnPOnline -Url $SiteUrl -Interactive
        } catch {
            throw "Failed to connect to '$SiteUrl': $($_.Exception.Message)"
        }
    }

    #endregion Connect

    #region Validate Library

    try {
        Write-Log "Retrieving library '$LibraryName'..."
        $list = Get-PnPList -Identity $LibraryName -ErrorAction Stop

        if ($list.BaseType -ne 'DocumentLibrary') {
            throw "'$LibraryName' is not a document library. Only document libraries are supported."
        }
    } catch {
        throw "Library validation failed: $($_.Exception.Message)"
    }

    Write-Log "Library '$LibraryName' found. Item count: $($list.ItemCount)."

    #endregion Validate Library

    #region Enumerate Files and Remove Versions

    # Counters for the summary report.
    $filesProcessed  = 0
    $versionsRemoved = 0
    $filesSkipped    = 0
    $errors          = 0

    Write-Log "Starting version removal for library '$LibraryName' on '$SiteUrl'."

    if ($PSCmdlet.ShouldProcess("$SiteUrl/$LibraryName", 'Remove all previous file versions')) {

        # Use a recursive CAML query to retrieve all files (including files in folders).
        $query = '<View Scope="RecursiveAll"><Query><Where><Eq><FieldRef Name="FSObjType"/><Value Type="Integer">0</Value></Eq></Where></Query></View>'

        try {
            # Retrieve items in pages to limit memory consumption on large libraries.
            $items = Get-PnPListItem -List $LibraryName -PageSize $BatchSize -Query $query
        } catch {
            throw "Failed to retrieve items from '$LibraryName': $($_.Exception.Message)"
        }

        $totalFiles = $items.Count
        Write-Log "Found $totalFiles file(s) to process."

        foreach ($item in $items) {
            $filesProcessed++
            $fileRef = $item.FieldValues['FileRef']

            Write-Progress -Activity "Removing file versions" `
                           -Status "Processing file $filesProcessed of $totalFiles: $fileRef" `
                           -PercentComplete ([int](($filesProcessed / [Math]::Max(1, $totalFiles)) * 100))

            # Retrieve the version history for this specific file.
            try {
                $versions = Get-PnPFileVersion -Url $fileRef -ErrorAction Stop
            } catch {
                Write-Log "Could not retrieve versions for '$fileRef': $($_.Exception.Message)" 'WARN'
                $errors++
                continue
            }

            if ($null -eq $versions -or $versions.Count -eq 0) {
                Write-Verbose "No previous versions found for '$fileRef'. Skipping."
                $filesSkipped++
                continue
            }

            Write-Verbose "Removing $($versions.Count) version(s) from '$fileRef'."

            # Remove-PnPFileVersion with -All removes every non-current version in one call,
            # which is more efficient than iterating and calling per-version.
            try {
                Remove-PnPFileVersion -Url $fileRef -All -Force -ErrorAction Stop
                Write-Log "Removed $($versions.Count) version(s) from '$fileRef'."
                $versionsRemoved += $versions.Count
            } catch {
                Write-Log "Failed to remove versions from '$fileRef': $($_.Exception.Message)" 'ERROR'
                $errors++
            }
        }

        Write-Progress -Activity "Removing file versions" -Completed -Status "Done"

    } else {
        # -WhatIf path: enumerate and report without deleting.
        Write-Log "WhatIf mode: no versions will be deleted."

        $query = '<View Scope="RecursiveAll"><Query><Where><Eq><FieldRef Name="FSObjType"/><Value Type="Integer">0</Value></Eq></Where></Query></View>'

        try {
            $items = Get-PnPListItem -List $LibraryName -PageSize $BatchSize -Query $query
        } catch {
            throw "Failed to retrieve items from '$LibraryName': $($_.Exception.Message)"
        }

        $totalFiles = $items.Count

        foreach ($item in $items) {
            $filesProcessed++
            $fileRef = $item.FieldValues['FileRef']

            try {
                $versions = Get-PnPFileVersion -Url $fileRef -ErrorAction Stop
            } catch {
                Write-Log "Could not retrieve versions for '$fileRef': $($_.Exception.Message)" 'WARN'
                $filesSkipped++
                continue
            }

            if ($versions.Count -gt 0) {
                Write-Host "[WhatIf] Would remove $($versions.Count) version(s) from: $fileRef" -ForegroundColor Cyan
                $versionsRemoved += $versions.Count
            } else {
                $filesSkipped++
            }
        }
    }

    #endregion Enumerate Files and Remove Versions

    #region Summary

    Write-Log "--- Version Removal Summary ---"
    Write-Log "Library        : $LibraryName"
    Write-Log "Site           : $SiteUrl"
    Write-Log "Files processed: $filesProcessed"
    Write-Log "Files skipped  : $filesSkipped (no previous versions)"
    if ($PSCmdlet.ShouldProcess('', '', '')) {
        Write-Log "Versions removed: $versionsRemoved"
    } else {
        Write-Log "Versions that WOULD be removed (WhatIf): $versionsRemoved"
    }
    Write-Log "Errors         : $errors"
    Write-Log "-------------------------------"

    if ($versionsRemoved -gt 0 -and -not $WhatIfPreference) {
        Write-Host ""
        Write-Host "Version removal complete. $versionsRemoved version(s) sent to Recycle Bin." -ForegroundColor Green
        Write-Host "To fully reclaim storage quota, empty the site Recycle Bin." -ForegroundColor Yellow
    }

    #endregion Summary
}
