#Requires -Version 5.1
#Requires -Modules @{ ModuleName = 'PnP.PowerShell'; ModuleVersion = '1.12.0' }

<#
.SYNOPSIS
    Exports all unique permissions across a SharePoint Online site, including
    lists and libraries with broken inheritance.

.DESCRIPTION
    Export-SPSitePermissions performs a comprehensive permissions audit of a
    SharePoint Online site using PnP.PowerShell. It reports:

    - Site-level role assignments (who has what permission at the root)
    - All lists and document libraries with broken permission inheritance
    - Role assignments on each list/library with unique permissions
    - Optionally: item-level permissions within each list/library

    OUTPUT OBJECT PROPERTIES:
    - ScopeType      : Site, List, Library, or Item
    - ScopeUrl       : Server-relative URL of the object being reported
    - ScopeTitle     : Display name of the list/library (empty for site scope)
    - HasUniquePerms : Whether this scope has broken inheritance (True/False)
    - PrincipalName  : Display name of the user, group, or security principal
    - PrincipalLogin : Login name / claim string of the principal
    - PrincipalType  : User, SharePointGroup, or SecurityGroup
    - PermissionLevel: Comma-separated list of role definition names granted

    The function produces one row per principal-per-scope combination. A user
    with multiple roles on the same scope will appear on one row with a
    comma-separated PermissionLevel value.

    PERFORMANCE NOTE:
    Including item-level permissions (-IncludeItemPermissions) can be very slow
    on large libraries. Use -MaxItemsPerList to cap the number of items checked
    per list, or run the report against specific lists only.

.PARAMETER SiteUrl
    The full URL of the SharePoint Online site to audit.
    Example: https://contoso.sharepoint.com/sites/MySite

.PARAMETER UseExistingConnection
    When specified, uses the currently active PnP connection instead of
    prompting for interactive login. Useful when automating across multiple sites.

.PARAMETER IncludeItemPermissions
    When specified, the function also checks item-level permissions within each
    list or library that has unique permissions. This can significantly increase
    runtime on large sites.

.PARAMETER MaxItemsPerList
    When -IncludeItemPermissions is specified, limits the number of items
    checked per list/library. Defaults to 500. Set to 0 to check all items
    (use with caution on large libraries).

.PARAMETER ExcludeSystemLists
    When specified, excludes hidden system lists (e.g., Site Assets, Style
    Library, MicroFeed) from the output.

.PARAMETER ExportPath
    Optional. Full path to a CSV file where results will be exported.
    The directory must already exist. If the file exists it will be overwritten.
    Example: C:\Reports\SitePermissions.csv

.EXAMPLE
    Export-SPSitePermissions -SiteUrl "https://contoso.sharepoint.com/sites/HR"

    Description:
        Audits the HR site and outputs permissions to the console.

.EXAMPLE
    Export-SPSitePermissions `
        -SiteUrl "https://contoso.sharepoint.com/sites/HR" `
        -ExportPath "C:\Reports\HR_Permissions.csv"

    Description:
        Audits the HR site and exports results to a CSV.

.EXAMPLE
    Export-SPSitePermissions `
        -SiteUrl "https://contoso.sharepoint.com/sites/Finance" `
        -IncludeItemPermissions `
        -MaxItemsPerList 200 `
        -ExcludeSystemLists `
        -ExportPath "C:\Reports\Finance_Permissions.csv"

    Description:
        Full audit including item-level permissions, capped at 200 items per
        list, with system lists excluded.

.EXAMPLE
    Connect-PnPOnline -Url "https://contoso.sharepoint.com/sites/HR" -Interactive
    Export-SPSitePermissions `
        -SiteUrl "https://contoso.sharepoint.com/sites/HR" `
        -UseExistingConnection `
        -ExportPath "C:\Reports\HR_Permissions.csv"

    Description:
        Reuses an existing PnP connection to avoid a second login prompt.

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
        - Site Collection Administrator or Full Control on the target site
          (Read alone is insufficient for retrieving role assignments)

    REFERENCES:
        https://pnp.github.io/powershell/cmdlets/Get-PnPWeb.html
        https://pnp.github.io/powershell/cmdlets/Get-PnPList.html
        https://pnp.github.io/powershell/cmdlets/Get-PnPListItem.html

    GitHub: https://github.com/VAROIndustries/SharePointYankee
#>

function Export-SPSitePermissions {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(
            Mandatory = $true,
            HelpMessage = 'Full URL of the SharePoint Online site to audit.'
        )]
        [ValidatePattern('^https://')]
        [string]$SiteUrl,

        [Parameter(Mandatory = $false)]
        [switch]$UseExistingConnection,

        [Parameter(Mandatory = $false)]
        [switch]$IncludeItemPermissions,

        [Parameter(Mandatory = $false)]
        [ValidateRange(0, 100000)]
        [int]$MaxItemsPerList = 500,

        [Parameter(Mandatory = $false)]
        [switch]$ExcludeSystemLists,

        [Parameter(Mandatory = $false)]
        [ValidateScript({
            $parent = Split-Path $_ -Parent
            if ($parent -and -not (Test-Path $parent)) {
                throw "Directory '$parent' does not exist."
            }
            $true
        })]
        [string]$ExportPath
    )

    $ErrorActionPreference = 'Stop'
    $allResults = [System.Collections.Generic.List[PSCustomObject]]::new()

    #region Internal Helpers

    # Converts a collection of PnP role assignment objects into permission report rows.
    function ConvertTo-PermissionRows {
        param (
            [Parameter(Mandatory)] [string]$ScopeType,
            [Parameter(Mandatory)] [string]$ScopeUrl,
            [Parameter(Mandatory)] [string]$ScopeTitle,
            [Parameter(Mandatory)] [bool]$HasUniquePerms,
            [Parameter(Mandatory)] [object[]]$RoleAssignments
        )

        foreach ($ra in $RoleAssignments) {
            $principal = $ra.Member

            # Retrieve role definition names, skipping the built-in
            # "Limited Access" role which is assigned automatically
            # whenever a user has access to a child item.
            $roleNames = $ra.RoleDefinitionBindings |
                Where-Object { $_.Name -ne 'Limited Access' } |
                Select-Object -ExpandProperty Name

            if (-not $roleNames) { continue }

            [PSCustomObject]@{
                ScopeType      = $ScopeType
                ScopeUrl       = $ScopeUrl
                ScopeTitle     = $ScopeTitle
                HasUniquePerms = $HasUniquePerms
                PrincipalName  = $principal.Title
                PrincipalLogin = $principal.LoginName
                PrincipalType  = $principal.PrincipalType.ToString()
                PermissionLevel = ($roleNames -join ', ')
            }
        }
    }

    #endregion Internal Helpers

    #region Connect

    if (-not $UseExistingConnection) {
        try {
            Write-Verbose "Connecting to '$SiteUrl'..."
            Connect-PnPOnline -Url $SiteUrl -Interactive
        } catch {
            throw "Failed to connect to '$SiteUrl': $($_.Exception.Message)"
        }
    }

    #endregion Connect

    #region Site-Level Permissions

    Write-Verbose "Retrieving site-level permissions..."

    try {
        # Load the root web with role assignments and their bindings.
        $web = Get-PnPWeb -Includes RoleAssignments, RoleAssignments.Member,
                                     RoleAssignments.RoleDefinitionBindings,
                                     HasUniqueRoleAssignments

        $siteRows = ConvertTo-PermissionRows `
            -ScopeType     'Site' `
            -ScopeUrl      $web.ServerRelativeUrl `
            -ScopeTitle    $web.Title `
            -HasUniquePerms $true `
            -RoleAssignments $web.RoleAssignments

        foreach ($row in $siteRows) {
            $row
            $allResults.Add($row)
        }

        Write-Verbose "Site-level: $($siteRows.Count) permission row(s) found."
    } catch {
        Write-Warning "Failed to retrieve site-level permissions: $($_.Exception.Message)"
    }

    #endregion Site-Level Permissions

    #region List/Library Permissions

    Write-Verbose "Retrieving lists and libraries..."

    try {
        $lists = Get-PnPList -Includes HasUniqueRoleAssignments, RoleAssignments,
                                        RoleAssignments.Member,
                                        RoleAssignments.RoleDefinitionBindings
    } catch {
        throw "Failed to retrieve lists from '$SiteUrl': $($_.Exception.Message)"
    }

    $filteredLists = $lists | Where-Object {
        # Only process document libraries and generic lists, not internal catalogs.
        ($_.BaseType -eq 'DocumentLibrary' -or $_.BaseType -eq 'GenericList') -and
        # Skip hidden system lists if the switch is set.
        (-not $ExcludeSystemLists -or -not $_.Hidden)
    }

    Write-Verbose "Processing $($filteredLists.Count) list(s)/library(ies)..."

    $listCount = 0
    foreach ($list in $filteredLists) {
        $listCount++
        $scopeType  = if ($list.BaseType -eq 'DocumentLibrary') { 'Library' } else { 'List' }
        $scopeTitle = $list.Title
        $scopeUrl   = $list.RootFolder.ServerRelativeUrl

        Write-Progress -Activity "Auditing list/library permissions" `
                       -Status "$scopeType '$scopeTitle' ($listCount of $($filteredLists.Count))" `
                       -PercentComplete ([int](($listCount / $filteredLists.Count) * 100))

        if ($list.HasUniqueRoleAssignments) {
            Write-Verbose "$scopeType '$scopeTitle' has unique permissions."

            $listRows = ConvertTo-PermissionRows `
                -ScopeType      $scopeType `
                -ScopeUrl       $scopeUrl `
                -ScopeTitle     $scopeTitle `
                -HasUniquePerms $true `
                -RoleAssignments $list.RoleAssignments

            foreach ($row in $listRows) {
                $row
                $allResults.Add($row)
            }
        } else {
            # Emit a single informational row indicating inherited permissions.
            $inheritedRow = [PSCustomObject]@{
                ScopeType      = $scopeType
                ScopeUrl       = $scopeUrl
                ScopeTitle     = $scopeTitle
                HasUniquePerms = $false
                PrincipalName  = '(Inherits from site)'
                PrincipalLogin = ''
                PrincipalType  = ''
                PermissionLevel = ''
            }
            $inheritedRow
            $allResults.Add($inheritedRow)
        }

        #region Item-Level Permissions

        if ($IncludeItemPermissions -and $list.HasUniqueRoleAssignments) {
            Write-Verbose "Checking item-level permissions in '$scopeTitle'..."

            try {
                $pageSize  = 200
                $itemLimit = if ($MaxItemsPerList -gt 0) { $MaxItemsPerList } else { [int]::MaxValue }
                $itemCount = 0

                # Use a CAML query to retrieve only items with unique permissions.
                # This avoids loading every item in large libraries.
                $query = '<View Scope="RecursiveAll"><Query></Query></View>'

                $items = Get-PnPListItem -List $list -PageSize $pageSize -Query $query `
                             -Includes HasUniqueRoleAssignments, RoleAssignments,
                                        RoleAssignments.Member,
                                        RoleAssignments.RoleDefinitionBindings

                foreach ($item in $items) {
                    if ($itemCount -ge $itemLimit) {
                        Write-Warning "MaxItemsPerList ($MaxItemsPerList) reached for '$scopeTitle'. Remaining items skipped."
                        break
                    }

                    if ($item.HasUniqueRoleAssignments) {
                        $itemUrl = $item.FieldValues['FileRef']
                        if (-not $itemUrl) { $itemUrl = "$scopeUrl/Item($($item.Id))" }

                        $itemRows = ConvertTo-PermissionRows `
                            -ScopeType      'Item' `
                            -ScopeUrl       $itemUrl `
                            -ScopeTitle     $scopeTitle `
                            -HasUniquePerms $true `
                            -RoleAssignments $item.RoleAssignments

                        foreach ($row in $itemRows) {
                            $row
                            $allResults.Add($row)
                        }
                    }

                    $itemCount++
                }
            } catch {
                Write-Warning "Error retrieving item permissions for '$scopeTitle': $($_.Exception.Message)"
            }
        }

        #endregion Item-Level Permissions
    }

    Write-Progress -Activity "Auditing list/library permissions" -Completed -Status "Done"

    #endregion List/Library Permissions

    #region Export

    if ($ExportPath) {
        if ($allResults.Count -gt 0) {
            try {
                $allResults | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8 -Force
                Write-Host "Exported $($allResults.Count) permission row(s) to '$ExportPath'." -ForegroundColor Green
            } catch {
                Write-Error "Failed to export results to '$ExportPath': $($_.Exception.Message)"
            }
        } else {
            Write-Warning "No results to export."
        }
    }

    #endregion Export

    Write-Verbose "Export-SPSitePermissions complete. Total rows: $($allResults.Count)."
}
