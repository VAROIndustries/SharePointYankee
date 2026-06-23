<#
.SYNOPSIS
    Copies item-level permissions from a source SharePoint Online document library
    in one tenant to a destination document library in another.

.DESCRIPTION
    This script uses PnP.PowerShell to read permissions from a single source document
    library in a source site and apply equivalent permissions to a single destination
    document library in a destination site.

    KEY FEATURES:
    - READ-ONLY ON SOURCE:
        The script does not modify the source site or library. All write
        operations happen only on the destination.

    - SINGLE LIBRARY SCOPE:
        One source site + library to one destination site + library.

    - PRINCIPAL MAPPING (USING CSV):
        Uses a ShareGate-style mapping CSV to translate principals (users and
        groups) from source to destination.

    - SHAREPOINT GROUP HANDLING:
        For SharePoint groups (excluding "SharingLinks" groups), the script will
        prompt whether to create them on the destination if they do not exist.

    - CHECKPOINT / RESUME:
        Maintains a JSON checkpoint file so that if the script is interrupted,
        it can resume from where it left off on the next run.

    - LOGGING:
        Writes multiple log files in the same directory as the script:
          * PermissionsCopy.log
              - High-level status, info, and error messages.
          * PermissionsCopyErrors.log
              - Only messages tagged as "ERROR".
          * PermissionsCopyTranscript_yyyyMMdd_HHmmss.log
              - Full PowerShell transcript for troubleshooting.
          * PermissionsCopyCheckpoint.json
              - JSON checkpoint file used to track the last processed item index.

    - PROGRESS BAR + ETA:
        Displays a progress bar showing:
          * Items processed vs total
          * Percentage complete
          * Elapsed time
          * Estimated time remaining (ETA) for the current run

    IMPORTANT ASSUMPTIONS:
    - Relative paths (FileRef) under the specified library are the same between
      source and destination. If your destination structure differs, you will
      need to adjust the lookup logic where the destination item is retrieved.

    - The destination site has compatible role definitions (e.g., "Read",
      "Contribute", "Edit", "Full Control"). If custom role definition names
      differ between tenants, you will need a role mapping layer.

    - The script resets item-level permissions on each destination item and then
      applies the source item's permissions explicitly. If you need different
      behavior (e.g., only add missing principals), you must adapt that logic.

.PARAMETER SourceSiteUrl
    URL of the source SharePoint site.
    Example:
        https://contoso.sharepoint.com/sites/SourceSite

.PARAMETER DestinationSiteUrl
    URL of the destination SharePoint site (can be a different tenant or cloud environment).
    Example:
        https://contoso.sharepoint.us/sites/DestSite

.PARAMETER LibraryName
    Name (title) of the document library to process in both source and destination.
    Examples:
        "Documents"
        "Shared Documents"

.PARAMETER MappingCsvPath
    Full or relative path to the user/group mapping CSV file.
    The CSV should have at least the following columns:
        SourcePrincipal,TargetPrincipal,Type

    Example:
        SourcePrincipal,TargetPrincipal,Type
        user1@contoso.com,user1@contoso.us,User
        sg-Finance-Contrib@contoso.com,sg-Finance-Contrib@contoso.us,Group

    Notes:
    - SourcePrincipal:
        The principal identifier as used in the source tenant (usually UPN or
        login name). This should match what appears in SharePoint for permissions.
    - TargetPrincipal:
        The corresponding principal identifier in the destination tenant.
        For users: typically the new UPN.
        For groups: typically the group's display name or mail address, depending
        on how you want to reference them in PnP.
    - Type:
        Optional informational field ("User", "Group", etc.). Not used by the
        script logic but useful for documentation.

.PARAMETER Restart
    Optional switch. If specified, the script deletes any existing checkpoint file
    (PermissionsCopyCheckpoint.json) and starts processing from the beginning.

.PARAMETER MaxItems
    Optional integer. If specified, limits the script to process only this many items
    (useful for testing on a subset of the library). If not specified or set to 0,
    all items in the library are processed. The checkpoint logic still applies, so if
    the script is interrupted and resumed, it will continue from the last checkpoint
    and still respect the MaxItems limit.
    Example: -MaxItems 10 will process only the first 10 items (or resume and process
    until 10 total items have been completed).

.PARAMETER MaxRetries
    Optional integer. The maximum number of times to retry processing an item if it fails.
    If an item fails more than this number of times, it will be skipped (or prompt if
    -InteractiveSkip is used). Default is 3. Set to 0 to disable retries.

.PARAMETER InteractiveSkip
    Optional switch. When an item fails after MaxRetries attempts, if this switch is
    specified, the script will prompt the user to choose: Skip the item, Retry more times,
    or Abort the script. If not specified, the item is automatically skipped.

.PARAMETER BatchSize
    Optional integer. The page size used when retrieving items from the source library.
    Larger values can improve performance for large libraries but may increase memory usage.
    Default is 500. Valid range: 1-5000.

.EXAMPLE
    PS C:\Scripts> .\Copy-SPPermissionsCrossTenant.ps1 `
        -SourceSiteUrl "https://contoso.sharepoint.com/sites/SourceSite" `
        -DestinationSiteUrl "https://contoso.sharepoint.us/sites/DestSite" `
        -LibraryName "Documents" `
        -MappingCsvPath ".\UserMapping.csv"

    Description:
        Runs the script using the specified source/destination sites and library,
        and a mapping CSV called UserMapping.csv in the current directory.
        If a checkpoint file exists, processing resumes after the last processed item.

.EXAMPLE
    PS C:\Scripts> .\Copy-SPPermissionsCrossTenant.ps1 `
        -SourceSiteUrl "https://contoso.sharepoint.com/sites/SourceSite" `
        -DestinationSiteUrl "https://contoso.sharepoint.us/sites/DestSite" `
        -LibraryName "Documents" `
        -MappingCsvPath ".\UserMapping.csv" `
        -Restart

    Description:
        Same as above, but explicitly discards any previous checkpoint and starts
        over from the first item in the library.

.EXAMPLE
    PS C:\Scripts> .\Copy-SPPermissionsCrossTenant.ps1 `
        -SourceSiteUrl "https://contoso.sharepoint.com/sites/SourceSite" `
        -DestinationSiteUrl "https://contoso.sharepoint.us/sites/DestSite" `
        -LibraryName "Documents" `
        -MappingCsvPath ".\UserMapping.csv" `
        -MaxItems 10

    Description:
        Test run: processes only the first 10 items. Useful for validating the
        script behavior on a small subset before running against the full library.
        If interrupted and rerun, resumes from the checkpoint and continues until
        all 10 items have been processed (based on when each was first started).

.EXAMPLE
    PS C:\Scripts> .\Copy-SPPermissionsCrossTenant.ps1 `
        -SourceSiteUrl "https://contoso.sharepoint.com/sites/SourceSite" `
        -DestinationSiteUrl "https://contoso.sharepoint.us/sites/DestSite" `
        -LibraryName "Documents" `
        -MappingCsvPath ".\UserMapping.csv" `
        -MaxRetries 5 `
        -InteractiveSkip `
        -BatchSize 1000

    Description:
        Runs the script with custom retry settings (5 retries, interactive skip on failure),
        and a larger batch size for better performance on large libraries.

.NOTES
    AUTHOR:
        Geoff Varosky

    VERSION:
        1.2.0

    LAST UPDATED:
        2026-03-23

    REQUIREMENTS:
        - PowerShell 5.1 or 7.x
        - PnP.PowerShell module (Install-Module PnP.PowerShell -Scope CurrentUser)
        - Site Owner or Site Collection Administrator rights on both source and destination

    LOGGING:
        The script creates the following files in the folder where it resides:

        - PermissionsCopy.log
            * Contains informational and status messages, including progress, mapping
              notes, and error summaries.

        - PermissionsCopyErrors.log
            * Contains only lines logged with Level "ERROR". Useful for quickly
              reviewing issues without scanning the full log.

        - PermissionsCopyTranscript_yyyyMMdd_HHmmss.log
            * A full transcript generated by Start-Transcript, including all commands
              and console output. Helpful for deep troubleshooting.

        - PermissionsCopyCheckpoint.json
            * A JSON file with a single property "LastProcessedIndex" that stores
              the index of the last processed item. This enables resumability.

    CHECKPOINT / RESUME BEHAVIOR:
        - On first run (no checkpoint file present):
            * The script initializes LastProcessedIndex to -1.
            * Processing starts from item index 0.

        - After processing each item:
            * The script updates the checkpoint file with the current item index.
            * If the script is interrupted, you can rerun it without -Restart and
              it will skip items up to the last stored index and continue.

        - Using -Restart:
            * Deletes the existing checkpoint file (if any).
            * Starts processing from the first item again (index 0).

    INSTALLATION / PRE-REQUISITES:

        1. PowerShell:
           - PowerShell 5.1 (Windows) or PowerShell 7.x recommended.
           - Run your PowerShell session with sufficient rights to install modules
             and run scripts (if Execution Policy is restrictive, you may need to
             adjust it: e.g., Set-ExecutionPolicy RemoteSigned -Scope CurrentUser).

        2. PnP.PowerShell Module:
           - Install PnP.PowerShell (once per machine or per user):
                Install-Module PnP.PowerShell -Scope CurrentUser
           - Importing is automatic when using the cmdlets, but you can explicitly
             run: Import-Module PnP.PowerShell

        3. Permissions (Accounts / Access Rights):
           - Source site:
             * The account must have at least READ permissions on the source library
               and its items (including item-level permissions).

           - Destination site:
             * The account must have enough rights to manage permissions on the
               destination document library and its items (e.g., Site Owner,
               Full Control, or Site Collection Admin).
             * To create SharePoint groups (when prompted), the account must have
               rights to create and manage groups on the destination site.

        4. Authentication:
           - The script uses Connect-PnPOnline -Interactive for both source and
             destination connections. You will be prompted to sign in for each.
           - MFA or Conditional Access may apply depending on tenant policies.

        5. Mapping CSV:
           - You must supply a mapping CSV file that maps principals (users/groups)
             from source to destination.
           - This may be derived from a ShareGate mapping CSV.
           - Minimal required columns:
                SourcePrincipal,TargetPrincipal
           - An additional "Type" column is recommended for clarity but not required
             by the script logic.
           - For SharePoint groups, use the group display name (Title) as SourcePrincipal
             to enable mapping of group names between tenants.

        6. Library / Path Assumptions:
           - The script assumes that for every item (folder/file) in the source
             library, there is a corresponding item with the same FileRef path in
             the destination library.
           - If your destination library has a different structure or URL path,
             you will need to modify the section that retrieves $destItem based
             on $relativeUrl.

    LIMITATIONS / CAVEATS:
        - This script processes item-level permissions one item at a time, which
          can be slow for very large libraries. It now processes items in batches
          to reduce memory usage, loading only one page of items at a time.
        - If you want to limit scope (e.g., only folders, only items with unique
          permissions), additional filtering logic is required.
        - Role definition names must exist on the destination. If a source role
          name is missing on the destination, the permission assignment may fail.
        - This is provided as a template and may require adjustments for your
          environment. Always test on a non-production library first.

    GitHub: https://github.com/VAROIndustries/SharePointYankee

#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string][ValidatePattern('^https://')]$SourceSiteUrl,

    [Parameter(Mandatory = $true)]
    [string][ValidatePattern('^https://')]$DestinationSiteUrl,

    [Parameter(Mandatory = $true)]
    [string][ValidateNotNullOrEmpty()]$LibraryName,

    [Parameter(Mandatory = $true)]
    [string][ValidateScript({Test-Path $_ -PathType Leaf})]$MappingCsvPath,

    [switch]$Restart,

    [Parameter(Mandatory = $false)]
    [int][ValidateRange(0, [int]::MaxValue)]$MaxItems = 0,  # 0 means process all items

    [Parameter(Mandatory = $false)]
    [int][ValidateRange(0, 10)]$MaxRetries = 3,

    [Parameter(Mandatory = $false)]
    [switch]$InteractiveSkip,

    [Parameter(Mandatory = $false)]
    [int][ValidateRange(1, 5000)]$BatchSize = 500
)

#region Initial Setup and Logging

# Determine the directory where the script is located.
# All logs and the checkpoint file will be stored here.
$ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

# Log file paths
$LogFile      = Join-Path $ScriptRoot "PermissionsCopy.log"
$ErrorLogFile = Join-Path $ScriptRoot "PermissionsCopyErrors.log"
$CheckpointFile = Join-Path $ScriptRoot "PermissionsCopyCheckpoint.json"

# Logging function for informational and error messages.
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"  # "INFO" or "ERROR"
    )

    $timestamp = (Get-Date).ToString("s")
    $line = "[$timestamp] [$Level] $Message"

    # Write to console
    Write-Host $line

    # Append to main log file
    Add-Content -Path $LogFile -Value $line

    # If this is an error, also append to the error log
    if ($Level -eq "ERROR") {
        Add-Content -Path $ErrorLogFile -Value $line
    }
}

# Start a PowerShell transcript for detailed troubleshooting.
$TranscriptFile = Join-Path $ScriptRoot "PermissionsCopyTranscript_$((Get-Date).ToString('yyyyMMdd_HHmmss')).log"
Start-Transcript -Path $TranscriptFile -Force | Out-Null

Write-Log "Script started."

# Handle the Restart switch: if specified, delete any existing checkpoint file.
if ($Restart -and (Test-Path $CheckpointFile)) {
    Write-Log "Restart switch specified. Removing existing checkpoint file '$CheckpointFile'."
    Remove-Item $CheckpointFile -Force
}

#endregion Initial Setup and Logging

#region Load Mapping CSV

# Ensure the mapping CSV exists.
if (-not (Test-Path $MappingCsvPath)) {
    Write-Log "Mapping CSV not found at path: $MappingCsvPath" "ERROR"
    Stop-Transcript | Out-Null
    throw "Mapping CSV missing. Please provide a valid MappingCsvPath."
}

Write-Log "Loading mapping CSV from '$MappingCsvPath'."

try {
    $mappingRows = Import-Csv -Path $MappingCsvPath
} catch {
    Write-Log "Failed to read mapping CSV: $($_.Exception.Message)" "ERROR"
    Stop-Transcript | Out-Null
    throw
}

# Build a hashtable for quick lookup:
# Key   = SourcePrincipal (lowercase, trimmed)
# Value = TargetPrincipal (as-is, trimmed)
$PrincipalMap = @{}
foreach ($row in $mappingRows) {
    if ($row.SourcePrincipal -and $row.TargetPrincipal) {
        $key = $row.SourcePrincipal.Trim().ToLower()
        $PrincipalMap[$key] = $row.TargetPrincipal.Trim()
    }
}

Write-Log "Loaded $($PrincipalMap.Count) principal mappings from CSV."

#endregion Load Mapping CSV

#region Connect to Source and Destination Sites

# CONNECT TO SOURCE SITE (READ-ONLY)
Write-Log "Connecting to source site: $SourceSiteUrl"
try {
    Connect-PnPOnline -Url $SourceSiteUrl -Interactive
    # Capture the source connection object so we can switch to it later without re-authenticating.
    $srcConnection = Get-PnPConnection
} catch {
    Write-Log "Failed to connect to source site '$SourceSiteUrl': $($_.Exception.Message)" "ERROR"
    Stop-Transcript | Out-Null
    throw
}

# CONNECT TO DESTINATION SITE (WRITE)
Write-Log "Connecting to destination site: $DestinationSiteUrl"
try {
    Connect-PnPOnline -Url $DestinationSiteUrl -Interactive -ReturnConnection:$true
    # Capture the destination connection object so we can switch to it later.
    $destConnection = Get-PnPConnection
} catch {
    Write-Log "Failed to connect to destination site '$DestinationSiteUrl': $($_.Exception.Message)" "ERROR"
    Stop-Transcript | Out-Null
    throw
}

#endregion Connect to Source and Destination Sites

#region Retrieve Source and Destination Libraries

# Retrieve the source document library by name.
# IMPORTANT: The script does not modify the source list or items.
Write-Log "Retrieving source library '$LibraryName'."
try {
    $srcList = Get-PnPList -Identity $LibraryName
} catch {
    Write-Log "Source library '$LibraryName' not found on '$SourceSiteUrl': $($_.Exception.Message)" "ERROR"
    Stop-Transcript | Out-Null
    throw
}

# Get total item count for progress and limits
$totalItems = $srcList.ItemCount

# Apply MaxItems limit if specified (for testing purposes)
if ($MaxItems -gt 0 -and $MaxItems -lt $totalItems) {
    $totalItems = $MaxItems
    Write-Log "MaxItems limit applied: processing only $totalItems items (out of $($srcList.ItemCount) total)." "INFO"
}

# Switch connection to destination to retrieve the destination library.
Write-Log "Switching to destination connection to retrieve destination library '$LibraryName'."
Set-PnPConnection -Connection $destConnection

try {
    $destList = Get-PnPList -Identity $LibraryName
} catch {
    Write-Log "Destination library '$LibraryName' not found on '$DestinationSiteUrl': $($_.Exception.Message)" "ERROR"
    Stop-Transcript | Out-Null
    throw
}

# Retrieve destination role definitions for validation
$destRoleNames = Get-PnPRoleDefinition | Select-Object -ExpandProperty Name
Write-Log "Retrieved $($destRoleNames.Count) role definitions from destination."

# Switch back to source connection for reading items.
Write-Log "Switching back to source connection to read items."

#endregion Retrieve Source and Destination Libraries

#region Checkpoint Handling

# Initialize checkpoint structure.
# LastProcessedIndex = index of last item processed.
$checkpoint = @{
    LastProcessedIndex = -1
}

# If a checkpoint file exists, load it and resume from that index.
if (Test-Path $CheckpointFile) {
    Write-Log "Checkpoint file found at '$CheckpointFile'. Loading checkpoint."
    try {
        $checkpoint = Get-Content -Raw -Path $CheckpointFile | ConvertFrom-Json
        Write-Log "Last processed index from checkpoint: $($checkpoint.LastProcessedIndex)"
    } catch {
        Write-Log "Failed to load checkpoint file. Starting from beginning. Error: $($_.Exception.Message)" "ERROR"
        $checkpoint = @{ LastProcessedIndex = -1 }
    }
} else {
    Write-Log "No checkpoint file found. Starting from the first item."
}

# Helper function to save checkpoint.
function Save-Checkpoint {
    param(
        [int]$Index
    )
    $checkpoint.LastProcessedIndex = $Index
    $checkpoint | ConvertTo-Json | Set-Content -Path $CheckpointFile -Encoding UTF8
}

# Set up batch processing variables
$currentGlobalIndex   = $checkpoint.LastProcessedIndex + 1
$currentPage          = [math]::Floor($currentGlobalIndex / $BatchSize)
$currentIndexInPage   = $currentGlobalIndex % $BatchSize
$itemsProcessed       = $currentGlobalIndex

#endregion Checkpoint Handling

#region Helper Functions

<#
.SYNOPSIS
    Maps a source principal to a destination principal using the mapping CSV.

.DESCRIPTION
    Looks up the provided source principal in the $PrincipalMap hashtable.
    If a mapping exists, returns the mapped value.
    If no mapping exists, returns the original source principal (and logs info).

.PARAMETER SourcePrincipal
    The principal identifier from the source tenant (e.g., user UPN).

.RETURNS
    The mapped destination principal, or the original source principal if not found.
#>
function Map-Principal {
    param(
        [string]$SourcePrincipal
    )

    if ([string]::IsNullOrWhiteSpace($SourcePrincipal)) {
        return $null
    }

    # We normalize by trimming and lowercasing the key for lookup.
    $key = $SourcePrincipal.Trim().ToLower()

    if ($PrincipalMap.ContainsKey($key)) {
        return $PrincipalMap[$key]
    } else {
        # If no mapping is found, we log this event but still return the original.
        Write-Log "No mapping found for principal '$SourcePrincipal'; leaving unchanged." "INFO"
        return $SourcePrincipal
    }
}

<#
.SYNOPSIS
    Ensures that a SharePoint group exists on the destination site.

.DESCRIPTION
    For a given source SharePoint group name, this function:
    - Skips groups with names containing "SharingLinks" (these are typically
      system-generated sharing link groups).
    - Checks whether a group with the same Title exists on the destination.
    - If it exists, returns the group.
    - If it does not exist, prompts the user to create it.
      If the user chooses "Y", creates the group and returns it.
      If the user chooses "N", returns $null.

.PARAMETER SourceGroupName
    The Title of the SharePoint group in the source tenant.

.RETURNS
    The destination group object if created or found, otherwise $null.
#>
function Ensure-SPGroupOnDestination {
    param(
        [string]$SourceGroupName
    )

    # Exclude "SharingLinks" groups from being created or migrated.
    if ($SourceGroupName -like "*SharingLinks*") {
        Write-Log "Skipping SharingLinks group '$SourceGroupName'."
        return $null
    }

    # Switch to destination connection before getting/creating groups.
    Set-PnPConnection -Connection $destConnection

    # Check if the group already exists.
    $existingGroup = Get-PnPGroup | Where-Object { $_.Title -eq $SourceGroupName }
    if ($existingGroup) {
        Write-Log "Destination SharePoint group '$SourceGroupName' already exists."
        return $existingGroup
    }

    # If not, ask the user whether to create it.
    Write-Host ""
    Write-Host "SharePoint group '$SourceGroupName' does not exist on the destination."
    $answer = Read-Host "Create it on destination? (Y/N)"
    if ($answer -match '^[Yy]') {
        try {
            $newGroup = New-PnPGroup -Title $SourceGroupName -Description "Created by permissions copy script"
            Write-Log "Created destination SharePoint group '$SourceGroupName'."
            return $newGroup
        } catch {
            Write-Log "Failed to create group '$SourceGroupName': $($_.Exception.Message)" "ERROR"
            return $null
        }
    } else {
        Write-Log "User declined creation of SharePoint group '$SourceGroupName' on destination."
        return $null
    }
}

#endregion Helper Functions

<#
.SYNOPSIS
    Processes a single item: copies its permissions to destination.
#>
function Process-Item {
    param(
        [object]$srcItem,
        [int]$index
    )

    $itemUrl = $srcItem.FieldValues.FileRef

    Write-Log "[$($index+1)/$totalItems] Processing source item: $itemUrl"

    $relativeUrl = $itemUrl

    # Look up the corresponding destination item by FileRef (server-relative URL).
    $queryXml = @"
<View>
  <Query>
    <Where>
      <Eq>
        <FieldRef Name='FileRef' />
        <Value Type='Text'>$relativeUrl</Value>
      </Eq>
    </Where>
  </Query>
</View>
"@

    $destItem = Get-PnPListItem -List $destList -Query $queryXml -PageSize 1 | Select-Object -First 1

    if (-not $destItem) {
        throw "Destination item not found for '$relativeUrl'."
    }

    # Switch to source connection to read permissions from the source item.
    Set-PnPConnection -Connection $srcConnection

    $srcRoleAssignments = Get-PnPProperty -ClientObject $srcItem -Property RoleAssignments

    # Switch back to destination connection for permission updates.
    Set-PnPConnection -Connection $destConnection

    # Reset role inheritance on the destination item so we can set explicit permissions.
    Set-PnPListItemPermission -List $destList -Identity $destItem.Id -InheritPermissions:$false

    # Iterate over each role assignment on the source item and attempt to recreate it on the destination item.
    foreach ($ra in $srcRoleAssignments) {
        $principal = Get-PnPProperty -ClientObject $ra -Property Member
        $roleDefs  = Get-PnPProperty -ClientObject $ra -Property RoleDefinitionBindings

        $principalName = $principal.LoginName

        # Skip system accounts if needed (optional).
        if ($principalName -like "SHAREPOINT\system") {
            Write-Log "Skipping system principal '$principalName'."
            continue
        }

        # Map principal via the mapping CSV.
        $mappedPrincipal = Map-Principal -SourcePrincipal $principalName

        if (-not $mappedPrincipal) {
            Write-Log "Mapped principal is null/empty for '$principalName'. Skipping this assignment."
            continue
        }

        # Determine whether the principal is a SharePoint group or a user/AAD group.
        if ($principal.PrincipalType -eq "SharePointGroup" -or $principalName -like "*|*") {
            # SharePoint group: use the group Title as the name.
            $groupName       = $principal.Title
            $mappedGroupName = Map-Principal -SourcePrincipal $groupName
            if (-not $mappedGroupName) {
                Write-Log "Mapped group name is null for '$groupName'. Skipping this assignment."
                continue
            }
            $spGroup = Ensure-SPGroupOnDestination -SourceGroupName $mappedGroupName
            if (-not $spGroup) {
                Write-Log "No destination SharePoint group available for '$mappedGroupName'; skipping this assignment."
                continue
            }
            $destPrincipal = $spGroup.Title
        } else {
            # For users / Azure AD groups, mappedPrincipal is the value we use (UPN or mail).
            $destPrincipal = $mappedPrincipal
        }

        # For each role definition (e.g., Read, Contribute) associated with this assignment, apply it on destination.
        foreach ($roleDef in $roleDefs) {
            $roleName = $roleDef.Name

            if ($roleName -notin $destRoleNames) {
                Write-Log "Role '$roleName' not found on destination. Skipping assignment." "ERROR"
                continue
            }

            Set-PnPListItemPermission -List $destList -Identity $destItem.Id -User $destPrincipal -AddRole $roleName
            Write-Log "Assigned role '$roleName' to '$destPrincipal' on '$relativeUrl'."
        }
    }
}

#region Process Items and Copy Permissions (with Progress Bar + ETA)

# Switch to destination connection for write operations.
Set-PnPConnection -Connection $destConnection

Write-Log "Starting permission copy for $totalItems items."

# Track start time for progress/ETA calculation.
$startTime = Get-Date

# Process items in batches to reduce memory usage
while ($itemsProcessed -lt $totalItems) {
    try {
        $items = Get-PnPListItem -List $srcList -PageSize $BatchSize -ScriptBlock { param($items) $items } -Page $currentPage
    } catch {
        Write-Log "Failed to retrieve page $currentPage from source library '$LibraryName': $($_.Exception.Message)" "ERROR"
        Stop-Transcript | Out-Null
        throw
    }

    if ($items.Count -eq 0) {
        Write-Log "No more items to process. Ending." "INFO"
        break
    }

    for ($i = $currentIndexInPage; $i -lt $items.Count; $i++) {
        $item        = $items[$i]
        $globalIndex = $currentPage * $BatchSize + $i

        if ($globalIndex -ge $totalItems) {
            break
        }

        $itemUrl = $item.FieldValues.FileRef

        # Progress metrics.
        $processedCount  = $globalIndex + 1
        $percentComplete = [int](($processedCount / $totalItems) * 100)

        $elapsed        = (Get-Date) - $startTime
        $elapsedSeconds = [math]::Max(1, $elapsed.TotalSeconds)  # Avoid division by zero.
        $secPerItem     = $elapsedSeconds / $processedCount
        $remainingItems = $totalItems - $processedCount
        $etaSeconds     = $secPerItem * $remainingItems
        $eta            = [TimeSpan]::FromSeconds($etaSeconds)

        $statusMessage = "Processing item $processedCount of $totalItems"
        $etaMessage    = "Elapsed: {0:hh\:mm\:ss} | Remaining: {1:hh\:mm\:ss}" -f $elapsed, $eta

        # Display progress bar with ETA.
        Write-Progress -Activity "Copying permissions" `
                       -Status "$statusMessage - $etaMessage" `
                       -PercentComplete $percentComplete

        $retryCount = 0
        while ($retryCount -le $MaxRetries) {
            try {
                Process-Item -srcItem $item -index $globalIndex
                Save-Checkpoint -Index $globalIndex
                $itemsProcessed++
                break
            } catch {
                $retryCount++
                Write-Log "Error processing item '$itemUrl': $($_.Exception.Message)" "ERROR"
                if ($retryCount -gt $MaxRetries) {
                    if ($InteractiveSkip) {
                        $answer = Read-Host "Item '$itemUrl' failed $MaxRetries times. Skip (S), Retry more (R), or Abort (A)?"
                        if ($answer -match '^[Ss]') {
                            Write-Log "Skipping item '$itemUrl' after $MaxRetries failures."
                            $itemsProcessed++
                            break
                        } elseif ($answer -match '^[Rr]') {
                            $retryCount = 0
                        } else {
                            throw "Aborting due to repeated failures on item '$itemUrl'."
                        }
                    } else {
                        Write-Log "Skipping item '$itemUrl' after $MaxRetries failures." "ERROR"
                        $itemsProcessed++
                        break
                    }
                } else {
                    Write-Log "Retrying item '$itemUrl' (attempt $retryCount/$MaxRetries)."
                }
            }
        }
    }

    $currentPage++
    $currentIndexInPage = 0
}

# Mark progress as completed.
Write-Progress -Activity "Copying permissions" -Completed -Status "Done"

Write-Log "Permission copy completed for all items (or up to the last available item)."

#endregion Process Items and Copy Permissions

# Stop the PowerShell transcript.
Stop-Transcript | Out-Null
Write-Log "Script finished."
