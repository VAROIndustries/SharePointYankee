param(
    [Parameter(Mandatory = $true)]
    [string]$SiteURL,

    [Parameter(Mandatory = $false)]
    [string]$TenantType = "Commercial",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\PreservationHold-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
)

# Validate tenant type
$validTenants = @("Commercial", "GCC")
if ($TenantType -notin $validTenants) {
    Write-Error "Invalid TenantType. Use 'Commercial' or 'GCC'"
    exit 1
}

Write-Host "=== SharePoint Online Preservation Hold Library Retrieval ===" -ForegroundColor Cyan
Write-Host "Site URL: $SiteURL" -ForegroundColor Yellow
Write-Host "Tenant Type: $TenantType" -ForegroundColor Yellow
Write-Host "Output File: $OutputPath" -ForegroundColor Yellow
Write-Host ""

# Connect to SharePoint Online
Write-Host "Connecting to SharePoint Online..." -ForegroundColor Cyan
try {
    Connect-PnPOnline -Url $SiteURL -UseWebLogin -WarningAction SilentlyContinue
    Write-Host "Connected successfully!" -ForegroundColor Green
}
catch {
    Write-Error "Failed to connect to SharePoint Online: $_"
    exit 1
}

$allItems     = @()
$totalSize    = 0
$totalVersionSize = 0
$totalVersionCount = 0

try {
    # Discover the Preservation Hold Library by scanning all lists (including hidden)
    Write-Host ""
    Write-Host "Discovering Preservation Hold Library..." -ForegroundColor Cyan

    $allLists = Get-PnPList -Includes BaseType, ItemCount -WarningAction SilentlyContinue
    $phl = $allLists | Where-Object { $_.Title -like "*Preservation*" -and $_.BaseType -eq "DocumentLibrary" } | Select-Object -First 1

    if ($null -eq $phl) {
        Write-Host "No Preservation Hold Library found on this site." -ForegroundColor Yellow
        Write-Host "It may not exist if no retention policies are applied." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "All libraries found on this site:" -ForegroundColor Cyan
        $allLists | Where-Object { $_.BaseType -eq "DocumentLibrary" } | ForEach-Object {
            Write-Host "  Title: '$($_.Title)'  InternalName: '$($_.InternalName)'" -ForegroundColor Gray
        }
        Disconnect-PnPOnline -WarningAction SilentlyContinue
        exit 0
    }

    Write-Host "[OK] Found: '$($phl.Title)' (internal name: '$($phl.InternalName)') — $($phl.ItemCount) items" -ForegroundColor Green
    Write-Host "Retrieving items (this may take a while for large libraries)..." -ForegroundColor Cyan

    $items = Get-PnPListItem -List $phl.Id -PageSize 5000 -WarningAction SilentlyContinue

    foreach ($item in $items) {
        $fields = $item.FieldValues

        $itemName       = $fields["FileLeafRef"]
        $itemURL        = $fields["FileRef"]
        $itemCreated    = $fields["Created"]
        $itemModified   = $fields["Modified"]
        $itemCreatedBy  = $fields["Author"]
        $itemModifiedBy = $fields["Editor"]
        $itemSize       = 0
        $versionCount   = 0
        $versionSize    = 0

        # Try both size fields
        if ($fields["File_x0020_Size"]) {
            $itemSize = [int64]$fields["File_x0020_Size"]
        }
        elseif ($fields["SMTotalFileStreamSize"]) {
            $itemSize = [int64]$fields["SMTotalFileStreamSize"]
        }

        # Preservation-specific fields
        $originalUrl    = $fields["_vti_ItemDeclaredRecord"] -or ""
        $holdDate       = $fields["_vti_ItemHoldRecordStatus"] -or ""
        $preservedFrom  = $fields["PreservationDateOfCapture"] -or ""
        $originalAuthor = $fields["PreservationOriginalAuthor"] -or ""
        $originalPath   = $fields["PreservationOriginalURL"] -or $fields["_dlc_DocId"] -or ""

        # Get version history
        if ($itemURL) {
            try {
                $versions = Get-PnPFileVersion -Url $itemURL -WarningAction SilentlyContinue
                if ($versions) {
                    $priorVersions = $versions | Select-Object -Skip 1
                    $versionCount  = @($priorVersions).Count
                    foreach ($v in $priorVersions) {
                        if ($v.Size) { $versionSize += $v.Size }
                    }
                }
            }
            catch {
                # Silently skip version errors per item
            }
        }

        $allItems += [PSCustomObject]@{
            "Item Name"            = $itemName
            "Item URL"             = $itemURL
            "Original Path"        = $originalPath
            "Original Author"      = $originalAuthor
            "Preserved Date"       = $preservedFrom
            "Size (bytes)"         = $itemSize
            "Version Count"        = $versionCount
            "Version Size (bytes)" = $versionSize
            "Created"              = $itemCreated
            "Modified"             = $itemModified
            "Created By"           = $itemCreatedBy
            "Modified By"          = $itemModifiedBy
        }

        $totalSize         += $itemSize
        $totalVersionSize  += $versionSize
        $totalVersionCount += $versionCount
    }

    Write-Host "[OK] Processed $($items.Count) items" -ForegroundColor Green
}
catch {
    Write-Error "Error retrieving Preservation Hold Library: $_"
    exit 1
}
finally {
    Disconnect-PnPOnline -WarningAction SilentlyContinue
}

# Export to CSV
Write-Host ""
Write-Host "Exporting results to CSV..." -ForegroundColor Cyan
try {
    $allItems | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8 -Force
    Write-Host "[OK] Exported $($allItems.Count) items to: $OutputPath" -ForegroundColor Green
}
catch {
    Write-Error "Failed to export CSV: $_"
    exit 1
}

# Summary
$totalSizeGB        = [math]::Round($totalSize / 1GB, 2)
$totalSizeMB        = [math]::Round($totalSize / 1MB, 2)
$totalVersionSizeGB = [math]::Round($totalVersionSize / 1GB, 2)
$totalVersionSizeMB = [math]::Round($totalVersionSize / 1MB, 2)
$combinedSize       = $totalSize + $totalVersionSize
$combinedSizeGB     = [math]::Round($combinedSize / 1GB, 2)
$combinedSizeMB     = [math]::Round($combinedSize / 1MB, 2)

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "       PRESERVATION HOLD LIBRARY SUMMARY             " -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "ITEM STATISTICS:" -ForegroundColor Yellow
Write-Host "  Total Items:       $($allItems.Count)"
Write-Host ""
Write-Host "CURRENT VERSION SIZE:" -ForegroundColor Yellow
Write-Host "  Size (GB):         $totalSizeGB GB"
Write-Host "  Size (MB):         $totalSizeMB MB"
Write-Host "  Size (bytes):      $totalSize bytes"
Write-Host ""
Write-Host "VERSION HISTORY SIZE:" -ForegroundColor Yellow
Write-Host "  Total Versions:    $totalVersionCount"
Write-Host "  Size (GB):         $totalVersionSizeGB GB"
Write-Host "  Size (MB):         $totalVersionSizeMB MB"
Write-Host "  Size (bytes):      $totalVersionSize bytes"
Write-Host ""
Write-Host "COMBINED SIZE (Current + Versions):" -ForegroundColor Yellow
Write-Host "  Size (GB):         $combinedSizeGB GB"
Write-Host "  Size (MB):         $combinedSizeMB MB"
Write-Host "  Size (bytes):      $combinedSize bytes"
Write-Host ""
Write-Host "Script completed successfully!"
Write-Host "Results saved to: $OutputPath"
Write-Host ""
