param(
    [Parameter(Mandatory = $true)]
    [string]$SiteURL,
    
    [Parameter(Mandatory = $false)]
    [string]$TenantType = "Commercial",
    
    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\SiteContents-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
)

# Tenant endpoints
$tenantConfig = @{
    "Commercial" = @{
        "SharePointUrl" = "https://$($SiteURL.Split('.')[0]).sharepoint.com"
    }
    "GCC"        = @{
        "SharePointUrl" = "https://$($SiteURL.Split('.')[0]).sharepoint.us"
    }
}

# Validate tenant type
if ($tenantConfig[$TenantType] -eq $null) {
    Write-Error "Invalid TenantType. Use 'Commercial' or 'GCC'"
    exit 1
}

Write-Host "=== SharePoint Online Site Contents Retrieval ===" -ForegroundColor Cyan
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

# Initialize collections and counters
$allItems = @()
$listCount = 0
$libraryCount = 0
$totalSize = 0
$totalVersionSize = 0
$totalItemCount = 0
$totalVersionCount = 0

Write-Host "Retrieving site contents..." -ForegroundColor Cyan
Write-Host ""

try {
    # Get all lists and libraries
    $lists = Get-PnPList -Includes BaseType, ItemCount | Where-Object { $_.Hidden -eq $false }
    
    foreach ($list in $lists) {
        $listTitle = $list.Title
        $isLibrary = $list.BaseType -eq "DocumentLibrary"
        
        if ($isLibrary) {
            $libraryCount++
            $listType = "Document Library"
        }
        else {
            $listCount++
            $listType = "List"
        }
        
        Write-Host "Processing: [$listType] $listTitle..." -ForegroundColor Yellow
        
        try {
            # Get all items in the list/library
            $items = Get-PnPListItem -List $list.Id -PageSize 5000 -WarningAction SilentlyContinue
            
            foreach ($item in $items) {
                $itemSize = 0
                $versionCount = 0
                $versionSize = 0
                $itemModified = $null
                $itemCreated = $null
                $itemCreatedBy = $null
                $itemModifiedBy = $null
                $itemType = "Item"
                
                # Get field values
                $fields = $item.FieldValues
                
                # Extract common fields
                $itemName = $fields["FileLeafRef"] -or $fields["Title"] -or "N/A"
                $itemModified = $fields["Modified"]
                $itemCreated = $fields["Created"]
                $itemCreatedBy = $fields["Author"] -or $fields["CreatedBy"] -or "N/A"
                $itemModifiedBy = $fields["Editor"] -or $fields["ModifiedBy"] -or "N/A"
                
                # Handle document-specific fields
                if ($isLibrary -and $fields["File_x0020_Size"]) {
                    $itemSize = [int]$fields["File_x0020_Size"]
                    $itemType = "Document"
                }
                elseif ($isLibrary -and $fields["SMTotalFileStreamSize"]) {
                    $itemSize = [int]$fields["SMTotalFileStreamSize"]
                    $itemType = "Document"
                }
                else {
                    $itemType = "Item"
                }
                
                # Get relative URL
                $itemURL = ""
                if ($fields["FileRef"]) {
                    $itemURL = $fields["FileRef"]
                }
                
                # Get version information for documents
                if ($isLibrary -and $itemType -eq "Document" -and $fields["FileRef"]) {
                    try {
                        $versions = Get-PnPFileVersion -Url $fields["FileRef"] -WarningAction SilentlyContinue
                        if ($versions) {
                            # Filter out the current version (it's the first one returned)
                            $priorVersions = $versions | Select-Object -Skip 1
                            $versionCount = @($priorVersions).Count
                            
                            # Calculate version sizes
                            foreach ($version in $priorVersions) {
                                if ($version.Size) {
                                    $versionSize += $version.Size
                                }
                            }
                        }
                    }
                    catch {
                        # Silently skip version retrieval errors for individual items
                    }
                }
                
                # Add to collection
                $allItems += [PSCustomObject]@{
                    "List/Library Name" = $listTitle
                    "List/Library Type" = $listType
                    "Item Name"         = $itemName
                    "Item Type"         = $itemType
                    "Size (bytes)"      = $itemSize
                    "Version Count"     = $versionCount
                    "Version Size (bytes)" = $versionSize
                    "Created"           = $itemCreated
                    "Modified"          = $itemModified
                    "Created By"        = $itemCreatedBy
                    "Modified By"       = $itemModifiedBy
                    "Item URL"          = $itemURL
                }
                
                $totalSize += $itemSize
                $totalVersionSize += $versionSize
                $totalVersionCount += $versionCount
                $totalItemCount++
            }
            
            Write-Host "  [OK] Processed $($items.Count) items" -ForegroundColor Green
        }
        catch {
            Write-Host "  [ERR] Error processing list '$listTitle': $_" -ForegroundColor Red
        }
    }
}
catch {
    Write-Error "Error retrieving site contents: $_"
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

# Calculate statistics
$totalSizeGB = [math]::Round($totalSize / 1GB, 2)
$totalSizeMB = [math]::Round($totalSize / 1MB, 2)
$totalVersionSizeGB = [math]::Round($totalVersionSize / 1GB, 2)
$totalVersionSizeMB = [math]::Round($totalVersionSize / 1MB, 2)
$combinedSizeGB = [math]::Round(($totalSize + $totalVersionSize) / 1GB, 2)
$combinedSizeMB = [math]::Round(($totalSize + $totalVersionSize) / 1MB, 2)
$avgItemSize = if ($totalItemCount -gt 0) { [math]::Round($totalSize / $totalItemCount, 2) } else { 0 }

# Display summary
Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "           SITE CONTENTS SUMMARY REPORT              " -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "CONTAINER STATISTICS:" -ForegroundColor Yellow
Write-Host "  Total Lists:           $listCount"
Write-Host "  Total Libraries:       $libraryCount"
Write-Host ""
Write-Host "ITEM STATISTICS:" -ForegroundColor Yellow
Write-Host "  Total Items:           $totalItemCount"
Write-Host "  Average Item Size:     $avgItemSize bytes"
Write-Host ""
Write-Host "CURRENT VERSION SIZE:" -ForegroundColor Yellow
Write-Host "  Total Size (GB):       $totalSizeGB GB"
Write-Host "  Total Size (MB):       $totalSizeMB MB"
Write-Host "  Total Size (Bytes):    $totalSize bytes"
Write-Host ""
Write-Host "VERSION HISTORY SIZE:" -ForegroundColor Yellow
Write-Host "  Total Versions:        $totalVersionCount"
Write-Host "  Version Size (GB):     $totalVersionSizeGB GB"
Write-Host "  Version Size (MB):     $totalVersionSizeMB MB"
Write-Host "  Version Size (Bytes):  $totalVersionSize bytes"
Write-Host ""
Write-Host "COMBINED SIZE (Current + Versions):" -ForegroundColor Yellow
Write-Host "  Combined Size (GB):    $combinedSizeGB GB"
Write-Host "  Combined Size (MB):    $combinedSizeMB MB"
Write-Host "  Combined Size (Bytes): $(($totalSize + $totalVersionSize)) bytes"
Write-Host ""
Write-Host "Script completed successfully!"
Write-Host "Results saved to: $OutputPath"
Write-Host ""
