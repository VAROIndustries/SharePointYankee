param(
    [Parameter(Mandatory = $true)]
    [string]$SiteURL,

    [Parameter(Mandatory = $false)]
    [string]$TenantType = "Commercial",

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\RecycleBin-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
)

# Validate tenant type
$validTenants = @("Commercial", "GCC")
if ($TenantType -notin $validTenants) {
    Write-Error "Invalid TenantType. Use 'Commercial' or 'GCC'"
    exit 1
}

Write-Host "=== SharePoint Online Recycle Bin Retrieval ===" -ForegroundColor Cyan
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

$allItems = @()
$firstStageCount = 0
$secondStageCount = 0
$firstStageSize = 0
$secondStageSize = 0

try {
    # First stage recycle bin
    Write-Host ""
    Write-Host "Retrieving First Stage recycle bin items..." -ForegroundColor Cyan
    try {
        $firstStage = Get-PnPRecycleBinItem -FirstStage -WarningAction SilentlyContinue
        foreach ($item in $firstStage) {
            $allItems += [PSCustomObject]@{
                "Stage"            = "First Stage"
                "Item Title"       = $item.Title
                "Original Location"= $item.DirName
                "Item Type"        = $item.ItemType
                "Size (bytes)"     = $item.Size
                "Deleted By"       = $item.DeletedByName
                "Deleted Date"     = $item.DeletedDate
                "Item ID"          = $item.Id
                "Original List ID" = $item.ListId
            }
            $firstStageSize += $item.Size
            $firstStageCount++
        }
        Write-Host "[OK] Found $firstStageCount first stage items" -ForegroundColor Green
    }
    catch {
        Write-Host "[ERR] Error retrieving first stage recycle bin: $_" -ForegroundColor Red
    }

    # Second stage recycle bin
    Write-Host "Retrieving Second Stage recycle bin items..." -ForegroundColor Cyan
    try {
        $secondStage = Get-PnPRecycleBinItem -SecondStage -WarningAction SilentlyContinue
        foreach ($item in $secondStage) {
            $allItems += [PSCustomObject]@{
                "Stage"            = "Second Stage"
                "Item Title"       = $item.Title
                "Original Location"= $item.DirName
                "Item Type"        = $item.ItemType
                "Size (bytes)"     = $item.Size
                "Deleted By"       = $item.DeletedByName
                "Deleted Date"     = $item.DeletedDate
                "Item ID"          = $item.Id
                "Original List ID" = $item.ListId
            }
            $secondStageSize += $item.Size
            $secondStageCount++
        }
        Write-Host "[OK] Found $secondStageCount second stage items" -ForegroundColor Green
    }
    catch {
        Write-Host "[ERR] Error retrieving second stage recycle bin: $_" -ForegroundColor Red
    }
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
$totalSize     = $firstStageSize + $secondStageSize
$firstSizeGB   = [math]::Round($firstStageSize  / 1GB, 2)
$firstSizeMB   = [math]::Round($firstStageSize  / 1MB, 2)
$secondSizeGB  = [math]::Round($secondStageSize / 1GB, 2)
$secondSizeMB  = [math]::Round($secondStageSize / 1MB, 2)
$totalSizeGB   = [math]::Round($totalSize        / 1GB, 2)
$totalSizeMB   = [math]::Round($totalSize        / 1MB, 2)

Write-Host ""
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host "           RECYCLE BIN SUMMARY REPORT                " -ForegroundColor Cyan
Write-Host "======================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "FIRST STAGE (user recycle bin):" -ForegroundColor Yellow
Write-Host "  Items:        $firstStageCount"
Write-Host "  Size (GB):    $firstSizeGB GB"
Write-Host "  Size (MB):    $firstSizeMB MB"
Write-Host "  Size (bytes): $firstStageSize bytes"
Write-Host ""
Write-Host "SECOND STAGE (site collection recycle bin):" -ForegroundColor Yellow
Write-Host "  Items:        $secondStageCount"
Write-Host "  Size (GB):    $secondSizeGB GB"
Write-Host "  Size (MB):    $secondSizeMB MB"
Write-Host "  Size (bytes): $secondStageSize bytes"
Write-Host ""
Write-Host "TOTAL (both stages):" -ForegroundColor Yellow
Write-Host "  Items:        $($allItems.Count)"
Write-Host "  Size (GB):    $totalSizeGB GB"
Write-Host "  Size (MB):    $totalSizeMB MB"
Write-Host "  Size (bytes): $totalSize bytes"
Write-Host ""
Write-Host "Script completed successfully!"
Write-Host "Results saved to: $OutputPath"
Write-Host ""
