# Get-SiteCollectionStorage.ps1
# Get the storage usage of a SharePoint site collection
#
# Blog Post: https://sharepointyankee.com/how-much-storage-space-is-my-site-collection-using
# Author: Geoff Varosky
# Website: https://sharepointyankee.com

$site = Get-SPSite "http://yoursite"
$storageUsedMB = [math]::Round($site.Usage.Storage / 1MB, 2)
Write-Host "Site Collection: $($site.Url)"
Write-Host "Storage Used: $storageUsedMB MB"
$site.Dispose()
