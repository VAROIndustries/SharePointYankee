# Remove-FileVersions.ps1
# Delete all previous versions of a SharePoint file (on-premises only)
#
# Usage: .\Remove-FileVersions.ps1 "http://mysharepoint.com/sites/foo/Documents/Document_1.docx"
#
# Blog Post: https://sharepointyankee.com/delete-all-versions-of-a-file-using-powershell
# Author: Geoff Varosky
# Website: https://sharepointyankee.com

param(
    [string] $UrlToFile
)

$Site = New-Object -Type Microsoft.SharePoint.SPSite -ArgumentList $UrlToFile
$Web = $Site.OpenWeb()
$SPFile = $Web.GetFile($UrlToFile)

Write-Host "Deleting all versions for file $($UrlToFile)..."

# Remove all versions for file...
$SPFile.Versions.DeleteAll()

# Dispose of the web object
$Web.Dispose()

# Dispose of the site object
$Site.Dispose()
