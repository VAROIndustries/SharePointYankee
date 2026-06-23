# Remove-MasterPage.ps1
# Delete a master page from a SharePoint site
#
# Blog Post: https://sharepointyankee.com/delete-a-master-page-in-sharepoint-using-powershell
# Author: Geoff Varosky
# Website: https://sharepointyankee.com

$web = Get-SPWeb "http://yoursite"
$lib = $web.GetFolder("_catalogs/masterpage")
$file = $lib.Files["your-master-page.master"]
$file.Delete()
$web.Dispose()
