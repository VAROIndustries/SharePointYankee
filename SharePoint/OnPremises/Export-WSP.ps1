# Export-WSP.ps1
# Extract solution packages (WSPs) from a SharePoint farm
#
# Blog Post: https://sharepointyankee.com/extracting-solution-packages-wsps-from-sharepoint-using-powershell
# Author: Geoff Varosky
# Website: https://sharepointyankee.com

$farm = Get-SPFarm
$farm.Solutions | ForEach-Object {
    $_.SolutionFile.SaveAs("C:\temp\$($_.Name)")
    Write-Host "Exported: $($_.Name)"
}
