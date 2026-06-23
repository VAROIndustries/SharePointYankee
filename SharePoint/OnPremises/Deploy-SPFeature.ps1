# Deploy-SPFeature.ps1
# Install and activate a SharePoint feature across all MySite personal sites
#
# Blog Post: https://sharepointyankee.com/deploying-and-activating-features-in-sharepoint-2010-with-powershell
# Author: Geoff Varosky
# Website: https://sharepointyankee.com

# Install the feature
Install-SPFeature -path "MyNewNavFeature"

# Enable on the MySite host
Enable-SPFeature -identity "MyNewNavFeature" -URL "http://mysitehost"

# Enable on all personal sites
$personalSites = Get-SPSite | Where-Object {$_.RootWeb.WebTemplate -eq "SPSPERS"}
foreach ($site in $personalSites) {
    Enable-SPFeature -Identity "MyNewNavFeature" -Url $site.Url
}
