# Disable-MDS.ps1
# Disable Minimal Download Strategy across all sites in a web application
#
# Blog Post: https://sharepointyankee.com/disable-minimal-download-strategy-across-all-sites-and-site-collections-via-powershell
# Author: Geoff Varosky
# Website: https://sharepointyankee.com

# Load the SharePoint PowerShell Module... if not running this in the SharePoint Console...
Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue

# URL to our web application
$WebApp = "https://sharepoint.contoso.com"

# Get all webs within a web application
$Webs = Get-SPWebApplication $WebApp | Get-SPSite -Limit All | Get-SPWeb -Limit All

# Loop through said webs
foreach ($Web in $Webs)
{
    # Is MDS enabled?
    $MDSEnabled = Get-SPFeature -web $Web.URL | Where-Object {$_.DisplayName -eq "MDSFeature"}

    # If it is... disable it!
    if ($MDSEnabled -ne $null)
    {
        Disable-SPFeature -identity "MDSFeature" -URL $Web.URL -confirm:$false
    }
}
