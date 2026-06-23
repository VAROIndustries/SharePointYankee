# Get-WebsAndTemplates.ps1
# List all webs and their site templates within a site collection
#
# Blog Post: https://sharepointyankee.com/powershell-script-to-list-all-webs-and-site-templates-in-use-within-a-site-collection
# Author: Geoff Varosky
# Website: https://sharepointyankee.com

$site = Get-SPSite "http://yoursite"

foreach ($web in $site.AllWebs) {
    $web | Select-Object -Property Title, Url, WebTemplate
}

$site.Dispose()
