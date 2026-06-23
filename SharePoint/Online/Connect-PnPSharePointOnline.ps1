# Connect-PnPSharePointOnline.ps1
# Connect to SharePoint Online using the PnP PowerShell Library with stored credentials
#
# Blog Post: https://sharepointyankee.com/connecting-to-sharepoint-online-using-the-pnp-powershell-library-and-not-having-to-log-in-every-single-time
# Author: Geoff Varosky
# Website: https://sharepointyankee.com

#region Imports
Import-Module SharePointPnPPowerShellOnline -WarningAction SilentlyContinue
#endregion Imports

#region Variables
$Username = "admin@yourtenant.onmicrosoft.com"
$Password = "YourPasswordHere"
$SiteCollection = "https://yourtenant.sharepoint.com/sites/yoursite"
#endregion Variables

#region Credentials
[SecureString]$SecurePass = ConvertTo-SecureString $Password -AsPlainText -Force
[System.Management.Automation.PSCredential]$PSCredentials = New-Object System.Management.Automation.PSCredential($Username, $SecurePass)
#endregion Credentials

#region ConnectPnPOnline
try {
    Connect-PnPOnline -Url $SiteCollection -Credentials $PSCredentials
    if (-not (Get-PnPContext)) {
        Write-Host "Error connecting to SharePoint Online, unable to establish context" -ForegroundColor Black -BackgroundColor Red
        return
    }
} catch {
    Write-Host "Error connecting to SharePoint Online: $($_.Exception.Message)" -ForegroundColor Black -BackgroundColor Red
    return
}
#endregion ConnectPnPOnline
