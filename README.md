# SharePoint Yankee Scripts

A collection of PowerShell scripts, code snippets, and configuration files from [SharePointYankee.com](https://sharepointyankee.com) covering SharePoint, Azure, Microsoft 365, and related Microsoft technologies.

Many of these scripts were originally published as part of blog posts over the years. Others come from a decade-plus of real-world consulting, administration, and development work. This repo puts them all in one place for easy reference and reuse.

**Blog:** [https://sharepointyankee.com](https://sharepointyankee.com)

---

## SharePoint On-Premises

| Script | Description | Blog Post |
|---|---|---|
| [Disable-MDS-OnPrem.ps1](SharePoint/OnPremises/Disable-MDS-OnPrem.ps1) | Disable Minimal Download Strategy across all sites in a web application | [Read More](https://sharepointyankee.com/disable-minimal-download-strategy-across-all-sites-and-site-collections-via-powershell) |
| [Remove-FileVersions-OnPrem.ps1](SharePoint/OnPremises/Remove-FileVersions-OnPrem.ps1) | Delete all previous versions of a SharePoint file | [Read More](https://sharepointyankee.com/delete-all-versions-of-a-file-using-powershell) |
| [Get-WebsAndTemplates.ps1](SharePoint/OnPremises/Get-WebsAndTemplates.ps1) | List all webs and site templates in a site collection | [Read More](https://sharepointyankee.com/powershell-script-to-list-all-webs-and-site-templates-in-use-within-a-site-collection) |
| [Approve-ListItems.ps1](SharePoint/OnPremises/Approve-ListItems.ps1) | Bulk-approve all items in a SharePoint list | [Read More](https://sharepointyankee.com/lotd-using-powershell-to-approve-list-items) |
| [Remove-MasterPage.ps1](SharePoint/OnPremises/Remove-MasterPage.ps1) | Delete a master page from a SharePoint site | [Read More](https://sharepointyankee.com/delete-a-master-page-in-sharepoint-using-powershell) |
| [Deploy-SPFeature.ps1](SharePoint/OnPremises/Deploy-SPFeature.ps1) | Install and activate features across MySite personal sites | [Read More](https://sharepointyankee.com/deploying-and-activating-features-in-sharepoint-2010-with-powershell) |
| [Export-WSP.ps1](SharePoint/OnPremises/Export-WSP.ps1) | Extract solution packages (WSPs) from a SharePoint farm | [Read More](https://sharepointyankee.com/extracting-solution-packages-wsps-from-sharepoint-using-powershell) |
| [Get-SiteCollectionStorage.ps1](SharePoint/OnPremises/Get-SiteCollectionStorage.ps1) | Get storage usage of a site collection | [Read More](https://sharepointyankee.com/how-much-storage-space-is-my-site-collection-using) |
| [Migrate-SharePointUsers.ps1](SharePoint/OnPremises/Migrate-SharePointUsers.ps1) | Migrate all users between domains across a SharePoint 2019 farm | |
| [Test-SPFarmHealth.ps1](SharePoint/OnPremises/Test-SPFarmHealth.ps1) | Monitor farm health: disk, CPU, memory, timer jobs, search, and IIS across all servers | |

## SharePoint Online

| Script | Description | Blog Post |
|---|---|---|
| [Connect-PnPSharePointOnline.ps1](SharePoint/Online/Connect-PnPSharePointOnline.ps1) | Connect to SharePoint Online using PnP PowerShell with stored credentials | [Read More](https://sharepointyankee.com/connecting-to-sharepoint-online-using-the-pnp-powershell-library-and-not-having-to-log-in-every-single-time) |
| [Copy-SPPermissionsCrossTenant.ps1](SharePoint/Online/Copy-SPPermissionsCrossTenant.ps1) | Copy item-level permissions between tenants (e.g. Commercial to GCC High) | |
| [Disable-SharePointMDS.ps1](SharePoint/Online/Disable-SharePointMDS.ps1) | Disable MDS across SharePoint Online sites via PnP PowerShell | |
| [Export-SPSitePermissions.ps1](SharePoint/Online/Export-SPSitePermissions.ps1) | Export all unique permissions across a site including broken inheritance | |
| [Get-SPGroupMembers.ps1](SharePoint/Online/Get-SPGroupMembers.ps1) | Retrieve members of a SharePoint group via PnP | |
| [Get-SPOStorageReport.ps1](SharePoint/Online/Get-SPOStorageReport.ps1) | Report on storage usage across SharePoint Online site collections | |
| [New-BulkSharePointSites.ps1](SharePoint/Online/New-BulkSharePointSites.ps1) | Bulk-create SharePoint Online sites from a CSV | |
| [Remove-SPFileVersions.ps1](SharePoint/Online/Remove-SPFileVersions.ps1) | Remove file versions in SharePoint Online via PnP | |
| [Set-SPOExternalSharing.ps1](SharePoint/Online/Set-SPOExternalSharing.ps1) | Configure external sharing settings across site collections | |

## Azure & Microsoft 365

| Script | Description | Blog Post |
|---|---|---|
| [Connect-AzureRunAs.ps1](Azure/Snippets/Connect-AzureRunAs.ps1) | Authenticate using a Run As service principal in Azure Automation | [Read More](https://sharepointyankee.com/creating-an-azure-run-as-account) |
| [Get-AzureKeyVaultSecretValue.ps1](Azure/Snippets/Get-AzureKeyVaultSecretValue.ps1) | Get the plain text value of an Azure Key Vault secret | [Read More](https://sharepointyankee.com/getting-the-plain-text-value-of-an-azure-key-vault-secret-with-powershell) |
| [Export-M365LicenseReport.ps1](Azure/Export-M365LicenseReport.ps1) | Export a comprehensive Microsoft 365 license usage report to CSV | |
| [New-GraphBulkUserReport.ps1](Azure/New-GraphBulkUserReport.ps1) | Generate bulk user reports with MFA status and last sign-in via Microsoft Graph | |
| [Get-AzureUpdatesReport.ps1](Azure/Monitoring/Get-AzureUpdatesReport.ps1) | Fetch and format Azure service update reports from RSS | |

## Monitoring

| Script | Description | Blog Post |
|---|---|---|
| [Test-WebsiteHealth.ps1](Monitoring/Test-WebsiteHealth.ps1) | Test website health (HTTP status, response time, SSL cert expiry) | [Read More](https://sharepointyankee.com/creating-runbooks-in-azure-and-calling-them-from-sharepoint-using-webhooks-and-flow) |
| [Get-HttpResponseHeaders.ps1](Monitoring/Get-HttpResponseHeaders.ps1) | Retrieve and display HTTP response headers from any URL | [Read More](https://sharepointyankee.com/powershell-script-to-get-http-headers) |

## Email

| Script | Description |
|---|---|
| [Send-SendGridEmail.ps1](Email/Send-SendGridEmail.ps1) | Send email via the SendGrid v3 REST API |

## SQL

| Script | Description | Blog Post |
|---|---|---|
| [Copy-FBAUsers.sql](SQL/Copy-FBAUsers.sql) | Generate PowerShell commands to migrate FBA users from SP 2007 to 2010 | [Read More](https://sharepointyankee.com/using-powershell-and-sql-to-copy-users-from-sharepoint-2007-to-2010) |
| [Get-SiteCollectionUsage.sql](SQL/Get-SiteCollectionUsage.sql) | Query a content database for site collection disk usage | [Read More](https://sharepointyankee.com/calculating-site-collection-usage-via-sql) |

## JavaScript

| Script | Description | Blog Post |
|---|---|---|
| [HideSearchScopes.js](JavaScript/HideSearchScopes.js) | Hide the search scopes drop-down in WSS v3/MOSS 2007 | [Read More](https://sharepointyankee.com/hiding-the-search-scopes-drop-down-in-wssv3moss-2007) |
| [NintexAutoComplete.js](JavaScript/NintexAutoComplete.js) | Replace a Nintex Forms drop-down with an autocompleting textbox | [Read More](https://sharepointyankee.com/replacing-a-drop-down-list-in-nintex-forms-2010-with-an-autocompleting-textbox-fix-for-version-1-11-4-0-update) |

## Registry

| File | Description | Blog Post |
|---|---|---|
| [IE-DisableScriptErrors.reg](Registry/IE-DisableScriptErrors.reg) | Disable IE script error notifications | [Read More](https://sharepointyankee.com/internet-explorer-registry-shortcuts-for-enabling-and-disabling-scripting-error-notifications) |
| [IE-EnableScriptErrors.reg](Registry/IE-EnableScriptErrors.reg) | Enable IE script error notifications | [Read More](https://sharepointyankee.com/internet-explorer-registry-shortcuts-for-enabling-and-disabling-scripting-error-notifications) |

## Config

| File | Description | Blog Post |
|---|---|---|
| [PeoplePickerWildcards.xml](Config/PeoplePickerWildcards.xml) | web.config snippet for FBA wildcard search in People Picker | [Read More](https://sharepointyankee.com/wildcard-search-for-forms-based-authentication-users-in-the-sharepoint-2010-people-picker-not-working) |
| [robots-sharepoint.txt](Config/robots-sharepoint.txt) | Block search engines from indexing SharePoint search result pages | [Read More](https://sharepointyankee.com/stay-away-from-my-search-result-pages-insert-search-engine-name-here-bot) |

---

## About

These scripts span over a decade of SharePoint, Azure, and Microsoft 365 administration and development work. Many include header comments linking back to the original blog post with full context, explanations, and screenshots.

**Blog:** [https://sharepointyankee.com](https://sharepointyankee.com)

## Disclaimer

These scripts are provided as-is for educational and reference purposes. Some target older versions of SharePoint (2007, 2010, 2013, 2016) while others target current SharePoint Online and Azure services. Always test in a non-production environment first.

## License

MIT License - See [LICENSE](LICENSE) for details.
