-- Copy-FBAUsers.sql
-- Generate New-SPUser PowerShell commands to migrate FBA users from SharePoint 2007 to 2010 claims format
--
-- Blog Post: https://sharepointyankee.com/using-powershell-and-sql-to-copy-users-from-sharepoint-2007-to-2010
-- Author: Geoff Varosky
-- Website: https://sharepointyankee.com

SELECT 'New-SPUser -UserAlias "'
    + REPLACE(tp_Login, 'acaspnetsqlmembershipprovider:', 'i:0#.f|sql-membershipprovider|')
    + '" -Web http://internet -DisplayName "'
    + tp_Title + '" -Email "'
    + tp_Email + '"'
FROM UserInfo
WHERE tp_SiteID = 'YOUR-SITE-COLLECTION-GUID-HERE'
AND tp_Login LIKE 'acaspnetsqlmembershipprovider%'
