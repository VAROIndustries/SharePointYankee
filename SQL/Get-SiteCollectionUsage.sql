-- Get-SiteCollectionUsage.sql
-- Query a SharePoint content database to list all site collections and their disk usage
--
-- Blog Post: https://sharepointyankee.com/calculating-site-collection-usage-via-sql
-- Author: Geoff Varosky
-- Website: https://sharepointyankee.com

SELECT FullUrl AS URL, (DiskUsed / 1024) AS SiteCollectionUsedKB
FROM Sites
