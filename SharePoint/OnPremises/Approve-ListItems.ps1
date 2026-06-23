# Approve-ListItems.ps1
# Bulk-approve all items in a SharePoint list
#
# Blog Post: https://sharepointyankee.com/lotd-using-powershell-to-approve-list-items
# Author: Geoff Varosky
# Website: https://sharepointyankee.com

$web = Get-SPWeb "http://site"
$list = $web.Lists["Posts"]
$items = $list.Items

foreach ($item in $items) {
    $item["_ModerationStatus"] = 0
    $item.Update()
}

$web.Dispose()
