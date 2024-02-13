
$siteUrl = "<--- SITE URL -->"
$documentLibrary = "<-- Name of the document library -->"
$serverRelativePath = "<-- server relative path (e.g., /sites/{SiteName}}/Shared Documents) -->"


Connect-PnPOnline -Url $siteUrl -Interactive
$list = Get-PnPList -Identity $documentLibrary


$global:counter = 0
$itemCount =  $list.ItemCount
$listItems = Get-PnPListItem -List $list -PageSize 500 -Fields ID, Title, FileDirRef, FileLeafRef  -ScriptBlock { 
  Param($items) 
  $global:counter += $items.Count
  Write-Progress -PercentComplete ($global:Counter / $itemCount * 100) -Activity "Getting Items from List" -Status "Getting Items $global:Counter of $($itemCount)"
}
Write-Progress -Activity "Completed Getting Items from Library $($list.Title)" -Completed


$itemsToDelete = $listItems | Where-Object {$_.FieldValues.FileDirRef -eq $serverRelativePath}

$batch = New-PnPBatch
foreach ($item in $itemsToDelete) {
  Remove-PnPListItem -List $list -Identity $item.ID -Recycle -Batch $batch
  #Write-host "Deleted Item:"$Item.ServerRelativeURL
}
Invoke-PnPBatch -Batch $batch -Details
#$items.FieldValues.FileLeafRefclear