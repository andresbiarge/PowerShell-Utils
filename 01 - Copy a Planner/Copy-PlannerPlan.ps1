#Update these variables first
#⚠️Ensure to close the Excel file before executing the script
$ExcelPath = "<-- Full path of the Excel you exported from the source Planner -->"
$TargetPlannerId = "<-- The new Planner's ID that you can get from its URL -->"


#------------------------------ Scrit starts here ------------------------------
$isExcelClosed = Read-Host -Prompt "⚠️ Please, confirm that the Excel to import is closed [Y/N]"
If ($isExcelClosed -ne "Y") {
  exit
}

#Authenticate interactively
Connect-MgGraph -Scopes Tasks.ReadWrite

#Check if Planner is OK
$TargetPlannerObject = Get-MgPlannerPlan -PlannerPlanId $TargetPlannerId

if (!$TargetPlannerObject) {
  Write-Error "The Planner ID is not valid"
  exit
}

#⚠️ Cleanup of resources in TARGET Planner
$isInitialCleanupConfirmed = Read-Host -Prompt "⚠️ Do you want to remove the existing tasks and buckets first? [Y/N]"
If ($isInitialCleanupConfirmed -eq "Y") {
  $existingTasks = Get-MgPlannerPlanTask -PlannerPlanId $TargetPlannerId
  foreach ($task in $existingTasks) {
    Remove-MgPlannerTask -PlannerTaskId $task.Id -IfMatch $task.AdditionalProperties."@odata.etag"
  }
  $existingBuckets = Get-MgPlannerPlanBucket -PlannerPlanId $TargetPlannerId
  foreach ($bucket in $existingBuckets) {
    Remove-MgPlannerBucket -PlannerBucketId $bucket.Id -IfMatch $bucket.AdditionalProperties."@odata.etag"
  }
}


Write-Host "Now we're going to create the buckets and tasks..."
$excelTemplate = Import-Excel -Path $ExcelPath

Start-Sleep 1
#0: Create Categories (Planner labels)
$excelTasksLabels = $excelTemplate | Select-Object -Property "Labels" -Unique | Sort-Object -Property "Labels"
$categoryDescriptions = [PSCustomObject]@{}
for ($i = 0; $i -lt $excelTasksLabels.Count; $i++) {
  $auxCategoryKey = "category$($i+1)"
  $auxCategoryValue = $excelTasksLabels[$i].Labels
  $categoryDescriptions | Add-Member -MemberType NoteProperty -Name $auxCategoryKey -Value $auxCategoryValue
}
$TargetPlannerDetailsObject = Get-MgPlannerPlanDetail -PlannerPlanId $TargetPlannerId
$params = [PSCustomObject]@{
  "sharedWith" = $TargetPlannerDetailsObject.SharedWith
  "categoryDescriptions" = $categoryDescriptions
}
Update-MgPlannerPlanDetail -PlannerPlanId $TargetPlannerId -BodyParameter ($params | ConvertTo-Json -Depth 100) -IfMatch $TargetPlannerDetailsObject.AdditionalProperties."@odata.etag"

Start-Sleep 1
#1: Create Buckets
$bucketsToCreate = $excelTemplate | Select-Object -Property "Bucket Name" -Unique | Sort-Object -Property "Bucket Name"
$createdBuckets = @()
foreach($bucketHashObject in $bucketsToCreate) {
  $params = @{
    name = $bucketHashObject.'Bucket Name'
    planId = $TargetPlannerId
    orderHint = " !"
  }
  $createdBuckets += New-MgPlannerBucket -BodyParameter $params
}
Start-Sleep 1
#2: Create all Tasks
foreach($excelRow in $excelTemplate) {
  #2.1: Create Task
  $targetBucketForCurrentTask = ($createdBuckets | Where-Object {$_."Name" -eq $excelRow."Bucket Name"}).Id
  $params = @{
    planId = $TargetPlannerId
    bucketId = $targetBucketForCurrentTask
    title = $excelRow."Task Name"
    appliedCategories = @{}
  }
  foreach ($key in $categoryDescriptions.PsObject.Properties) {
    if ($excelRow."Labels" -match $categoryDescriptions."$($key.Name)") {
      $params.appliedCategories.Add($key.Name, $true)
    }
  }
  $auxCreatedTask = New-MgPlannerTask -BodyParameter $params

  #2.2: Update Task Details
  $checklistItems = [PSCustomObject]@{}
  if ($excelRow."Checklist Items") {
    foreach ($checkItem in ($excelRow."Checklist Items" -Split ";")) {
      $auxChecklistItem = [PSCustomObject]@{
        "isChecked" = $false
        "orderHint" = " !"
        "title" = $checkItem
      }
      $auxNewGuid = $(New-Guid).Guid
      $checklistItems | Add-Member -MemberType NoteProperty  -Name $auxNewGuid -Value $auxChecklistItem
    }
  }
  $params = @{
    previewType = "description"
    checklist = $checklistItems
    description = $excelRow."Description" -replace "_x000d_", ""
  }
  $auxCurrentTaskDetails = Get-MgPlannerTaskDetail -PlannerTaskId $auxCreatedTask.Id
  Update-MgPlannerTaskDetail -PlannerTaskId $auxCreatedTask.Id -BodyParameter $params -IfMatch $auxCurrentTaskDetails.AdditionalProperties."@odata.etag"
}

