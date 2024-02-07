# 1. Dependencies

## Microsoft.Graph.Planner
Used to manage interaction with Planner via the Graph API.

Docs at: https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation?view=graph-powershell-1.0

```PowerShell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

Install-Module Microsoft.Graph.Planner -Scope CurrentUser
```

## ImportExcell
Used to interact with the exported Excel file from the original Template Planner

Docs at: https://www.powershellgallery.com/packages/ImportExcel/7.8.4



```PowerShell
Install-Module -Name ImportExcel -Scope CurrentUser
```

# 2. Usage
1. Export your Planner to Excel through the Planner Web Application.
2. Create a new Planner in your target M365 Tenant.
3. Update the following variables of the Script file "Copy-PlannerPlan.ps1"
  ```PowerShell
  $ExcelPath = "<-- Full path of the Excel you exported from the source Planner -->"
  $TargetPlannerId = "<-- The new Planner's ID that you can get from its URL -->"
```
4. Execute the file and follow the interactive login and instructions from the console.