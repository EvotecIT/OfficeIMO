---
title: Excel Cmdlets
description: PSWriteOffice cmdlets for creating and editing Excel workbooks in PowerShell.
order: 62
---

# Excel Cmdlets

PSWriteOffice provides PowerShell cmdlets for creating, editing, and saving Excel workbooks. These cmdlets wrap the OfficeIMO.Excel .NET library and let you build spreadsheets directly from PowerShell scripts.

## Creating a Workbook

```powershell
Import-Module PSWriteOffice

# Create a new workbook with a default sheet
$excel = New-OfficeExcel -FilePath "C:\Output\data.xlsx" -WorkSheetName "Sheet1"
```

## Adding Worksheets

```powershell
$excel | Add-OfficeExcelWorkSheet -Name "Summary"
$excel | Add-OfficeExcelWorkSheet -Name "Details"
$excel | Add-OfficeExcelWorkSheet -Name "Charts"
```

## Setting Cell Values

Access cells through the sheet object:

```powershell
$sheet = $excel.Sheets[0]

# Headers
$sheet.Cells["A1"].Value = "Name"
$sheet.Cells["B1"].Value = "Department"
$sheet.Cells["C1"].Value = "Salary"

# Data
$sheet.Cells["A2"].Value = "Alice"
$sheet.Cells["B2"].Value = "Engineering"
$sheet.Cells["C2"].Value = 95000

$sheet.Cells["A3"].Value = "Bob"
$sheet.Cells["B3"].Value = "Design"
$sheet.Cells["C3"].Value = 85000

$sheet.Cells["A4"].Value = "Carol"
$sheet.Cells["B4"].Value = "Marketing"
$sheet.Cells["C4"].Value = 90000
```

## Formulas

```powershell
$sheet.Cells["C5"].Value = "=SUM(C2:C4)"
$sheet.Cells["C6"].Value = "=AVERAGE(C2:C4)"
```

## Cell Formatting

```powershell
# Bold headers
$sheet.Cells["A1"].Bold = $true
$sheet.Cells["B1"].Bold = $true
$sheet.Cells["C1"].Bold = $true

# Number format for salary
$sheet.Cells["C2"].NumberFormat = "$#,##0"
$sheet.Cells["C3"].NumberFormat = "$#,##0"
$sheet.Cells["C4"].NumberFormat = "$#,##0"
```

## Adding Tables

```powershell
$excel | Add-OfficeExcelTable -SheetName "Sheet1" -Name "EmployeeTable" -Range "A1:C4" -Style "Medium2"
```

## Populating from PowerShell Objects

A common workflow is to export PowerShell object data to Excel:

```powershell
$processes = Get-Process | Sort-Object -Property WorkingSet64 -Descending |
    Select-Object -First 20 -Property Name, Id, @{N='Memory (MB)';E={[math]::Round($_.WorkingSet64 / 1MB, 1)}}, CPU

$excel = New-OfficeExcel -FilePath "C:\Output\processes.xlsx" -WorkSheetName "Top Processes"

$sheet = $excel.Sheets[0]

# Write headers
$headers = @("Process Name", "PID", "Memory (MB)", "CPU Time")
for ($col = 0; $col -lt $headers.Count; $col++) {
    $cellRef = [char](65 + $col) + "1"
    $sheet.Cells[$cellRef].Value = $headers[$col]
    $sheet.Cells[$cellRef].Bold = $true
}

# Write data
for ($row = 0; $row -lt $processes.Count; $row++) {
    $p = $processes[$row]
    $r = $row + 2
    $sheet.Cells["A$r"].Value = $p.Name
    $sheet.Cells["B$r"].Value = $p.Id
    $sheet.Cells["C$r"].Value = $p.'Memory (MB)'
    $sheet.Cells["D$r"].Value = if ($null -ne $p.CPU) { [math]::Round($p.CPU, 2) } else { 0 }
}

$excel | Save-OfficeExcel
$excel | Close-OfficeExcel
```

## Opening Existing Workbooks

```powershell
$excel = Get-OfficeExcel -FilePath "C:\Data\existing.xlsx"

foreach ($sheet in $excel.Sheets) {
    Write-Host "Sheet: $($sheet.Name)"
}
```

## Saving and Closing

```powershell
$excel | Save-OfficeExcel
$excel | Close-OfficeExcel
```

## Complete Example: System Inventory Report

```powershell
Import-Module PSWriteOffice

$excel = New-OfficeExcel -FilePath "C:\Reports\Inventory.xlsx" -WorkSheetName "Disks"

# Disk information
$sheet = $excel.Sheets[0]
$sheet.Cells["A1"].Value = "Drive"
$sheet.Cells["B1"].Value = "Label"
$sheet.Cells["C1"].Value = "Size (GB)"
$sheet.Cells["D1"].Value = "Free (GB)"
$sheet.Cells["A1"].Bold = $true
$sheet.Cells["B1"].Bold = $true
$sheet.Cells["C1"].Bold = $true
$sheet.Cells["D1"].Bold = $true

$disks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3"
$row = 2
foreach ($disk in $disks) {
    $sheet.Cells["A$row"].Value = $disk.DeviceID
    $sheet.Cells["B$row"].Value = $disk.VolumeName
    $sheet.Cells["C$row"].Value = [math]::Round($disk.Size / 1GB, 1)
    $sheet.Cells["D$row"].Value = [math]::Round($disk.FreeSpace / 1GB, 1)
    $row++
}

# Network adapters on a second sheet
$excel | Add-OfficeExcelWorkSheet -Name "Network"
$netSheet = $excel.Sheets[1]
$netSheet.Cells["A1"].Value = "Adapter"
$netSheet.Cells["B1"].Value = "IP Address"
$netSheet.Cells["C1"].Value = "Status"
$netSheet.Cells["A1"].Bold = $true
$netSheet.Cells["B1"].Bold = $true
$netSheet.Cells["C1"].Bold = $true

$adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
$row = 2
foreach ($adapter in $adapters) {
    $ip = (Get-NetIPAddress -InterfaceIndex $adapter.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).IPAddress
    $netSheet.Cells["A$row"].Value = $adapter.Name
    $netSheet.Cells["B$row"].Value = if ($ip) { $ip } else { "N/A" }
    $netSheet.Cells["C$row"].Value = $adapter.Status
    $row++
}

$excel | Save-OfficeExcel
$excel | Close-OfficeExcel

Write-Host "Inventory report saved."
```
