---
title: Excel Cmdlets
description: PSWriteOffice cmdlets for creating and editing Excel workbooks in PowerShell.
order: 62
---

# Excel Cmdlets

PSWriteOffice provides PowerShell cmdlets for creating, editing, and saving Excel workbooks. The examples below use the workbook DSL and the concrete commands present in the module help.

## Creating a Workbook

```powershell
Import-Module PSWriteOffice

# Create a new workbook and return the workbook object
$excel = New-OfficeExcel -Path "C:\Output\data.xlsx" -PassThru
```

## Adding Worksheets

```powershell
New-OfficeExcel -Path "C:\Output\data.xlsx" {
    Add-OfficeExcelSheet -Name "Summary"
    Add-OfficeExcelSheet -Name "Details"
    Add-OfficeExcelSheet -Name "Charts"
}
```

## Setting Cell Values

Use `Set-OfficeExcelCell` inside an `Add-OfficeExcelSheet` block:

```powershell
New-OfficeExcel -Path "C:\Output\data.xlsx" {
    Add-OfficeExcelSheet -Name "Employees" {
        Set-OfficeExcelCell -Address "A1" -Value "Name"
        Set-OfficeExcelCell -Address "B1" -Value "Department"
        Set-OfficeExcelCell -Address "C1" -Value "Salary"

        Set-OfficeExcelCell -Address "A2" -Value "Alice"
        Set-OfficeExcelCell -Address "B2" -Value "Engineering"
        Set-OfficeExcelCell -Address "C2" -Value 95000

        Set-OfficeExcelCell -Address "A3" -Value "Bob"
        Set-OfficeExcelCell -Address "B3" -Value "Design"
        Set-OfficeExcelCell -Address "C3" -Value 85000
    }
}
```

## Formulas

```powershell
New-OfficeExcel -Path "C:\Output\data.xlsx" {
    Add-OfficeExcelSheet -Name "Summary" {
        Set-OfficeExcelCell -Address "A1" -Value "Revenue"
        Set-OfficeExcelCell -Address "A2" -Value 1200
        Set-OfficeExcelCell -Address "A3" -Value 1800
        Set-OfficeExcelCell -Address "A4" -Formula "SUM(A2:A3)"
    }
}
```

## Adding Tables

```powershell
New-OfficeExcel -Path "C:\Output\data.xlsx" {
    Add-OfficeExcelSheet -Name "Employees" {
        Set-OfficeExcelCell -Address "A1" -Value "Name"
        Set-OfficeExcelCell -Address "B1" -Value "Department"
        Set-OfficeExcelCell -Address "C1" -Value "Salary"
        Set-OfficeExcelCell -Address "A2" -Value "Alice"
        Set-OfficeExcelCell -Address "B2" -Value "Engineering"
        Set-OfficeExcelCell -Address "C2" -Value 95000
        Set-OfficeExcelCell -Address "A3" -Value "Bob"
        Set-OfficeExcelCell -Address "B3" -Value "Design"
        Set-OfficeExcelCell -Address "C3" -Value 85000
        Add-OfficeExcelTable -Range "A1:C3" -TableName "EmployeeTable" -Style "Medium2"
    }
}
```

## Populating from PowerShell Objects

A common workflow is to export PowerShell object data to Excel:

```powershell
$processes = Get-Process | Sort-Object -Property WorkingSet64 -Descending |
    Select-Object -First 20 -Property Name, Id, @{N='Memory (MB)';E={[math]::Round($_.WorkingSet64 / 1MB, 1)}}, CPU

New-OfficeExcel -Path "C:\Output\processes.xlsx" {
    Add-OfficeExcelSheet -Name "Top Processes" {
        Set-OfficeExcelCell -Address "A1" -Value "Process Name"
        Set-OfficeExcelCell -Address "B1" -Value "PID"
        Set-OfficeExcelCell -Address "C1" -Value "Memory (MB)"
        Set-OfficeExcelCell -Address "D1" -Value "CPU Time"

        $row = 2
        foreach ($p in $processes) {
            Set-OfficeExcelCell -Row $row -Column 1 -Value $p.Name
            Set-OfficeExcelCell -Row $row -Column 2 -Value $p.Id
            Set-OfficeExcelCell -Row $row -Column 3 -Value $p.'Memory (MB)'
            Set-OfficeExcelCell -Row $row -Column 4 -Value $(if ($null -ne $p.CPU) { [math]::Round($p.CPU, 2) } else { 0 })
            $row++
        }

        Add-OfficeExcelTable -Range "A1:D$($row - 1)" -TableName "TopProcesses" -Style "Medium2"
    }
}
```

## Opening Existing Workbooks

```powershell
$excel = Get-OfficeExcel -Path "C:\Data\existing.xlsx"

foreach ($sheet in $excel.Sheets) {
    Write-Host "Sheet: $($sheet.Name)"
}
```

## Saving and Closing

```powershell
$excel | Save-OfficeExcel
Close-OfficeExcel -Document $excel
```

## Complete Example: System Inventory Report

```powershell
Import-Module PSWriteOffice

New-OfficeExcel -Path "C:\Reports\Inventory.xlsx" {
    Add-OfficeExcelSheet -Name "Disks" {
        Set-OfficeExcelCell -Address "A1" -Value "Drive"
        Set-OfficeExcelCell -Address "B1" -Value "Label"
        Set-OfficeExcelCell -Address "C1" -Value "Size (GB)"
        Set-OfficeExcelCell -Address "D1" -Value "Free (GB)"

        $disks = Get-CimInstance Win32_LogicalDisk -Filter "DriveType=3"
        $row = 2
        foreach ($disk in $disks) {
            Set-OfficeExcelCell -Row $row -Column 1 -Value $disk.DeviceID
            Set-OfficeExcelCell -Row $row -Column 2 -Value $disk.VolumeName
            Set-OfficeExcelCell -Row $row -Column 3 -Value ([math]::Round($disk.Size / 1GB, 1))
            Set-OfficeExcelCell -Row $row -Column 4 -Value ([math]::Round($disk.FreeSpace / 1GB, 1))
            $row++
        }
    }

    Add-OfficeExcelSheet -Name "Network" {
        Set-OfficeExcelCell -Address "A1" -Value "Adapter"
        Set-OfficeExcelCell -Address "B1" -Value "IP Address"
        Set-OfficeExcelCell -Address "C1" -Value "Status"

        $adapters = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
        $row = 2
        foreach ($adapter in $adapters) {
            $ip = (Get-NetIPAddress -InterfaceIndex $adapter.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).IPAddress
            Set-OfficeExcelCell -Row $row -Column 1 -Value $adapter.Name
            Set-OfficeExcelCell -Row $row -Column 2 -Value $(if ($ip) { $ip } else { "N/A" })
            Set-OfficeExcelCell -Row $row -Column 3 -Value $adapter.Status
            $row++
        }
    }
}
```
