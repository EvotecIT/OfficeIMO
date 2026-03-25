---
title: "PowerShell Meets Office: PSWriteOffice in 5 Minutes"
description: "Get started with PSWriteOffice, the PowerShell module that wraps OfficeIMO for quick Word and Excel document generation from the command line."
date: 2025-06-15
tags: [powershell, pswriteoffice, automation]
categories: [Tutorial]
author: "Przemyslaw Klys"
---

Not every document automation task warrants a full C# project. Sometimes you just need to generate an Excel summary from a CSV or spin up a Word letter from a template, and you want to do it in ten lines of PowerShell. That is exactly what **PSWriteOffice** is built for.

## Installation

PSWriteOffice is published to the PowerShell Gallery:

```powershell
Install-Module -Name PSWriteOffice -Scope CurrentUser
```

It pulls in OfficeIMO as a dependency automatically. No manual DLL management required.

## Creating a Word Document

```powershell
Import-Module PSWriteOffice

New-OfficeWord -FilePath "Welcome.docx" {
    New-OfficeWordText -Text "Welcome to Contoso" -Bold -FontSize 22
    New-OfficeWordText -Text "We are pleased to confirm your enrolment."
    New-OfficeWordText -Text ""
    New-OfficeWordText -Text "Your start date is January 6, 2025."
    New-OfficeWordTable -DataTable @(
        [PSCustomObject]@{ Item = "Laptop"; Status = "Shipped" }
        [PSCustomObject]@{ Item = "Badge";  Status = "Ready" }
        [PSCustomObject]@{ Item = "Parking"; Status = "Assigned" }
    )
}
```

Run the script and open `Welcome.docx`. You get a formatted heading, body text, and a table, all generated without Microsoft Word installed.

## Creating an Excel Workbook

```powershell
$sales = Import-Csv "sales.csv"

New-OfficeExcel -FilePath "SalesReport.xlsx" {
    New-OfficeExcelSheet -Name "Q1 Sales" {
        # Headers
        New-OfficeExcelRow -Values "Region", "Revenue", "Units"

        foreach ($row in $sales) {
            New-OfficeExcelRow -Values $row.Region, $row.Revenue, $row.Units
        }
    }
    New-OfficeExcelSheet -Name "Summary" {
        New-OfficeExcelRow -Values "Total Regions", $sales.Count
    }
}
```

The DSL reads naturally: you declare a workbook, add sheets, and fill rows. PSWriteOffice translates each call into the corresponding OfficeIMO.Excel API.

## Automating with Scheduled Tasks

Combine PSWriteOffice with a Windows Scheduled Task or a cron job on Linux to generate recurring reports:

```powershell
# daily-report.ps1
$events = Get-EventLog -LogName Application -Newest 100

New-OfficeExcel -FilePath "C:\Reports\EventLog_$(Get-Date -Format yyyyMMdd).xlsx" {
    New-OfficeExcelSheet -Name "Events" {
        New-OfficeExcelRow -Values "Time", "Source", "Message"
        foreach ($e in $events) {
            New-OfficeExcelRow -Values $e.TimeGenerated, $e.Source, $e.Message
        }
    }
}
```

Schedule it with `schtasks` or `Register-ScheduledTask` and you have hands-free daily reporting.

## Tips and Tricks

- **Pipeline input.** Pipe `Get-Process`, `Get-Service`, or any cmdlet output directly into `New-OfficeExcelRow` via `ForEach-Object`.
- **Styling.** Use `-Bold`, `-Italic`, `-FontSize`, `-FontColor`, and `-BackgroundColor` on text and cell commands.
- **Multiple documents.** Call `New-OfficeWord` and `New-OfficeExcel` in the same script to produce a matched pair of files from one data source.
- **Cross-platform.** PSWriteOffice runs on PowerShell 7+ on Windows, macOS, and Linux.

## Why PSWriteOffice Over COM?

The legacy approach, `New-Object -ComObject Word.Application`, requires Office installed, leaks COM handles, and fails silently in non-interactive sessions. PSWriteOffice avoids all of that by using OfficeIMO's pure managed-code engine. No Office installation, no COM headaches, no leaked `WINWORD.EXE` processes.

Give it five minutes. You will not go back.
