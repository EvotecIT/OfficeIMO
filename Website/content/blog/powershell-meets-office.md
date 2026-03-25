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

It ships as a PowerShell module with the OfficeIMO document engines underneath, so you do not need to manage Open XML assemblies yourself.

## Creating a Word Document

```powershell
Import-Module PSWriteOffice

New-OfficeWord -Path ".\Welcome.docx" {
    Add-OfficeWordSection {
        Add-OfficeWordParagraph -Style Heading1 -Text "Welcome to Contoso"
        Add-OfficeWordParagraph -Text "We are pleased to confirm your enrolment."
        Add-OfficeWordParagraph -Text "Your start date is January 6, 2025."
        Add-OfficeWordTable -InputObject @(
            [PSCustomObject]@{ Item = "Laptop";  Status = "Shipped" }
            [PSCustomObject]@{ Item = "Badge";   Status = "Ready" }
            [PSCustomObject]@{ Item = "Parking"; Status = "Assigned" }
        ) -Style GridTable4Accent1
    }
}
```

Run the script and open `Welcome.docx`. You get a heading, body text, and a generated table without Microsoft Word installed.

## Creating an Excel Workbook

```powershell
$sales = Import-Csv ".\sales.csv"

New-OfficeExcel -Path ".\SalesReport.xlsx" {
    Add-OfficeExcelSheet -Name "Q1 Sales" {
        Set-OfficeExcelCell -Address "A1" -Value "Region"
        Set-OfficeExcelCell -Address "B1" -Value "Revenue"
        Set-OfficeExcelCell -Address "C1" -Value "Units"

        $rowIndex = 2
        foreach ($row in $sales) {
            Set-OfficeExcelCell -Row $rowIndex -Column 1 -Value $row.Region
            Set-OfficeExcelCell -Row $rowIndex -Column 2 -Value $row.Revenue
            Set-OfficeExcelCell -Row $rowIndex -Column 3 -Value $row.Units
            $rowIndex++
        }
    }

    Add-OfficeExcelSheet -Name "Summary" {
        Set-OfficeExcelCell -Address "A1" -Value "Total Regions"
        Set-OfficeExcelCell -Address "B1" -Value $sales.Count
    }
}
```

The DSL reads naturally: create a workbook, add sheets, and populate cells from script data.

## Automating with Scheduled Tasks

Combine PSWriteOffice with a Windows Scheduled Task or a cron job on Linux to generate recurring reports:

```powershell
# daily-report.ps1
$events = Get-EventLog -LogName Application -Newest 100

New-OfficeExcel -Path "C:\Reports\EventLog_$(Get-Date -Format yyyyMMdd).xlsx" {
    Add-OfficeExcelSheet -Name "Events" {
        Set-OfficeExcelCell -Address "A1" -Value "Time"
        Set-OfficeExcelCell -Address "B1" -Value "Source"
        Set-OfficeExcelCell -Address "C1" -Value "Message"

        $rowIndex = 2
        foreach ($e in $events) {
            Set-OfficeExcelCell -Row $rowIndex -Column 1 -Value $e.TimeGenerated
            Set-OfficeExcelCell -Row $rowIndex -Column 2 -Value $e.Source
            Set-OfficeExcelCell -Row $rowIndex -Column 3 -Value $e.Message
            $rowIndex++
        }
    }
}
```

Schedule it with `schtasks` or `Register-ScheduledTask` and you have hands-free daily reporting.

## Tips and Tricks

- **Tables and structured data.** Use `Add-OfficeWordTable -InputObject` or `Add-OfficePowerPointTable -Data` when your source is already object-based.
- **Inline text formatting.** Use `Add-OfficeWordText` inside `Add-OfficeWordParagraph` blocks when you need bold, italic, or underlined runs.
- **Multiple documents.** Call `New-OfficeWord` and `New-OfficeExcel` in the same script to produce a matched pair of files from one data source.
- **Cross-platform.** PSWriteOffice runs on PowerShell 7+ on Windows, macOS, and Linux.

## Why PSWriteOffice Over COM?

The legacy approach, `New-Object -ComObject Word.Application`, requires Office installed, leaks COM handles, and fails silently in non-interactive sessions. PSWriteOffice avoids all of that by using OfficeIMO's managed document engine. No Office installation, no COM headaches, no leaked `WINWORD.EXE` processes.

Give it five minutes. You will not go back.
