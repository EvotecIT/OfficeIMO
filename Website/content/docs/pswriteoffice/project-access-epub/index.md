---
title: "Project, Access, and EPUB Files"
description: "Create and read Microsoft Project schedules, Access databases and EPUB books from PowerShell, without Project or Access installed."
layout: docs
---

These three families cover files that usually need their own desktop application: Microsoft Project schedules, Access table databases, and EPUB books. PSWriteOffice creates and reads them with OfficeIMO's managed engines, so they work in a scheduled job or a container on Windows, Linux, or macOS.

Each family follows the same pattern as Word and Excel: `New-` or `Open-` returns a live object, `Add-`, `Set-`, and `Get-` commands work on it through the pipeline, and `Save-` and `Close-` finish the job.

## Project schedules

Build a schedule from tasks and dependencies, let the scheduler calculate the dates, then save it as Project XML or render a Gantt view to PDF.

```powershell
$project = New-OfficeProject -Name 'Handbook delivery' -StartDate '2026-10-20 08:00'
try {
    $draft  = $project | Add-OfficeProjectTask -Name Draft -DurationDays 3 -PassThru
    $review = $project | Add-OfficeProjectTask -Name Review -DurationDays 2 -PassThru
    $project | Add-OfficeProjectDependency -Predecessor $draft -Successor $review
    $project | Invoke-OfficeProjectSchedule -Apply | Out-Null

    $project | Save-OfficeProject -Path .\delivery.xml
    $project | Get-OfficeProjectView -Kind Gantt | Export-OfficeProjectReport -Path .\delivery.pdf
    $project | Get-OfficeProjectTask | Select-Object Uid, Name, Start, Finish
} finally {
    $project | Close-OfficeProject
}
```

```text
Uid Name   Start               Finish
--- ----   -----               ------
  1 Draft  20.10.2026 08:00:00 22.10.2026 17:00:00
  2 Review 23.10.2026 08:00:00 26.10.2026 17:00:00
```

Review starts on the next working day after Draft finishes, and the weekend is skipped. `Invoke-OfficeProjectSchedule` without `-Apply` returns the calculated dates without changing the project, which is useful for a what-if check.

- **Resources and assignments:** `Add-OfficeProjectResource` and `Add-OfficeProjectAssignment`, read back with `Get-OfficeProjectResource` and `Get-OfficeProjectAssignment`.
- **Tables in and out:** `ConvertTo-OfficeProjectTable` maps ordinary rows, such as a CSV export from a tracker, to a task, resource, assignment, or calendar table, and `Import-OfficeProjectTable` builds a new project from those tables. `Get-OfficeProjectTable` goes the other way and returns a project's tables as rows for Excel or a report.

  ```powershell
  $tasks = Import-Csv .	asks.csv | ConvertTo-OfficeProjectTable -Kind Tasks -ColumnMap @{ Uid = 'ID'; Name = 'Task' }
  $project = $tasks | Import-OfficeProjectTable -Name Delivery -ReportVariable importReport
  ```

- **Native MPP:** saving to `.mpp` reports anything that version of the format cannot carry. Inspect it with `-ReportVariable` and add `-AllowLoss` only after reviewing those reductions. Project XML keeps the full supported scheduling model.
- **Rendering:** `Export-OfficeProjectReport` needs PowerShell 7.4 or newer.

## Access table databases

Create MDB and ACCDB files with tables, typed columns, and rows, or read rows back as objects.

```powershell
$database = New-OfficeAccess -Title Contacts
try {
    $contacts = $database | Add-OfficeAccessTable -Name Contacts -PassThru
    $contacts | Add-OfficeAccessColumn -Name Id -Type AutoNumber
    $contacts | Add-OfficeAccessColumn -Name Name -MaxLength 100
    $contacts | Add-OfficeAccessColumn -Name Active -Type Boolean

    Import-Csv .\teams.csv | Add-OfficeAccessRow -Table $contacts
    $database | Save-OfficeAccess -Path .\contacts.accdb
} finally {
    $database | Close-OfficeAccess
}

$database = Open-OfficeAccess -Path .\contacts.accdb
$database | Get-OfficeAccessRow -Table Contacts
$database | Close-OfficeAccess
```

These are file operations only. They don't start the Access database engine, run SQL, or execute macros or VBA stored in the file, so opening a database from someone else is a read of its tables and nothing more.

## EPUB publications

Compose a book from chapters written as plain text, Markdown, or XHTML, then add images and set the metadata a reader app shows.

```powershell
$handbook = New-OfficeEpub -Title 'Operations handbook' -Language en -Author 'Operations team'
$handbook | Add-OfficeEpubChapter -Title Welcome -Text 'Safety & care come first.'
$handbook | Add-OfficeEpubChapter -Title Procedures -Markdown (Get-Content .\procedures.md -Raw)
$handbook | Save-OfficeEpub -Path .\handbook.epub
```

Plain chapter text is escaped, so `&` and `<` appear as written. Use `-Markdown` or `-Xhtml` when the chapter has its own structure. `Add-OfficeEpubResource` adds images and other files, and `Set-OfficeEpubMetadata` changes the title, author, and language. `Open-OfficeEpub` and `Get-OfficeEpubChapter` read an existing book.

To publish an EPUB as PDF or another format, use `Convert-OfficeDocument`; the [conversion routes guide](/docs/pswriteoffice/conversion-routes/) lists what each route keeps.

## Runnable example

The [native document families example](https://github.com/EvotecIT/OfficeIMO/blob/e71f09178f1e1231bd7901cbda405bd4f6b4718b/PowerShell/PSWriteOffice/Examples/Workflows/Example-NativeDocumentFamilies.ps1) creates all three files in one run:

```powershell
.\Example-NativeDocumentFamilies.ps1 -OutputDirectory .\out
```

For every parameter, search the [command reference](/api/powershell/) for `OfficeProject`, `OfficeAccess`, or `OfficeEpub`. The .NET libraries behind these commands are described in the [Project](/docs/project/), [Access](/docs/access/), and [EPUB](/docs/epub/) guides.
