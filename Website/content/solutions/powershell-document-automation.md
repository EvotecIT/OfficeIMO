---
title: "Automate Office Documents with PowerShell"
description: "Use PSWriteOffice to create, inspect, and transform OfficeIMO documents from PowerShell scripts, scheduled jobs, CI/CD, and administration workflows."
meta.eyebrow: "PowerShell document automation"
meta.outcome: "Repeatable document jobs without COM automation"
meta.primary_label: "Read the PSWriteOffice guide"
meta.primary_url: "/docs/pswriteoffice/"
meta.powershell: true
---

PSWriteOffice is the first-party PowerShell surface over OfficeIMO. It is intended for administrators, automation engineers, and build pipelines that need document output without embedding C# in every script or launching desktop Office.

## Use it for repeatable jobs

Common workflows include scheduled reports, infrastructure documentation, inventory exports, audit evidence, migration jobs, and CI artifacts. Commands and task-oriented aliases cover Word, Excel, PowerPoint, PDF, email, Markdown, CSV, OpenDocument, Reader, and related package families.

```powershell
New-OfficeWord -Path '.\report.docx' {
    Add-OfficeWordSection {
        Add-OfficeWordParagraph -Text 'Quarterly report' -Style Heading1
    }
}
```

The public command catalog is generated from the module so help and API discovery remain aligned, while the task guides and examples are written for people solving real automation problems.

## Prefer data in, documents out

Keep collection and business logic separate from presentation. Pass structured objects into a document-building function, use a template or style policy consistently, and return the output path plus diagnostics. This makes scripts easier to test and reuse.

## Avoid COM as a scheduler dependency

Desktop Office automation is fragile in unattended sessions and difficult to run in Linux containers or hosted CI. PSWriteOffice uses managed libraries and does not require Word, Excel, or PowerPoint to be installed.

## Scale to .NET when needed

PowerShell is a thin automation surface over the same engines. If a job grows into a service, needs custom concurrency, or becomes a reusable application capability, move the orchestration into .NET without changing the document platform underneath.

Browse the [PSWriteOffice product page](/products/pswriteoffice/), [task guides](/docs/pswriteoffice/), and [cmdlet reference](/api/powershell/).
