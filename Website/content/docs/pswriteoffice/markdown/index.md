---
title: Markdown Cmdlets
description: PSWriteOffice cmdlets and DSL aliases for generating Markdown files from PowerShell.
order: 64
---

# Markdown Cmdlets

PSWriteOffice includes a Markdown surface for generating reports, READMEs, changelogs, and automation output from PowerShell. The cmdlets are especially useful when you already have structured objects in a script and want clean Markdown without hand-building strings.

## Core workflow

1. Start a document with `New-OfficeMarkdown`.
2. Add headings, paragraphs, tables, code blocks, or task lists.
3. Save the `.md` output as a report, repository artifact, or generated documentation file.

## Quick start

```powershell
$summary = @(
    [pscustomobject]@{ Metric = 'Uptime'; Value = '99.97%' }
    [pscustomobject]@{ Metric = 'Latency'; Value = '42 ms' }
)

New-OfficeMarkdown -Path .\report.md {
    MarkdownHeading -Level 1 -Text 'Summary'
    MarkdownParagraph -Text 'Generated automatically from the nightly validation run.'
    MarkdownTable -InputObject $summary
    MarkdownCode -Language 'powershell' -Content 'Get-Service | Select-Object -First 5'
}
```

## Multi-section report pattern

```powershell
$summary = Get-Process | Select-Object -First 5 Name, Id, CPU
$details = Get-Service | Select-Object -First 5 Name, Status, StartType

New-OfficeMarkdown -Path .\operations-report.md {
    MarkdownHeading -Level 1 -Text 'Operations Report'
    MarkdownTableOfContents -Title 'Contents' -MinLevel 2 -MaxLevel 2 -PlaceAtTop

    MarkdownHeading -Level 2 -Text 'Processes'
    MarkdownTable -InputObject $summary

    MarkdownHeading -Level 2 -Text 'Services'
    MarkdownTable -InputObject $details

    MarkdownCallout -Kind note -Title 'Next step' -Body 'Review the failed services before publishing.'
}
```

## Useful commands

- `New-OfficeMarkdown` starts a document and can save immediately or return the document for further piping.
- `Add-OfficeMarkdownHeading` and `Add-OfficeMarkdownParagraph` build the narrative structure.
- `Add-OfficeMarkdownTable` turns PowerShell objects into Markdown tables quickly.
- `Add-OfficeMarkdownCode` and `Add-OfficeMarkdownCallout` are useful for README-style and operational documentation.
- `Add-OfficeMarkdownTaskList` and `Add-OfficeMarkdownTableOfContents` help when the file needs to be actionable or navigable.

## When to use it

- Generate release notes, status reports, and inventory summaries from scheduled jobs.
- Produce README fragments or docs artifacts during CI.
- Emit Markdown that can later be consumed by `OfficeIMO.Markdown` or site generation workflows.

## Related guides

- [PSWriteOffice overview](/docs/pswriteoffice/) -- Module-level installation and command map.
- [Markdown overview](/docs/markdown/) -- .NET package concepts and rendering pipeline.
- [OfficeIMO.Markdown product page](/products/markdown/) -- Install command and package positioning.
