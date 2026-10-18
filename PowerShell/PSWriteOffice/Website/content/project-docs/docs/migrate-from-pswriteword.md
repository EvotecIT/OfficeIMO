---
title: "Migrate from PSWriteWord to PSWriteOffice"
description: "Replace archived PSWriteWord scripts with maintained PSWriteOffice Word commands for DOCX creation, editing, review, merge, and conversion."
layout: docs
---

[PSWriteWord](https://github.com/EvotecIT/PSWriteWord) is archived and replaced by PSWriteOffice. New development should use the `OfficeWord` command family, which runs on the current OfficeIMO.Word engine without requiring Microsoft Word.

PSWriteOffice is not a drop-in namespace change. PSWriteWord exposed Xceed-era objects and the Documentimo DSL; PSWriteOffice uses OfficeIMO objects and a scoped document DSL. Rewrite the composition block and validate templates, styles, fields, headers, footers, and pagination with representative documents.

## Common command replacements

| PSWriteWord | PSWriteOffice | Migration note |
| --- | --- | --- |
| `New-WordDocument` | `New-OfficeWord` | Prefer the script-block DSL when creating a complete document. |
| `Get-WordDocument` | `Get-OfficeWord` | Use targeted `Get-OfficeWord*` commands to inspect content. |
| `Add-WordParagraph`, `Add-WordText` | `Add-OfficeWordParagraph`, `Add-OfficeWordText` | Text runs can be added inside an active paragraph context. |
| `Add-WordTable` | `Add-OfficeWordTable` | Pass objects with `-InputObject`; map old design names to current table styles. |
| `Add-WordPicture` | `Add-OfficeWordImage` | Recheck sizing, wrapping, and placement. |
| `Add-WordHeader`, `Add-WordFooter` | `Add-OfficeWordHeader`, `Add-OfficeWordFooter` | Configure them inside the appropriate section. |
| `Add-WordList` | `Add-OfficeWordList` | Nested list behavior should be verified against the generated DOCX. |
| `Add-WordTOC` | `Add-OfficeWordTableOfContents` | Use update commands when fields or headings change. |
| `Merge-WordDocument` | `Join-OfficeWordDocument` | Supply a base input and one or more append paths. |
| `Save-WordDocument` | `Save-OfficeWord` | `New-OfficeWord` saves automatically unless `-NoSave` is used. |
| `Documentimo`, `DocText`, `DocTable` | `New-OfficeWord` with `Add-OfficeWord*` commands | Rewrite the old DSL; aliases are available for interactive use, but canonical names are clearer in maintained scripts. |

## Before: PSWriteWord

```powershell
$document = New-WordDocument -FilePath '.\Output\Status.docx'
Add-WordText -WordDocument $document -Text 'Service status'
Add-WordTable -WordDocument $document -DataTable $services
Save-WordDocument -WordDocument $document
```

## After: PSWriteOffice

```powershell
$services = Get-Service |
    Select-Object -First 10 Name, Status

New-OfficeWord -Path '.\Output\Status.docx' {
    Add-OfficeWordSection {
        Add-OfficeWordParagraph -Text 'Service status' -Style Heading1
        Add-OfficeWordTable -InputObject $services -Style 'GridTable1LightAccent2'
    }
}
```

For an existing file, load it with `Get-OfficeWord`, make targeted changes, and pipe the returned document to `Save-OfficeWord`. Use `Compare-OfficeWordDocument` when the migration needs machine-readable evidence of document changes.

## Validate before replacing the old job

- compare page setup, section breaks, styles, lists, tables, images, and headers or footers;
- check fields and table-of-contents updates in the target Word viewer;
- verify password protection and template behavior explicitly;
- test charts, equations, content controls, comments, revisions, and mail merge when the old workflow used them.

See [Word automation](/docs/pswriteoffice/word/) and the generated [PowerShell command reference](/api/powershell/) for current syntax.
