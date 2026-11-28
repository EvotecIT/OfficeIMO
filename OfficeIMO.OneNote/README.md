# OfficeIMO.OneNote

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.OneNote)](https://www.nuget.org/packages/OfficeIMO.OneNote)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.OneNote?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.OneNote)

Read, create, edit and write OneNote files in .NET without OneNote, Microsoft Graph or COM. `OfficeIMO.OneNote` works with desktop and FSSHTTP `.one` sections, `.onetoc2` notebooks, `.onepkg` exports and notebook folders through one typed model.

## Install

```sh
dotnet add package OfficeIMO.OneNote
```

## Quick start

<!-- powerforge:code source="Build/DocumentationExamples/OneNote/Program.cs" region="QuickStart" -->
```csharp
using OfficeIMO.OneNote;

using var section = OneNoteSection.Create("Planning.one");
section.Name = "Planning";
var page = new OneNotePage { Title = "Release checklist" };
var item = new OneNoteParagraph {
    List = new OneNoteListInfo { Ordered = false, Level = 0 }
};
item.Runs.Add(new OneNoteTextRun { Text = "Validate the packaged artifact" });
page.DirectContent.Add(item);
section.Pages.Add(page);

section.Save();
```
<!-- /powerforge:code -->

## What it does

- **Read and edit** sections, pages, outlines, paragraphs, lists, tables, images, note and task tags, authors and revisions.
- **Create** native sections and notebooks, reproducibly when you pass a fixed identity seed and timestamp.
- **Ink and math**: author ink strokes and editable equations.
- **Notebooks and packages**: read and write `.onetoc2` tables of contents and `.onepkg` Cabinet exports with bounded expansion.
- **Render** pages to images.
- **Bounded input** and streaming reads for large notebooks.

The [OneNote guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.onenote-guide.md) shows each of these with examples.

## Converting to other formats

| Target | Package |
| --- | --- |
| Markdown | [OfficeIMO.OneNote.Markdown](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.OneNote.Markdown/README.md) |
| HTML | [OfficeIMO.OneNote.Html](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.OneNote.Html/README.md) |
| PDF | [OfficeIMO.OneNote.Pdf](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.OneNote.Pdf/README.md) |

## Limits

Encrypted sections are detected but not decrypted, and some objects are written only within deliberate boundaries. The [OneNote current state](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.onenote.current-state.md) lists coverage and evidence.

## Learn more

- [OfficeIMO.OneNote on officeimo.com](https://officeimo.com/products/onenote/) and the [OneNote docs](https://officeimo.com/docs/onenote/).
- [OneNote guide](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.onenote-guide.md) and [examples](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Examples/OneNote).
- Reading for search and AI: [OfficeIMO.Reader.OneNote](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Reader.OneNote/README.md).

## Targets and dependencies

- Targets: `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.
- Dependencies: `OfficeIMO.Core` only.
- License: MIT.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 1 | 0 | 0 | 0 | 1 | 0 |
| Read | 1 | 0 | 0 | 0 | 1 | 0 |
| Edit | 1 | 0 | 0 | 1 | 0 | 0 |
| Preserve | 1 | 0 | 0 | 0 | 0 | 0 |
| Inspect | 1 | 1 | 0 | 0 | 0 | 0 |
| Validate | 0 | 0 | 0 | 0 | 0 | 1 |
| Remove | 0 | 0 | 0 | 0 | 1 | 0 |
| Export | 0 | 5 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.OneNote` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
