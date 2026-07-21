---
title: "OneNote"
description: "Work with native OneNote sections, notebooks, tables of contents, and packages offline."
order: 37
meta.seo_title: "OneNote file automation | OfficeIMO"
meta.head_html: '<link rel="alternate" hreflang="en" href="https://officeimo.com/docs/onenote/" /><link rel="alternate" hreflang="x-default" href="https://officeimo.com/docs/onenote/" />'
---

## Install

```shell
dotnet add package OfficeIMO.OneNote
```

## Read, edit, and save a section

```csharp
using OfficeIMO.OneNote;

OneNoteSection section = OneNoteSectionReader.Read("Projects.one");
OneNotePage page = section.Pages[0];

var paragraph = new OneNoteParagraph();
paragraph.Runs.Add(new OneNoteTextRun { Text = "Added offline" });
page.DirectContent.Add(paragraph);

section.Save("Projects-updated.one");
```

## Supported artifacts

| Artifact | Read | Create and write |
|---|:---:|:---:|
| Desktop `.one` revision store | Yes | Yes |
| FSSHTTP-encoded `.one` | Yes | Yes |
| `.onetoc2` notebook table of contents | Yes | Yes |
| `.onepkg` notebook export | Yes | Yes |
| Notebook directory | Yes | Yes |

Add `OfficeIMO.OneNote.Markdown`, `.Html`, or `.Pdf` for conversion, or `OfficeIMO.Reader.OneNote` for normalized extraction. See the [OneNote API reference](/api/onenote/) for storage and model details.
