---
title: "OfficeIMO.OneNote"
description: "Read, create, edit, and write native OneNote sections, notebooks, tables of contents, and packages offline."
layout: product
meta.seo_title: "OfficeIMO.OneNote for .NET applications"
meta.head_html: '<link rel="alternate" hreflang="en" href="https://officeimo.com/products/onenote/" /><link rel="alternate" hreflang="x-default" href="https://officeimo.com/products/onenote/" />'
product_label: "Offline OneNote engine"
product_color: "#7c3aed"
install: "dotnet add package OfficeIMO.OneNote"
nuget: "OfficeIMO.OneNote"
docs_url: "/docs/onenote/"
api_url: "/api/onenote/"
---

## Native OneNote files without Graph or OneNote

`OfficeIMO.OneNote` works with desktop and FSSHTTP `.one` files, `.onetoc2` notebook tables of contents, `.onepkg` exports, and notebook directories. Both OneNote encodings project into the same typed model.

```csharp
using OfficeIMO.OneNote;

OneNoteSection section = OneNoteSectionReader.Read("Projects.one");
OneNotePage page = section.Pages[0];
var paragraph = new OneNoteParagraph();
paragraph.Runs.Add(new OneNoteTextRun { Text = "Validated offline" });
page.DirectContent.Add(paragraph);
section.Save("Projects-updated.one");
```

## From notebook storage to useful output

| Package | Role |
|---|---|
| `OfficeIMO.OneNote` | Native storage, revisions, pages, content, notebook structure, and writing |
| `OfficeIMO.OneNote.Markdown` | Shared semantic projection for text, lists, tables, links, assets, math, and conflicts |
| `OfficeIMO.OneNote.Html` | Responsive semantic HTML and position-preserving visual HTML |
| `OfficeIMO.OneNote.Pdf` | Semantic or position-preserving PDF output |
| `OfficeIMO.Reader.OneNote` | Normalized extraction and chunking through OfficeIMO.Reader |

File-format operations remain local and managed. Microsoft Graph, a OneNote installation, COM automation, and commercial SDKs are not required.
