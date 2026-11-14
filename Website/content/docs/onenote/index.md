---
title: Read and Edit OneNote Files
description: Open OneNote .one section files offline in .NET, list pages and their text, add notes, lists and tags, and create new sections without OneNote, Graph or COM.
order: 37
meta.seo_title: "Read and write OneNote .one files in .NET | OfficeIMO"
search_terms:
  - onenote
  - one file
  - read onenote
  - create onenote section
  - onenote without graph
packages:
  - OfficeIMO.OneNote
  - OfficeIMO.OneNote.Markdown
cmdlets:
  - Open-OfficeOneNote
  - New-OfficeOneNoteReadOptions
commands:
  - officeimo reader read
limits: Encrypted (password-protected) sections are detected and reported but not decrypted.
verified: 2026-11-14
---

`OfficeIMO.OneNote` reads and writes OneNote files directly: desktop `.one` revision stores, FSSHTTP-encoded `.one` files, `.onetoc2` notebook tables of contents, `.onepkg` exports and notebook folders. A section loads into `OneNoteSection`, with pages, outlines, paragraphs, runs, lists, tables, images, ink, math and tags. Content the model does not cover is kept and written back when you save an edited file.

| Artifact | Read | Create and write |
|---|:---:|:---:|
| Desktop `.one` revision store | Yes | Yes |
| FSSHTTP-encoded `.one` | Yes | Yes |
| `.onetoc2` notebook table of contents | Yes | Yes |
| `.onepkg` notebook export | Yes | Yes |
| Notebook folder | Yes | Yes |

```sh
dotnet add package OfficeIMO.OneNote
```

## Create a section with pages

New sections use the same model as loaded ones. Writers read the output back by default to make sure nothing was lost.

### C#

```csharp
using OfficeIMO.OneNote;

using OneNoteSection section = OneNoteSection.Create("Planning.one");
section.Name = "Planning";

var page = new OneNotePage { Title = "Release checklist" };

foreach (string task in new[] { "Validate the package", "Update the docs", "Publish" }) {
    var item = new OneNoteParagraph {
        List = new OneNoteListInfo { Ordered = false, Level = 0 }
    };
    item.Runs.Add(new OneNoteTextRun { Text = task });
    page.DirectContent.Add(item);
}

section.Pages.Add(page);
section.Save();
```

## List pages and read their text

`ToMarkdown` from `OfficeIMO.OneNote.Markdown` projects the current pages into a Markdown document, which is the simplest way to get text out.

### C#

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Markdown;

using OneNoteSection section = OneNoteSection.Load("Planning.one");

foreach (OneNotePage page in section.Pages) {
    string indent = new string(' ', page.Level * 2);
    Console.WriteLine($"{indent}{page.Title} (modified {page.LastModifiedUtc:u}, by {page.MostRecentAuthor})");
}

MarkdownDocument markdown = section.ToMarkdown();
Console.WriteLine(markdown.ToMarkdown());
```

### PowerShell

```powershell
$section = Open-OfficeOneNote -Path .\Planning.one
$section.Pages | Select-Object Title, CreatedUtc, LastModifiedUtc
```

### CLI

```sh
officeimo reader read Planning.one --format markdown --output Planning.md
```

## Add a note to an existing page

Content added to `DirectContent` is moved into a standard outline when the section is written, which is the shape OneNote itself uses.

### C#

```csharp
using OfficeIMO.OneNote;

using OneNoteSection section = OneNoteSection.Load("Planning.one");
OneNotePage page = section.Pages[0];

var paragraph = new OneNoteParagraph();
paragraph.Runs.Add(new OneNoteTextRun { Text = "Added offline by OfficeIMO" });
page.DirectContent.Add(paragraph);

section.Save("Planning-updated.one");
```

## Read a large or untrusted file with limits

`OneNoteLoadOptions` bounds input bytes, objects, properties, nesting and assets.

### C#

```csharp
using OfficeIMO.OneNote;

var options = new OneNoteLoadOptions {
    MaxInputBytes = 64L * 1024 * 1024
};

using OneNoteSection section = OneNoteSection.Load("Planning.one", options);
Console.WriteLine($"{section.Pages.Count} pages");
```

### PowerShell

```powershell
$options = New-OfficeOneNoteReadOptions -MaxInputBytes 67108864
Open-OfficeOneNote -Path .\Planning.one -Options $options
```

## Save in FSSHTTP format

A loaded file keeps its original encoding. Choose the storage format explicitly when you need the FSSHTTP package form.

### C#

```csharp
using OfficeIMO.OneNote;

using OneNoteSection section = OneNoteSection.Load("Planning.one");
section.Save("Planning-fsshttp.one", new OneNoteSaveOptions {
    StorageFormat = OneNoteStorageFormat.FileSynchronizationPackage
});
```
