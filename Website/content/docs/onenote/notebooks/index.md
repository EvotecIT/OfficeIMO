---
title: Read and Write OneNote Notebooks and Packages
description: Open OneNote notebooks from a folder, an Open Notebook.onetoc2 file or a .onepkg export, walk section groups and sections, and write notebooks and packages offline.
order: 38
search_terms:
  - onenote notebook
  - onetoc2
  - onepkg
  - onenote section groups
  - export onenote notebook
packages:
  - OfficeIMO.OneNote
cmdlets:
  - Open-OfficeOneNote
  - New-OfficeOneNoteNotebookOptions
related:
  - /docs/onenote/
  - /docs/onenote/export-and-convert/
limits: .onepkg reading supports single Cabinet archives with uncompressed, MSZIP or LZX folders; Quantum compression and multi-file cabinet sets are not supported.
verified: 2026-11-14
---

A OneNote notebook is a folder with an `Open Notebook.onetoc2` table of contents, `.one` section files and subfolders for section groups. A `.onepkg` file is the same notebook packed into a Cabinet archive. `OneNoteNotebookReader` opens the folder form and `OneNotePackageReader` opens the package; both return a `OneNoteNotebook` with section groups, sections and pages. `OneNoteNotebookWriter` and `OneNotePackageWriter` write them back.

## Open a notebook folder or table of contents

### C#

```csharp
using OfficeIMO.OneNote;

OneNoteNotebook notebook = OneNoteNotebookReader.Read(
    Path.Combine("Work notebook", "Open Notebook.onetoc2"),
    new OneNoteNotebookReaderOptions {
        IncludeRecycleBin = false,
        ContinueOnSectionError = true
    });

Console.WriteLine(notebook.Name);
foreach (OneNoteDiagnostic diagnostic in notebook.Diagnostics) {
    Console.WriteLine(diagnostic);
}
```

### PowerShell

```powershell
$notebook = Open-OfficeOneNote -Path '.\Work notebook\Open Notebook.onetoc2'
$notebook.Sections | Select-Object Name
```

## Walk section groups, sections and pages

### C#

```csharp
using OfficeIMO.OneNote;

OneNoteNotebook notebook = OneNotePackageReader.Read("Work notebook.onepkg");

void PrintGroup(string name, IList<OneNoteSection> sections, IList<OneNoteSectionGroup> groups, int depth) {
    string indent = new string(' ', depth * 2);
    Console.WriteLine($"{indent}[{name}]");
    foreach (OneNoteSection section in sections) {
        Console.WriteLine($"{indent}  {section.Name} ({section.Pages.Count} pages)");
        foreach (OneNotePage page in section.Pages) {
            Console.WriteLine($"{indent}    - {page.Title}");
        }
    }
    foreach (OneNoteSectionGroup group in groups) {
        PrintGroup(group.Name, group.Sections, group.SectionGroups, depth + 1);
    }
}

PrintGroup(notebook.Name, notebook.Sections, notebook.SectionGroups, 0);
```

### PowerShell

```powershell
$options = New-OfficeOneNoteNotebookOptions -IncludeRecycleBin
$notebook = Open-OfficeOneNote -Path '.\Work notebook.onepkg' -Options $options
$notebook.Sections | ForEach-Object { $_.Pages } | Select-Object Title, LastModifiedUtc
```

## List the notebook structure without loading pages

Set `LoadSectionContent = false` to read only the table of contents, which is fast for large notebooks.

### C#

```csharp
using OfficeIMO.OneNote;

OneNoteNotebook notebook = OneNoteNotebookReader.Read(
    Path.Combine("Work notebook", "Open Notebook.onetoc2"),
    new OneNoteNotebookReaderOptions { LoadSectionContent = false });

foreach (OneNoteSection section in notebook.Sections) {
    Console.WriteLine(section.Name);
}
foreach (OneNoteSectionGroup group in notebook.SectionGroups) {
    Console.WriteLine($"{group.Name}/ ({group.Sections.Count} sections)");
}
```

### PowerShell

```powershell
$options = New-OfficeOneNoteNotebookOptions -NoSectionContent
Open-OfficeOneNote -Path '.\Work notebook\Open Notebook.onetoc2' -Options $options
```

## Create a notebook and a .onepkg package

### C#

```csharp
using OfficeIMO.OneNote;

var notebook = new OneNoteNotebook { Name = "Offline notebook" };

var section = OneNoteSection.Create();
section.Name = "Meetings";
var page = new OneNotePage { Title = "Kickoff" };
var paragraph = new OneNoteParagraph();
paragraph.Runs.Add(new OneNoteTextRun { Text = "Agreed scope and timeline." });
page.DirectContent.Add(paragraph);
section.Pages.Add(page);
notebook.Sections.Add(section);

OneNoteNotebookWriter.Write(notebook, "Offline notebook");
OneNotePackageWriter.Write(notebook, "Offline notebook.onepkg");

OneNoteNotebook reopened = OneNotePackageReader.Read("Offline notebook.onepkg");
Console.WriteLine($"{reopened.Name}: {reopened.Sections.Count} section(s)");
```
