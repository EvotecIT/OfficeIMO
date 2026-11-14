---
title: Create and Edit RTF Documents
description: Create Rich Text Format (RTF) files in .NET, edit text and bookmarks, merge documents, apply a theme, and make surgical edits that keep the original RTF syntax.
order: 40
meta.seo_title: "Create and edit RTF in .NET with OfficeIMO"
search_terms:
  - rtf
  - rich text format
  - create rtf
  - edit rtf
  - rtf replace text
packages:
  - OfficeIMO.Rtf
cmdlets:
  - New-OfficeRtf
  - Add-OfficeRtfParagraph
  - Save-OfficeRtf
  - Update-OfficeRtfText
limits: Semantic saves normalize the document; use the lossless editor when untouched control words, unknown destinations or binary data must stay byte for byte.
verified: 2026-11-14
---

`OfficeIMO.Rtf` reads RTF into two views. `RtfDocument` is the editable semantic model: paragraphs, runs, tables, lists, notes, headers and footers. Saving it with `ToRtf()` writes normalized RTF. The syntax tree behind `EditLossless()` keeps every control word, so a targeted edit leaves the rest of the file exactly as it was. Pick the semantic model for authoring and the lossless editor for changing documents produced elsewhere.

```sh
dotnet add package OfficeIMO.Rtf
```

## Create an RTF document

### C#

```csharp
using OfficeIMO.Rtf;

RtfDocument document = RtfDocument.Create();
document.AddParagraph().AddText("Quarterly report").SetBold();
document.AddParagraph("Prepared by OfficeIMO");
document.AddParagraph("Revenue grew in every region.");

File.WriteAllText("report.rtf", document.ToRtf());
```

### PowerShell

```powershell
New-OfficeRtf -Path .\Note.rtf -Text 'Server maintenance starts at 22:00.'

$rtf = New-OfficeRtf
Add-OfficeRtfParagraph -Document $rtf -Text 'Maintenance completed.' -Bold
Add-OfficeRtfParagraph -Document $rtf -Text 'All services are back online.'
Save-OfficeRtf -Document $rtf -Path .\Maintenance.rtf
```

## Replace text and fill bookmarks

Semantic edits work on the model and write normalized RTF.

### C#

```csharp
using OfficeIMO.Rtf;

RtfDocument template = RtfDocument.Create();
template.AddParagraph("Contract between Contoso Ltd. and the customer.");
File.WriteAllText("contract.rtf", template.ToRtf());

RtfDocument document = RtfDocument.Load("contract.rtf");
document.InsertParagraph(0, "Confidential");
document.ReplaceText("Contoso Ltd.", "Contoso Europe");

File.WriteAllText("contract-updated.rtf", document.ToRtf());
```

`ReplaceBookmarkText("CustomerName", "…")` replaces the text inside a named bookmark in the same way.

## Edit an existing file without rewriting it

The lossless editor changes only the nodes you touch. Everything else, including control words OfficeIMO does not model, is written back unchanged.

### C#

```csharp
using OfficeIMO.Rtf;

RtfDocument source = RtfDocument.Create();
source.AddParagraph("Status: Draft");
File.WriteAllText("input.rtf", source.ToRtf());

RtfLosslessEditor editor = RtfDocument.LoadResult("input.rtf").EditLossless();
editor.ReplaceText("Draft", "Final");
editor.SetInfo(RtfDocumentInfoField.Title, "Status report");
editor.InsertRootParagraph(editor.RootNodeCount, "Approved by the review board.");
editor.SaveLossless("output.rtf");
```

### PowerShell

```powershell
Update-OfficeRtfText -Path .\Input.rtf -OutputPath .\Output.rtf -OldText Draft -NewText Final -PassThru
```

## Merge RTF documents

`AppendDocument` imports the other document's content with its fonts, colors, styles and lists. With `PreserveSections`, each appended document keeps its page setup and headers and footers.

### C#

```csharp
using OfficeIMO.Rtf;

RtfDocument cover = RtfDocument.Create();
cover.AddParagraph("Annual report");

RtfDocument body = RtfDocument.Create();
body.AddParagraph("Chapter 1");

RtfDocumentMergeResult merge = cover.AppendDocument(body,
    new RtfDocumentMergeOptions { PreserveSections = true });

foreach (var diagnostic in merge.Report.Diagnostics) {
    Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
}

File.WriteAllText("annual-report.rtf", merge.Document.ToRtf());
```

## Apply a theme

`ApplyTheme` restyles typography, spacing and table decoration while keeping content, lists and page geometry.

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Rtf;

RtfDocument document = RtfDocument.Create();
document.AddParagraph("Meeting notes");
document.ApplyTheme(OfficeTheme.Boardroom);

File.WriteAllText("notes-styled.rtf", document.ToRtf());
```
