---
title: "Remove hidden data and metadata"
description: "Inspect and remove properties, comments, tracked changes, hidden content, notes, attachments, scripts and GPS data from Office, PDF, EPUB and image files."
layout: docs
---

`InspectHiddenData()` lists the metadata and hidden data in a document and changes nothing. `RemoveHiddenData(options)` removes the selected categories and returns the same report, with an outcome on every finding. The categories follow what Office's Document Inspector checks, and the same report covers Word, Excel, PowerPoint, Visio, OpenDocument, EPUB, PDF and JPEG, PNG, WebP, GIF and TIFF images.

## Clean a file

```csharp
using OfficeIMO;

OfficeHiddenDataReport report = OfficeDocument.RemoveHiddenData("Contract.docx", "Contract-shared.docx",
    new OfficeHiddenDataOptions {
        Kinds = OfficeHiddenDataKind.Default | OfficeHiddenDataKind.HiddenContent,
        TrackedChanges = OfficeTrackedChangesAction.AcceptAll,
        AnonymousAuthor = "Reviewer"
    });

foreach (OfficeHiddenDataFinding finding in report.Findings) {
    Console.WriteLine($"{finding.Outcome} {finding.Kind} at {finding.Location}: {finding.Description}");
}
```

`OfficeDocument.RemoveHiddenData` in the `OfficeIMO` package picks the format from the file extension, writes the cleaned copy atomically and never changes the input. The output must keep the input's extension. To work in memory, pass the bytes and a file name instead; the result holds the report and the cleaned bytes:

```csharp
using OfficeIMO;

byte[] original = File.ReadAllBytes("photo.jpg");
OfficeHiddenDataResult result = OfficeDocument.RemoveHiddenData(original, "photo.jpg");
File.WriteAllBytes("photo-shared.jpg", result.ToArray());
Console.WriteLine($"Changed: {result.Report.Changed}, {result.Length} bytes");
```

To only look, inspect. `OfficeDocument.InspectHiddenData` takes a path, or bytes and a file name, and changes nothing:

```csharp
using OfficeIMO;

OfficeHiddenDataReport found = OfficeDocument.InspectHiddenData("Contract.docx");
foreach (OfficeHiddenDataFinding finding in found.Findings) {
    Console.WriteLine($"{finding.Kind}: {finding.Description} {finding.Value}");
}
```

`OfficeDocument.SupportsHiddenData(path)` tells you whether a file name has a supported extension. `OfficeDocument.Inspect` with `IncludeHiddenData = true` adds the same report to a full inspection.

## On a loaded document

A document you already have open has the same two methods. Save it to keep the result.

```csharp
using OfficeIMO;
using OfficeIMO.Excel;

using var workbook = ExcelDocument.Load("Budget.xlsx");

OfficeHiddenDataReport report = workbook.RemoveHiddenData();
foreach (OfficeHiddenDataFinding finding in report.GetFindings(OfficeHiddenDataKind.ExternalReferences)) {
    Console.WriteLine($"{finding.Outcome}: {finding.Value} {finding.Note}");
}

workbook.Save("Budget-shared.xlsx");
```

`WordDocument`, `ExcelDocument`, `PowerPointDocument`, `VisioDocument`, `OdfDocument` and `EpubPublication` have `InspectHiddenData()` and `RemoveHiddenData(options)`. `PdfDocument`'s versions take a `CancellationToken`, and its `RemoveHiddenData` returns an `OfficeHiddenDataResult` with the rewritten bytes, because a PDF is cleaned in one verified rewrite:

```csharp
using OfficeIMO;
using OfficeIMO.Pdf;

using PdfDocument pdf = PdfDocument.Load("Report.pdf");
OfficeHiddenDataResult result = pdf.RemoveHiddenData(new OfficeHiddenDataOptions {
    Kinds = OfficeHiddenDataKind.Default | OfficeHiddenDataKind.Attachments
});
File.WriteAllBytes("Report-shared.pdf", result.ToArray());
```

## What each format reports

- **Word:** properties, comments, tracked changes, reviewer names, hidden text, document variables, custom XML, revision save IDs (RSIDs), the attached template and picture metadata. Comments and revisions that stay keep an author, because Word requires one; `PersonalIdentities` replaces the name with `AnonymousAuthor` (`Author` by default). Bibliography sources and custom XML bound to content controls are kept.
- **Excel:** notes, threaded comments and their authors, hidden sheets, rows, columns and names, external links, pivot cache source paths, custom XML, picture metadata and printer settings. Cached values from external workbooks are removed, but the links themselves are kept, because formulas and refreshes need them. Power Query data is kept and reported with a note. Deleting hidden sheets, rows and columns shifts or breaks formulas that refer to them, exactly as deleting them in Excel does.
- **PowerPoint:** properties, classic and modern comments, speaker notes, the thumbnail, custom XML and picture metadata. Hidden slides, hidden shapes and shapes placed entirely off the slide count as hidden content.
- **PDF:** document properties and XMP metadata, comments and markup, attachments, JavaScript and other actions, bookmarks, optional-content layers and Content Credentials. Signed PDFs are refused.
- **Images:** Exif, GPS, XMP and IPTC metadata, Content Credentials and AI source labels. Pixels are not re-encoded.
- **Visio, OpenDocument and EPUB:** properties, comments, pictures and their metadata, and provenance, as far as each format carries them.

The [hidden data support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.hidden-data-support-matrix.md) lists exactly what each category covers in each format.

## Choose categories

`OfficeHiddenDataOptions.Kinds` is a flags value. `Default` selects every category that leaves visible content alone; `All` adds the rest.

| Category | In `Default` | Covers |
| --- | --- | --- |
| `DocumentProperties` | Yes | Author, last modified by, last printed, revision number, company, manager, template, editing time |
| `CustomProperties` | Yes | Custom document properties |
| `Comments` | Yes | Comments, replies, notes, threaded comments and PDF markup |
| `PersonalIdentities` | Yes | People parts, and authors on retained comments and revisions |
| `SpeakerNotes` | Yes | PowerPoint notes slides |
| `Thumbnail` | Yes | The package preview picture |
| `CustomXml` | Yes | Custom XML data parts |
| `ImageMetadata` | Yes | Exif, GPS, XMP and IPTC inside pictures; pixels are not re-encoded |
| `PrinterSettings` | Yes | Embedded printer settings |
| `ExternalReferences` | Yes | Attached templates and cached external values; links that content depends on are kept |
| `DocumentVariables` | Yes | Word document variables |
| `RevisionIdentifiers` | Yes | Word RSIDs |
| `ActiveContent` | Yes | Scripts and actions that run or open something, such as PDF JavaScript |
| `TrackedChanges` | No | Reported; changed only through `TrackedChanges` |
| `HiddenContent` | No | Hidden text, sheets, rows, columns, names, slides, shapes and PDF layers |
| `Attachments` | No | Files embedded in the document |
| `Bookmarks` | No | PDF bookmarks, whose titles can show draft headings |
| `Provenance` | No | Content Credentials and AI source labels; see [provenance](/docs/before-sharing/provenance/) |

## Pictures and provenance

Picture metadata cleanup and provenance overlap: an AI-generated picture often says so in its XMP metadata, and Content Credentials are bound to the exact bytes of the image. While `Provenance` is not selected, `ImageMetadata` keeps a picture's digital source type declaration and removes the rest, and leaves pictures that carry Content Credentials untouched. Those findings say so through `Reason`. Add `OfficeHiddenDataKind.Provenance` to remove the records too, and use `options.Provenance` to choose which kinds of record go.

A picture stored in a package that nothing in the document shows is reported with `Reason` `UnusedImage`; cleaning removes the whole picture.

## Read the report

```csharp
using OfficeIMO;

OfficeHiddenDataReport report = OfficeDocument.RemoveHiddenData("Contract.docx", "Contract-shared.docx");

foreach (OfficeHiddenDataFinding finding in report.Findings.Where(f => f.Outcome == OfficeHiddenDataOutcome.Retained)) {
    Console.WriteLine($"Kept {finding.Kind} at {finding.Location} ({finding.Reason}): {finding.Note}");
}
```

| Property | Meaning |
| --- | --- |
| `Kind` | The category |
| `Outcome` | `Found`, `Removed`, `Anonymized`, `Accepted`, `Rejected` or `Retained` |
| `Location` | Package part, sheet cell, slide, page or document story |
| `Description` | What was found |
| `Value` | A short preview, such as an author name or path |
| `Scope` | `Document`, `EmbeddedImage` for a picture inside it, or `Package` for the file as a whole |
| `Part` | The package part or entry the finding belongs to, such as `/word/media/image1.png` |
| `Reason` | Why the finding ended as it did, such as `NotSelected`, `NeededByDocument` or `ProtectsContentCredentials` |
| `Note` | The same reason as a sentence |
| `Provenance` | For provenance findings, the same evidence the provenance inspector returns |

`HasHiddenData` tells you whether anything was found, `Kinds` which categories were found, `Changed` whether removal changed the document, and `Format` which format reported it. Match on `Reason` rather than the `Note` text when your code needs to act on why something stayed.

To save the report as JSON, `OfficeHiddenDataReportSerializer.Serialize(report, sourceName, outputName)` in `OfficeIMO.Workflows` writes the `officeimo.hidden-data.result.v1` schema that the command line, the MCP tools and the browser tool also use.

## Signed documents

Removing hidden data changes the file, which breaks a digital signature. Saving a signed Word, Excel or PowerPoint document is blocked by default. Clean the document first and sign it afterwards. To clean a signed package anyway, pass `OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures` to `OfficeDocument.RemoveHiddenData`, or set `SignatureMutationPolicy` on `WordSaveOptions`, `ExcelSaveOptions` or `PowerPointSaveOptions`; the signatures the change would break are removed. Signed PDFs are always refused.

## Command line

```powershell
officeimo hidden-data inspect .\proposal.docx
officeimo hidden-data remove .\proposal.docx
officeimo hidden-data remove .\proposal.docx --output .\shared.docx --force --kinds default,provenance --tracked-changes accept
officeimo hidden-data remove .\poster.pdf --kinds default,provenance --keep-provenance c2pa
```

The first command reports everything without changing the file. The second writes `proposal.cleaned.docx` with the default categories. The third also removes origin records and accepts tracked changes, into a named file. The last removes only the AI label and keeps the Content Credentials.

`remove` never changes the input. It writes `<name>.cleaned<ext>` unless you pass `--output`, and refuses to replace an existing file without `--force`. `--kinds` takes `default`, `all` or category names such as `hidden-content`, `image-metadata`, `attachments`, `active-content`, `bookmarks` and `provenance`. `--keep-provenance` keeps some kinds of record when `provenance` is selected: `c2pa`, `external` or `ai-source`. `--anonymous-author` sets the replacement name, and `--remove-invalidated-signatures` allows cleaning a signed package. Output is JSON in the `officeimo.hidden-data.result.v1` schema, one entry per finding with its outcome; `--format text` prints a readable list instead.

## PowerShell

```powershell
Get-OfficeHiddenData -Path .\proposal.docx
Remove-OfficeHiddenData -Path .\proposal.docx
$options = New-OfficeHiddenDataOptions -Kind Default, Provenance -TrackedChanges AcceptAll
Get-ChildItem .\Outgoing -File | Remove-OfficeHiddenData -OutputDirectory .\Share -Options $options
```

`Remove-OfficeHiddenData` writes `<name>.cleaned<ext>` unless you pass `-OutputPath` or `-OutputDirectory`, replaces an existing output unless you pass `-NoClobber`, and returns the findings with their outcomes. [Clean files before sharing](/docs/pswriteoffice/before-sharing/) shows real output for Word, Excel, PowerPoint and PDF files.

## For AI agents

The [OfficeIMO MCP server](/agents/) has the same two operations:

- `officeimo_inspect_hidden_data` reports what would leave with a file and changes nothing.
- `officeimo_remove_hidden_data` writes a cleaned copy to `outputPath`. It takes `kinds`, `trackedChanges` (`keep`, `accept` or `reject`), `keepProvenance`, `anonymousAuthor`, `removeInvalidatedSignatures` and `overwrite`, with the same meanings as the command-line options.

## In the browser

[Clean a file before sharing](/browser/clean-before-sharing/) runs `InspectHiddenData` and `RemoveHiddenData` in your browser. It lists each category it found with the `Default` categories ticked, and saves a clean copy after checking it again. The file is not uploaded.
