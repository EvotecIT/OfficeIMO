---
title: "Remove hidden data and metadata"
description: "Inspect and remove document properties, comments, tracked changes, hidden content, speaker notes, custom XML and image GPS data from Word, Excel and PowerPoint files with OfficeIMO."
layout: docs
---

`InspectHiddenData()` lists the metadata and hidden data in a loaded Word, Excel or PowerPoint document and changes nothing. `RemoveHiddenData(options)` removes the selected categories in memory and returns the same report, with an outcome on every finding. Save the document to keep the result. The categories match what Office's Document Inspector checks.

## Word

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Load("Contract.docx");

OfficeHiddenDataReport found = document.InspectHiddenData();
Console.WriteLine($"Categories found: {found.Kinds}");

OfficeHiddenDataReport removed = document.RemoveHiddenData(new OfficeHiddenDataOptions {
    Kinds = OfficeHiddenDataKind.Default | OfficeHiddenDataKind.HiddenContent,
    TrackedChanges = OfficeTrackedChangesAction.AcceptAll,
    AnonymousAuthor = "Reviewer"
});
foreach (OfficeHiddenDataFinding finding in removed.Findings) {
    Console.WriteLine($"{finding.Outcome} {finding.Kind} at {finding.Location}: {finding.Description}");
}

document.Save("Contract-shared.docx");
```

Word reports properties, comments, tracked changes, reviewer names, hidden text, document variables, custom XML, revision save IDs (RSIDs), the attached template and image metadata. Comments and revisions that stay in the document keep an author, because Word requires one; `PersonalIdentities` replaces the name with `AnonymousAuthor` (`Author` by default). Bibliography sources and custom XML bound to content controls are kept.

## Excel

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

Excel reports notes, threaded comments and their authors, hidden sheets, rows, columns and names, external links, pivot cache source paths, custom XML, image metadata and printer settings. Cached values from external workbooks are removed, but the links themselves are reported and kept, because formulas and refreshes need them. Power Query data is kept and reported with a note.

Hidden sheets, rows and columns are removed only when you add `OfficeHiddenDataKind.HiddenContent`. Deleting them shifts or breaks formulas that refer to them, exactly as deleting them in Excel does.

## PowerPoint

```csharp
using OfficeIMO;
using OfficeIMO.PowerPoint;

using var presentation = PowerPointDocument.Load("Pitch.pptx");

OfficeHiddenDataReport report = presentation.RemoveHiddenData(new OfficeHiddenDataOptions {
    Kinds = OfficeHiddenDataKind.Default & ~OfficeHiddenDataKind.SpeakerNotes
});
Console.WriteLine($"Changed: {report.Changed}");

presentation.Save("Pitch-shared.pptx");
```

PowerPoint reports properties, classic and modern comments, speaker notes, the thumbnail, custom XML and image metadata. Hidden slides, hidden shapes and shapes placed entirely off the slide are removed only with `OfficeHiddenDataKind.HiddenContent`. The example keeps speaker notes by taking them out of `Default`.

## Choose categories

`OfficeHiddenDataOptions.Kinds` is a flags value. `Default` selects every category that leaves visible content alone; `All` adds tracked changes, hidden content and provenance.

| Category | In `Default` | Covers |
| --- | --- | --- |
| `DocumentProperties` | Yes | Author, last modified by, last printed, revision number, company, manager, template, editing time |
| `CustomProperties` | Yes | Custom document properties |
| `Comments` | Yes | Comments, replies, notes and threaded comments |
| `PersonalIdentities` | Yes | People parts, and authors on retained comments and revisions |
| `SpeakerNotes` | Yes | PowerPoint notes slides |
| `Thumbnail` | Yes | The package preview picture |
| `CustomXml` | Yes | Custom XML data parts |
| `ImageMetadata` | Yes | Exif, GPS, XMP and IPTC inside pictures; pixels are not re-encoded |
| `PrinterSettings` | Yes | Embedded printer settings |
| `ExternalReferences` | Yes | Attached templates and cached external values; links that content depends on are kept |
| `DocumentVariables` | Yes | Word document variables |
| `RevisionIdentifiers` | Yes | Word RSIDs |
| `TrackedChanges` | No | Reported; changed only through `TrackedChanges` |
| `HiddenContent` | No | Hidden text, sheets, rows, columns, names, slides and shapes |
| `Provenance` | No | Content Credentials and AI source labels; see [provenance](/docs/before-sharing/provenance/) |

The [hidden data support matrix](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/officeimo.hidden-data-support-matrix.md) lists exactly what each category covers in Word, Excel and PowerPoint.

## Pictures and provenance

Image metadata cleanup and provenance overlap: an AI-generated picture often says so in its XMP metadata, and Content Credentials are bound to the exact bytes of the image. While `Provenance` is not selected, `ImageMetadata` reduces a picture's XMP to its digital source type declaration instead of deleting it, and leaves pictures that carry Content Credentials untouched, reported as `Retained`. Add `OfficeHiddenDataKind.Provenance` to remove those records too.

## Read the report

```csharp
using OfficeIMO;
using OfficeIMO.Word;

using var document = WordDocument.Load("Contract.docx");
OfficeHiddenDataReport report = document.RemoveHiddenData();

foreach (OfficeHiddenDataFinding finding in report.Findings.Where(f => f.Outcome == OfficeHiddenDataOutcome.Retained)) {
    Console.WriteLine($"Kept {finding.Kind} at {finding.Location}: {finding.Note}");
}
```

| Property | Meaning |
| --- | --- |
| `Kind` | The category |
| `Location` | Package part, sheet cell, slide or document story |
| `Description` | What was found |
| `Value` | A short preview, such as an author name or path |
| `Outcome` | `Found`, `Removed`, `Anonymized`, `Accepted`, `Rejected` or `Retained` |
| `Note` | Why a finding was retained, or other detail about the outcome |
| `Provenance` | For provenance findings, the same evidence the provenance inspector returns |

`HasHiddenData` tells you whether anything was found, `Kinds` which categories were found, and `Changed` whether removal changed the document.

## Signed documents

Removing hidden data changes the package, which breaks a digital signature. Saving a signed document is blocked by default in Word, Excel and PowerPoint. Clean the document first and sign it afterwards.

## Command line

```powershell
# Report everything without changing the file
officeimo hidden-data inspect .\proposal.docx

# Write proposal.cleaned.docx with the default categories
officeimo hidden-data remove .\proposal.docx

# Also remove origin records and accept tracked changes, into a named file
officeimo hidden-data remove .\proposal.docx --output .\shared.docx --force --kinds default,provenance --tracked-changes accept
```

`remove` never changes the input. It writes `<name>.cleaned<ext>` unless you pass `--output`, and refuses to replace an existing file without `--force`. `--kinds` takes `default`, `all` or category names such as `hidden-content`, `image-metadata` and `provenance`; `--anonymous-author` sets the replacement name. Output is JSON in the `officeimo.hidden-data.result.v1` schema, one entry per finding with its outcome; `--format text` prints a readable list instead. Signed documents are refused.

## In the browser

[Clean a file before sharing](/browser/clean-before-sharing/) runs `InspectHiddenData` and `RemoveHiddenData` in your browser. It lists each category it found with the `Default` categories ticked, and saves a clean copy after checking it again. The file is not uploaded.
