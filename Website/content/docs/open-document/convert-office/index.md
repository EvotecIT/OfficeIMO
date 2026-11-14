---
title: Convert OpenDocument to and from Word, Excel and PowerPoint
description: Convert ODT to DOCX, ODS to XLSX and ODP to PPTX and back with the OfficeIMO OpenDocument adapters, and read the per-feature report of what was mapped, approximated or lost.
order: 40
verified: 2026-11-14
commands:
  - officeimo convert
packages:
  - OfficeIMO.Word.OpenDocument
  - OfficeIMO.Excel.OpenDocument
  - OfficeIMO.PowerPoint.OpenDocument
cmdlets:
  - Convert-OfficeDocument
  - New-OfficeWordOpenDocumentOptions
  - New-OfficeExcelOpenDocumentOptions
  - New-OfficePowerPointOpenDocumentOptions
related:
  - /docs/open-document/
  - /docs/open-document/create-read/
  - /docs/open-document/convert-pdf/
  - /docs/cli/convert/
limits: Conversion maps a documented feature subset; tracked changes, charts, SmartArt, content controls and section-specific layout are reported as loss rather than carried over.
search_terms:
  - odt to docx
  - docx to odt
  - ods to xlsx
  - xlsx to ods
  - odp to pptx
  - pptx to odp
---

Each Office format has its own adapter package, so a Word-only application does not carry the Excel or PowerPoint engines. Conversion is explicit and returns an `OdfConversionReport` that lists every feature as converted, approximated, skipped or unsupported, so a pipeline can reject unexpected changes instead of trusting a successful save.

| From | To | Package | Methods |
|---|---|---|---|
| `WordDocument` | ODT | `OfficeIMO.Word.OpenDocument` | `ToOpenDocument`, `SaveAsOpenDocument` |
| `OdtDocument` | DOCX | `OfficeIMO.Word.OpenDocument` | `ToWord`, `SaveAsWord` |
| `ExcelDocument` | ODS | `OfficeIMO.Excel.OpenDocument` | `ToOpenDocument`, `SaveAsOpenDocument` |
| `OdsDocument` | XLSX | `OfficeIMO.Excel.OpenDocument` | `ToExcel`, `SaveAsExcel` |
| `PowerPointDocument` | ODP | `OfficeIMO.PowerPoint.OpenDocument` | `ToOpenDocument`, `SaveAsOpenDocument` |
| `OdpPresentation` | PPTX | `OfficeIMO.PowerPoint.OpenDocument` | `ToPowerPoint`, `SaveAsPowerPoint` |

Every method has an `Async` form and a stream overload.

## Convert Word to ODT and back

### C#

```csharp
using OfficeIMO;
using OfficeIMO.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;

using (WordDocument word = WordDocument.Load("report.docx")) {
    using OdtDocument odt = word.ToOpenDocument(null, out OfficeConversionReport report);
    odt.Save("report.odt");

    foreach (OdfConversionMapping mapping in ((OdfConversionReport)report).Mappings) {
        if (mapping.Status != OdfConversionMappingStatus.Converted) {
            Console.WriteLine($"{mapping.Feature}: {mapping.Status} ({mapping.Count}) {mapping.Message}");
        }
    }
}

using (OdtDocument source = OdtDocument.Load("proposal.odt")) {
    OfficeConversionReport saved = source.SaveAsWord("proposal.docx");
    Console.WriteLine(saved.HasLoss ? "Converted with loss" : "Converted");
}
```

The Word adapter maps headings, paragraphs and their formatting, lists, tables and merges, inline images, hyperlinks, bookmarks, page layout, default and first-page headers and footers, footnotes and endnotes, and page number, page count, date and time fields.

### PowerShell

```powershell
Convert-OfficeDocument -Path ./report.docx -OutputPath ./report.odt
Convert-OfficeDocument -Path ./proposal.odt -OutputPath ./proposal.docx -FailOnLoss

$options = New-OfficeWordOpenDocumentOptions -NoImages -NoHeadersAndFooters
Convert-OfficeDocument -Path ./report.docx -To OpenDocument -OutputPath ./report.odt -Options $options
```

### CLI

```sh
officeimo convert report.docx report.odt
officeimo convert proposal.odt proposal.docx --require-no-loss
```

## Convert Excel to ODS and back

### C#

```csharp
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;

using (ExcelDocument workbook = ExcelDocument.Load("report.xlsx")) {
    OfficeConversionReport report = workbook.SaveAsOpenDocument("report.ods", new ExcelOpenDocumentConversionOptions {
        LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported
    });
    Console.WriteLine($"{report.Diagnostics.Count} diagnostics");
}

using (OdsDocument source = OdsDocument.Load("budget.ods"))
using (ExcelDocument converted = source.ToExcel()) {
    converted.Save("budget.xlsx");
}
```

The Excel adapter maps worksheets, typed values, formulas (translated structurally between A1 and OpenFormula syntax), hyperlinks, comments, merges, row and column layout, named ranges, number formats, a basic style subset, common validations, simple charts and pivot tables. `ExcelOpenDocumentConversionOptions` bounds rows, columns and converted cells; content outside those limits is reported as `Skipped`.

### PowerShell

```powershell
$options = New-OfficeExcelOpenDocumentOptions -IncludeBasicStyles -MaximumRows 10000 -MaximumColumns 100
Convert-OfficeDocument -Path ./report.xlsx -OutputPath ./report.ods -Options $options
Convert-OfficeDocument -Path ./budget.ods -OutputPath ./budget.xlsx
```

### CLI

```sh
officeimo convert report.xlsx report.ods
officeimo convert budget.ods budget.xlsx
```

## Convert PowerPoint to ODP and back

### C#

```csharp
using OfficeIMO;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;

using (PowerPointDocument deck = PowerPointDocument.Load("deck.pptx")) {
    deck.SaveAsOpenDocument("deck.odp");
}

using (OdpPresentation source = OdpPresentation.Load("slides.odp")) {
    using PowerPointDocument converted = source.ToPowerPoint(null, out OfficeConversionReport report);
    converted.Save("slides.pptx");
    report.ThrowIfLoss();   // fail the job if anything was approximated or dropped
}
```

The PowerPoint adapter maps slide size and order, hidden slides, text boxes and runs, images and crops, tables, basic shapes, solid backgrounds, common transitions, speaker notes, and master and layout associations with common placeholder roles.

### PowerShell

```powershell
$options = New-OfficePowerPointOpenDocumentOptions -NoSpeakerNotes
Convert-OfficeDocument -Path ./deck.pptx -To OpenDocument -Options $options
```

### CLI

```sh
officeimo convert deck.pptx deck.odp
officeimo convert slides.odp slides.pptx
```

## Reject conversions that lose content

Set `LossPolicy` on the adapter options to make a conversion throw instead of returning a lossy result. `ThrowOnSkippedOrUnsupported` accepts documented approximations; `ThrowOnAnyLoss` rejects those too.

### C#

```csharp
using OfficeIMO.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;

var options = new WordOpenDocumentConversionOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnAnyLoss };

using OdtDocument source = OdtDocument.Load("contract.odt");
try {
    using WordDocument word = source.ToWord(options);
    word.Save("contract.docx");
} catch (Exception exception) {
    Console.WriteLine("Not converted: " + exception.Message);
}
```

### PowerShell

```powershell
Convert-OfficeDocument -Path ./contract.odt -OutputPath ./contract.docx -FailOnLoss
```

### CLI

```sh
officeimo convert contract.odt contract.docx --require-no-loss
```
