---
title: Create, Read and Edit ODT, ODS and ODP Files
description: Create OpenDocument text documents, spreadsheets and presentations, read their content, edit them without rewriting untouched package entries, and open password-protected files.
order: 39
verified: 2026-11-14
commands:
  - officeimo read
packages:
  - OfficeIMO.OpenDocument
cmdlets:
  - New-OfficeOpenDocument
  - Open-OfficeOpenDocument
  - Add-OfficeOpenDocumentHeading
  - Add-OfficeOpenDocumentParagraph
  - Add-OfficeOpenDocumentSheet
  - Set-OfficeOpenDocumentCell
  - Add-OfficeOpenDocumentSlide
  - Add-OfficeOpenDocumentTextBox
related:
  - /docs/open-document/
  - /docs/open-document/convert-office/
  - /docs/open-document/convert-pdf/
  - /docs/security/sign-documents/
limits: Formula evaluation covers a bounded local subset of OpenFormula; external data, volatile functions and matrix formulas are not evaluated.
search_terms:
  - odt
  - ods
  - odp
  - libreoffice file
  - opendocument spreadsheet
  - opendocument presentation
---

`OfficeIMO.OpenDocument` reads and writes OpenDocument packages directly, with no LibreOffice or Office process. Use `OdtDocument` for text documents, `OdsDocument` for spreadsheets and `OdpPresentation` for presentations. Each also reads and writes the flat XML variants (`.fodt`, `.fods`, `.fodp`).

## Create an ODT text document

### C#

```csharp
using OfficeIMO.OpenDocument;

using OdtDocument document = OdtDocument.Create();
document.AddHeading("Service report", 1);
document.AddParagraph("All services are available.");

OdtParagraph emphasis = document.AddParagraph();
emphasis.AddRun("Read ").Bold = true;
emphasis.AddHyperlink("the runbook", "https://example.com/runbook");

OdtTable table = document.AddTable(3, 2, "Results");
table.Cell(0, 0).Text = "Service";
table.Cell(0, 1).Text = "Status";
table.Cell(1, 0).Text = "Web";
table.Cell(1, 1).Text = "Healthy";
table.Cell(2, 0).Text = "Database";
table.Cell(2, 1).Text = "Healthy";

OdtParagraph footer = document.AddParagraph("Page ");
footer.AddField(OdtFieldKind.PageNumber, "1");
footer.AddFootnote("Collected by the nightly check.");

document.Save("report.odt");
```

### PowerShell

```powershell
New-OfficeOpenDocument -Path ./report.odt {
    Add-OfficeOpenDocumentHeading -Text 'Service report' -Level 1
    Add-OfficeOpenDocumentParagraph -Text 'All services are available.' -Bold
}
```

## Create an ODS spreadsheet

Cells are zero-based and typed: `SetString`, `SetNumber`, `SetDecimal`, `SetBoolean`, `SetDate` and `SetPercentage` write the matching OpenDocument value type. Formulas use OpenFormula syntax, and `Recalculate` fills their cached values.

### C#

```csharp
using OfficeIMO.OpenDocument;

using OdsDocument workbook = OdsDocument.Create();
OdsSheet sheet = workbook.AddSheet("Budget");
sheet.Cell(0, 0).SetString("Item");
sheet.Cell(0, 1).SetString("Amount");
sheet.Cell(1, 0).SetString("Hosting");
sheet.Cell(1, 1).SetDecimal(1200m);
sheet.Cell(2, 0).SetString("Licenses");
sheet.Cell(2, 1).SetDecimal(850m);
sheet.Cell(3, 0).SetString("Total");
sheet.Cell(3, 1).Formula = "of:=SUM([.B2:.B3])";

OdsRecalculationReport calculation = workbook.Recalculate();
Console.WriteLine($"{calculation.FailedCells} formulas failed");

workbook.Save("budget.ods");
```

`AddValidation`, `AddChart` and named ranges are also available; see the [package README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.OpenDocument/README.md) for the full spreadsheet surface.

### PowerShell

```powershell
New-OfficeOpenDocument -Path ./budget.ods {
    Add-OfficeOpenDocumentSheet -Name 'Budget' -Content {
        Set-OfficeOpenDocumentCell -Row 0 -Column 0 -Value 'Item'
        Set-OfficeOpenDocumentCell -Row 0 -Column 1 -Value 'Amount'
        Set-OfficeOpenDocumentCell -Row 1 -Column 0 -Value 'Hosting'
        Set-OfficeOpenDocumentCell -Row 1 -Column 1 -Value 1200
    }
}
```

## Create an ODP presentation

Positions and sizes use `OdfRect`, here in centimeters.

### C#

```csharp
using OfficeIMO.OpenDocument;

using OdpPresentation presentation = OdpPresentation.Create();
OdpSlide slide = presentation.AddSlide("Summary");
slide.AddTextBox(OdfRect.FromCentimeters(2, 1, 24, 3), "Quarterly summary");
slide.AddTextBox(OdfRect.FromCentimeters(2, 5, 24, 8), "Revenue grew 12% over the previous quarter.");
slide.AddRectangle(OdfRect.FromCentimeters(2, 14, 8, 2)).FillColor = OdfColor.Parse("#D1E9FF");
slide.GetOrCreateSpeakerNotes().AddParagraph("Mention the new region.");

presentation.Save("summary.odp");
```

### PowerShell

```powershell
New-OfficeOpenDocument -Path ./summary.odp {
    Add-OfficeOpenDocumentSlide -Name 'Summary' -Content {
        Add-OfficeOpenDocumentTextBox -Text 'Quarterly summary' -X 2cm -Y 1cm -Width 24cm -Height 3cm
    }
}
```

## Read an existing file

### C#

```csharp
using OfficeIMO.OpenDocument;

using (OdtDocument text = OdtDocument.Load("report.odt")) {
    foreach (OdtParagraph paragraph in text.Paragraphs) Console.WriteLine(paragraph.Text);
    foreach (OdtTable table in text.Tables) Console.WriteLine($"table with {table.Rows.Count} rows");
}

using (OdsDocument workbook = OdsDocument.Load("budget.ods")) {
    foreach (OdsSheet sheet in workbook.Sheets) {
        if (sheet.UsedRange is not OdsUsedRange used) continue;
        for (long row = used.FirstRow; row <= used.LastRow; row++) {
            var cells = new List<string>();
            for (long column = used.FirstColumn; column <= used.LastColumn; column++) {
                cells.Add(sheet.Cell(row, column).Text);
            }
            Console.WriteLine($"{sheet.Name}: {string.Join(" | ", cells)}");
        }
    }
}

using (OdpPresentation presentation = OdpPresentation.Load("summary.odp")) {
    foreach (OdpSlide slide in presentation.Slides) {
        foreach (OdpTextBox box in slide.Shapes.OfType<OdpTextBox>()) {
            Console.WriteLine($"{slide.Name}: {string.Join(" ", box.Paragraphs.Select(p => p.Text))}");
        }
    }
}
```

When you do not know the kind in advance, `OdfDocument.Load(path)` returns the matching `OdtDocument`, `OdsDocument`, `OdpPresentation` or `OdgDocument`. Loading is bounded: `OdfLoadOptions` limits package size, entry count, expansion ratio and XML depth, and the library never runs scripts, macros or external links.

### CLI

```sh
officeimo read report.odt
officeimo read budget.ods --format json --output budget.json
```

## Edit a file and keep the rest of the package

Typed objects stay backed by the source XML. An edit rewrites only its owning XML part; untouched entries, unknown XML and vendor extensions keep their original bytes. The save report lists what was rewritten and what could not be carried.

### C#

```csharp
using OfficeIMO.OpenDocument;

using OdtDocument document = OdtDocument.Load("input.odt");
document.Paragraphs[0].Text = "Updated title";
document.AddParagraph("Reviewed and approved.");

OdfSaveResult result = document.Save("output.odt", new OdfSaveOptions {
    CompatibilityProfile = OdfCompatibilityProfile.PreserveSource
});

foreach (string entry in result.Report.RewrittenEntries) Console.WriteLine("rewritten " + entry);
foreach (string entry in result.Report.LossyEntries) Console.WriteLine("lossy " + entry);
```

New documents use ODF 1.4. Set `CompatibilityProfile = OdfCompatibilityProfile.Odf13` when a consumer needs ODF 1.3. Changing a signed package fails by default, because saving would invalidate its signatures.

### PowerShell

```powershell
Open-OfficeOpenDocument -Path ./input.odt -Content {
    Add-OfficeOpenDocumentParagraph -Text 'Reviewed and approved.'
}
```

## Open and save password-protected files

ODF password encryption is built in and does not need `OfficeIMO.Security`. Output uses AES-256-CBC with per-entry PBKDF2 keys. The password is used only for the current load or save and is not kept.

### C#

```csharp
using OfficeIMO.OpenDocument;

string password = Environment.GetEnvironmentVariable("ODF_PASSWORD") ?? throw new InvalidOperationException("Set ODF_PASSWORD.");

using OdtDocument document = OdtDocument.Load("protected.odt", new OdfLoadOptions { Password = password });
document.AddParagraph("Updated while decrypted in memory.");
document.Save("protected-updated.odt", new OdfSaveOptions {
    Encryption = new OdfEncryptionOptions { Password = password }
});
```

A missing or wrong password throws `OdfEncryptedPackageException` with a `Reason`. Saving an encrypted source without `Encryption` also fails, so protection is never removed by accident; set `EncryptionHandling = OdfEncryptionHandling.Remove` to write plaintext on purpose.

### PowerShell

```powershell
$document = Open-OfficeOpenDocument -Path ./protected.odt -Password (Read-Host -AsSecureString)
$document.Paragraphs[0].Text
Close-OfficeOpenDocument -Document $document
```
