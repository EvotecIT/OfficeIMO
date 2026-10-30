# OfficeIMO.Pdf

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Pdf)](https://www.nuget.org/packages/OfficeIMO.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Pdf)

Create, read, edit and combine PDF files in .NET. `OfficeIMO.Pdf` is OfficeIMO's own PDF engine, with no third-party PDF library, browser or Office installation. It also fills forms, stamps and redacts pages, protects files with passwords and checks PDF/A and PDF/UA output.

## Install

```sh
dotnet add package OfficeIMO.Pdf
```

Optional packages:

```sh
dotnet add package OfficeIMO.Security       # certificate signatures
dotnet add package OfficeIMO.Pdf.Ocr        # searchable scans, with an OCR provider
dotnet add package OfficeIMO.Pdf.Word       # PDF to Word; also .Pdf.Excel and .Pdf.PowerPoint
```

To turn Word, Excel, PowerPoint or Markdown into PDF, install the source format's adapter (`OfficeIMO.Word.Pdf`, `OfficeIMO.Excel.Pdf`, `OfficeIMO.PowerPoint.Pdf`, `OfficeIMO.Markdown.Pdf`) and call `SaveAsPdf` on that document.

Targets `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows.

## Create a report

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Service review")
        .Paragraph(p => p
            .Text("All scheduled checks passed. ")
            .Bold("Two follow-ups remain open."))
        .Table(new[] {
            new[] { "Task", "Owner", "Due" },
            new[] { "Rotate certificates", "Platform", "4 Nov" },
            new[] { "Review backups", "Operations", "11 Nov" }
        })))
    .Meta(title: "Service review", author: "Operations")
    .Save("Service-review.pdf");
```

Page margins follow the `Modern` theme unless `PdfCreateOptions.Theme` names another. `PdfCreateOptions` also sets the page size, fonts, tagging and compliance profile.

## Read text and tables

```csharp
using OfficeIMO.Pdf;

PdfDocument pdf = PdfDocument.Load("Service-review.pdf");
PdfDocumentReadResult result = pdf.Read();

foreach (PdfLogicalPage page in result.Pages) {
    Console.WriteLine($"Page {page.PageNumber}: {page.Paragraphs.Count} paragraphs, {page.Tables.Count} tables");
}
Console.WriteLine(result.ExtractMarkdown());
```

For files from outside your organization, pass `PdfLoadOptions` with read limits, as shown in the [guide](../Docs/officeimo.pdf-guide.md).

## Combine, split and reorder pages

```csharp
using OfficeIMO.Pdf;

PdfDocument.Load("Service-review.pdf")
    .MergeWith("Service-review.pdf")
    .Pages.Rotate(90, "2")
    .UpdateMetadata(title: "Review pack")
    .Save("Review-pack.pdf");

PdfDocument pack = PdfDocument.Load("Review-pack.pdf");
pack.Pages.Extract("1").Save("Cover.pdf");
IReadOnlyList<PdfDocument> pages = pack.Pages.Split();
for (int i = 0; i < pages.Count; i++) pages[i].Save($"Review-pack-{i + 1}.pdf");
```

Page operations return a new document; the source file is never changed in place. Page numbers and ranges are 1-based.

## Fill a form

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Application")
        .TextField("Applicant.Name")
        .TextField("Approval.Status")))
    .Save("Application.pdf");

PdfDocument.Load("Application.pdf")
    .Forms.FillAndFlatten(new Dictionary<string, string> {
        ["Applicant.Name"] = "Adele Vance",
        ["Approval.Status"] = "Approved"
    })
    .Save("Application-filled.pdf");
```

`Forms.Fill` keeps the fields editable; `FillAndFlatten` makes the values part of the page.

## Protect with a password

```csharp
using OfficeIMO.Pdf;

PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Salaries 2027")
        .Paragraph(p => p.Text("Confidential."))),
        new PdfCreateOptions().SetEncryption("open-password", ownerPassword: "owner-password"))
    .Save("Salaries.pdf");
```

## Convert a PDF back to Word

```csharp
using OfficeIMO;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Word;

PdfDocument pdf = PdfDocument.Load("Service-review.pdf");
OfficeConversionReport report = pdf.SaveAsWord("Service-review.docx");
Console.WriteLine($"Loss: {report.HasLoss}");
```

Reconstruction from a fixed-layout PDF is semantic: paragraphs, headings, lists and tables come back as editable Word content, and the report lists what could not be recovered.

## More

- [PDF guide](../Docs/officeimo.pdf-guide.md): page layout and typography, fonts, headers and footers, images, stamps and watermarks, redaction, annotations, signatures, attachments, tagged output, PDF/A and PDF/X, page images, invoices and inspection.
- [PDF current state](../Docs/officeimo.pdf.current-state.md), [PDF conversion support](../Docs/officeimo.pdf-conversion-support-matrix.md) and the [generated operation contract](../Docs/Compatibility/generated/package-operations.md).
- [Examples](../OfficeIMO.Examples/Pdf) and the [migration guide](../MIGRATION.md).
- PowerShell: [PSWriteOffice](../PowerShell/PSWriteOffice/README.MD) (`New-OfficePdf`, `Open-OfficePdf`).

Dependencies: `OfficeIMO.Core` only; the PDF parser, writer, renderer and password encryption are OfficeIMO code. MIT licensed.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 2 | 0 | 0 | 0 | 0 | 0 |
| Read | 2 | 0 | 0 | 0 | 0 | 0 |
| Edit | 2 | 0 | 0 | 0 | 0 | 0 |
| Preserve | 0 | 1 | 0 | 0 | 0 | 0 |
| Inspect | 3 | 0 | 0 | 0 | 0 | 0 |
| Validate | 2 | 0 | 0 | 0 | 0 | 1 |
| Remove | 2 | 0 | 0 | 0 | 0 | 0 |
| Export | 5 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
