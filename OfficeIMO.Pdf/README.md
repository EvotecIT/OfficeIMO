# OfficeIMO.Pdf

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Pdf)](https://www.nuget.org/packages/OfficeIMO.Pdf)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Pdf?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Pdf)

Create, read and change PDF files in .NET with OfficeIMO's own PDF engine — no third-party PDF library, browser or Office installation.

## What it does

- **Create** reports, invoices and forms from headings, paragraphs, tables, lists, images, drawings, headers and footers, with embedded fonts.
- **Read** text, tables, images, links, attachments and form values, or export the content as Markdown or JSON.
- **Merge, split and organize** pages: extract, reorder, rotate, crop, duplicate, delete, stamp and watermark.
- **Fill forms** and keep them editable, or flatten the values into the page.
- **Protect** files with AES-256 passwords and permissions, or remove protection with the owner password.
- **Sign** with certificate-based signatures (PAdES B-B and B-T) and validate existing signatures.
- **Redact** text and areas permanently, with a verification report for every removed match.
- **Remove hidden data** before sharing: metadata, comments, scripts and other active content, attachments, bookmarks and layers, with the same `InspectHiddenData` and `RemoveHiddenData` report as the Office formats.
- **Produce standard output**: PDF/A, PDF/UA, PDF/X, Factur-X and ZUGFeRD, checked against the saved bytes.
- **Render** pages to PNG, JPEG, TIFF, WebP or SVG, and compare two PDFs page by page.
- **Convert** Word, Excel, PowerPoint, HTML and Markdown to PDF, and PDF back to Word, Excel, PowerPoint, HTML or Markdown.

## Install

```sh
dotnet add package OfficeIMO.Pdf
```

`OfficeIMO.Pdf` depends only on `OfficeIMO.Core`. It targets `netstandard2.0`, `net8.0` and `net10.0`, plus `net472` on Windows. Add an optional package only when you need its workflow:

| To | Add |
| --- | --- |
| Sign with certificates or validate CMS signatures | `OfficeIMO.Security` |
| Save Word, Excel or PowerPoint as PDF | `OfficeIMO.Word.Pdf`, `OfficeIMO.Excel.Pdf`, `OfficeIMO.PowerPoint.Pdf` |
| Convert HTML to PDF, or PDF to HTML | `OfficeIMO.Html.Pdf` |
| Save Markdown as PDF | `OfficeIMO.Markdown.Pdf` |
| Convert PDF to Word, Excel or PowerPoint | `OfficeIMO.Pdf.Word`, `OfficeIMO.Pdf.Excel`, `OfficeIMO.Pdf.PowerPoint` |
| Read scanned pages with an OCR provider | `OfficeIMO.Pdf.Ocr` |
| Render an invoice with its Factur-X XML | `OfficeIMO.Invoicing.Pdf` |
| Full OpenType shaping for complex scripts | `OfficeIMO.Drawing.HarfBuzz` |

## Examples

Each example uses `using OfficeIMO.Pdf;`, and page numbers and ranges are 1-based. Page, merge and form operations return a new document; save it to a different path to keep the source file.

### Create a report

<!-- powerforge:code source="Build/DocumentationExamples/Pdf/Program.cs" region="Create" -->
```csharp
using OfficeIMO.Pdf;

PdfDocument.Create(pdf => pdf.Content(content => content
        .H1("Service review")
        .AddParagraph("All scheduled checks passed. Two follow-ups remain open.")
        .AddTable(new[] {
            new[] { "Task", "Owner", "Due" },
            new[] { "Rotate certificates", "Platform", "4 Nov" } })))
    .Meta(title: "Service review", author: "Operations")
    .Save("Service-review.pdf");
```
<!-- /powerforge:code -->

### Read text and tables

<!-- powerforge:code source="Build/DocumentationExamples/Pdf/Program.cs" region="Read" -->
```csharp
PdfDocumentReadResult result = PdfDocument.Load("Service-review.pdf").Read();
foreach (PdfLogicalPage page in result.Pages) {
    Console.WriteLine($"Page {page.PageNumber}: {page.Paragraphs.Count} paragraphs, {page.Tables.Count} tables");
}
Console.WriteLine(result.ExtractMarkdown());
```
<!-- /powerforge:code -->

`Read()` returns paragraphs, headings, lists and tables per page. The [reading guide](../Docs/officeimo.pdf-guide.md#load-once-and-build-one-semantic-read-result) covers decoding diagnostics, form values and scanned pages; set [read limits](../Docs/officeimo.pdf-guide.md#deploy) for files from outside your organization.

### Merge, rotate and split

This example combines the report with an `Appendix.pdf` of your own.

<!-- powerforge:code source="Build/DocumentationExamples/Pdf/Program.cs" region="Combine" -->
```csharp
PdfDocument.Load("Service-review.pdf")
    .MergeWith("Appendix.pdf")
    .Pages.Rotate(90, "2")
    .Save("Review-pack.pdf");

IReadOnlyList<PdfDocument> parts = PdfDocument.Load("Review-pack.pdf").Pages.Split();
for (int i = 0; i < parts.Count; i++) parts[i].Save($"Review-pack-{i + 1}.pdf");
```
<!-- /powerforge:code -->

### Fill a form

This example fills `Application.pdf`, a form with `Applicant.Name` and `Approval.Status` text fields. The [forms guide](../Docs/officeimo.pdf-guide.md#forms) shows how to create one.

<!-- powerforge:code source="Build/DocumentationExamples/Pdf/Program.cs" region="FillForm" -->
```csharp
PdfDocument.Load("Application.pdf")
    .Forms.FillAndFlatten(new Dictionary<string, string> {
        ["Applicant.Name"] = "Adele Vance",
        ["Approval.Status"] = "Approved"
    })
    .Save("Application-filled.pdf");
```
<!-- /powerforge:code -->

Use `Forms.Fill` to keep the fields editable.

### Protect with a password

<!-- powerforge:code source="Build/DocumentationExamples/Pdf/Program.cs" region="Protect" -->
```csharp
PdfSecurityMutationResult protectedPdf = PdfDocument.Load("Service-review.pdf")
    .Security.Encrypt(new PdfStandardEncryptionOptions("open-password") {
        OwnerPassword = "owner-password"
    });
File.WriteAllBytes("Service-review-protected.pdf", protectedPdf.Pdf);
```
<!-- /powerforge:code -->

### Render a page to PNG

<!-- powerforge:code source="Build/DocumentationExamples/Pdf/Program.cs" region="Render" -->
```csharp
PdfPageRenderResult preview = PdfDocument.Load("Service-review.pdf")
    .Render.Pages("1", new PdfPageRenderOptions { Dpi = 144 })[0];
File.WriteAllBytes("Service-review-page-1.png", preview.Bytes!);
```
<!-- /powerforge:code -->

## More

- [PDF guide](../Docs/officeimo.pdf-guide.md): every workflow with examples, from page layout and fonts to signatures, redaction, PDF/A and deployment.
- [PDF current state](../Docs/officeimo.pdf.current-state.md): what each workflow covers and where its limits are.
- [PDF conversion support](../Docs/officeimo.pdf-conversion-support-matrix.md) and the [generated operation contract](../Docs/Compatibility/generated/package-operations.md).
- [Examples](../OfficeIMO.Examples/Pdf), the [migration guide](../MIGRATION.md) and [PSWriteOffice](../PowerShell/PSWriteOffice/README.MD) for PowerShell (`New-OfficePdf`, `Open-OfficePdf`).

MIT licensed.

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
| Convert | 0 | 2 | 0 | 0 | 0 | 0 |
| Export | 5 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Pdf` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
