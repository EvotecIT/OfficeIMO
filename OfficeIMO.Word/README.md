# OfficeIMO.Word - Word documents for .NET

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Word)](https://www.nuget.org/packages/OfficeIMO.Word)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Word?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Word)

`OfficeIMO.Word` is the main Word document package in the OfficeIMO family. It creates, edits, inspects, converts, and saves `.docx` files, and can import and write supported legacy `.doc` files, without COM automation and without Microsoft Office installed.

If OfficeIMO saves you time, please consider supporting the work through [GitHub Sponsors](https://github.com/sponsors/PrzemyslawKlys) or [PayPal](https://paypal.me/PrzemyslawKlys). PowerShell users should use [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice) for the PowerShell-facing experience.

## Install

```powershell
dotnet add package OfficeIMO.Word
```

Upgrading from OfficeIMO 3.0? The regular create, load, edit, and save workflow
keeps the same shape. Start with the [short Word-first 3.1 migration path](../MIGRATION.md#start-here-most-officeimoword-applications) for the package, enum, and type replacements most Word applications need.

Document authoring, reading, signature inspection, and safe signed-package handling do not require the security
package. Install it only when the application creates or cryptographically validates OPC or VBA signatures:

```powershell
dotnet add package OfficeIMO.Security
```

## Quick start

```csharp
using OfficeIMO.Word;

using var document = WordDocument.Create("report.docx");

document.AddParagraph("Quarterly report").Style = WordParagraphStyles.Heading1;
document.AddParagraph("Created with OfficeIMO.Word.");

var table = document.AddTable(2, 2, WordTableStyle.TableGrid);
table.Rows[0].Cells[0].Paragraphs[0].Text = "Area";
table.Rows[0].Cells[1].Paragraphs[0].Text = "Status";
table.Rows[1].Cells[0].Paragraphs[0].Text = "Documents";
table.Rows[1].Cells[1].Paragraphs[0].Text = "Generated";
table.RepeatHeaderRowAtTheTopOfEachPage = true;
table.Style = WordTableStyle.TableGrid;

document.Save();
```

`AsFluent()` wraps the same `WordDocument`; `End()` returns that document for
direct object-model work:

```csharp
using var document = WordDocument.Create("report.docx");

document.AsFluent()
    .H1("Quarterly report")
    .Paragraph(paragraph => paragraph.Text("Created with OfficeIMO.Word."))
    .End();

document.Save();
```

## What it does

- Creates, loads, edits, saves, and appends `.docx` documents.
- Opens supported Word 97-2003 `.doc` files through the normal `WordDocument.Load(...)` path and projects them into the regular OfficeIMO Word model.
- Writes native `.doc` files for the currently supported simple-document subset, with preflight checks that block unsupported content before saving.
- Converts supported `.doc` and `.docx` files with `WordDocument.Convert(...)`, using the same import diagnostics and save preflight as normal load/save workflows.
- Works with paragraphs, runs, styles, sections, headers, footers, page numbers, tables, images, hyperlinks, bookmarks, fields, footnotes, endnotes, content controls, charts, shapes, and document protection.
- Applies optional shared package-security policy before parsing Open XML or compound DOC files, and preflights read, edit, template, render, and save capabilities.
- Inspects and manages VBA and embedded package/OLE/ActiveX payload bytes without executing active content.
- Creates and validates cross-platform OPC XML package signatures and managed VBA legacy, agile, and V3 signatures through an explicitly supplied `IOfficeSecurityProvider`; structural inspection remains provider-free.
- Exports estimated document page ranges as dependency-free PNG or SVG previews through `ExportImages(...)`, `SaveAsImages(...)`, and `ToImages()`.
- Keeps Office automation out of the runtime path, making it suitable for services, scheduled jobs, CI, desktop apps, and automation hosts.
- Provides fluent helpers for common authoring flows while keeping the lower-level Word object model available.
- Uses `OfficeIMO.Drawing` for shared colors, image metadata, page rendering, and the reusable math expression tree.

Advanced drawing, structured comparison, field evaluation, and evidence boundaries are documented in [Word advanced editing and evidence contracts](../Docs/officeimo.word-advanced-contracts.md). The contracts distinguish persisted drawing geometry from desktop Word layout, detected relocation from native move revisions, and supported legacy DOC writing from arbitrary DOC authoring.

For untrusted files, capability preflight, binary DOC/XLS/XLSB loss policies,
macro and embedded-payload handling, and the executable compatibility corpus,
see the [Word and Excel interoperability guide](../Docs/officeimo.word-excel-interoperability.md).

## Examples

The quick start shows the smallest useful document. These examples show the kinds of document work that belong in `OfficeIMO.Word` itself.

### Paragraphs and runs

```csharp
var paragraph = document.AddParagraph("Status: ");
paragraph.AddText("Approved").Bold = true;
paragraph.AddText(" on ");
paragraph.AddText(DateTime.Today.ToString("yyyy-MM-dd")).Italic = true;
```

### Tables with structure

```csharp
var table = document.AddTable(3, 3);
table.Rows[0].Cells[0].Paragraphs[0].Text = "Area";
table.Rows[0].Cells[1].Paragraphs[0].Text = "Owner";
table.Rows[0].Cells[2].Paragraphs[0].Text = "Status";
table.RepeatHeaderRowAtTheTopOfEachPage = true;
table.Style = WordTableStyle.TableGrid;

table.Rows[1].Cells[0].Paragraphs[0].Text = "Documents";
table.Rows[1].Cells[1].Paragraphs[0].Text = "Operations";
table.Rows[1].Cells[2].Paragraphs[0].Text = "Ready";

table.MergeCells(rowIndex: 2, columnIndex: 0, rowSpan: 1, colSpan: 3);
table.Rows[2].Cells[0].Paragraphs[0].Text = "Generated by OfficeIMO.Word";
```

### Headers and footers

```csharp
document.HeaderDefaultOrCreate.AddParagraph("Internal report");
document.FooterDefaultOrCreate.AddParagraph()
    .AddText("Page ")
    .AddPageNumber();
```

### Images

```csharp
var paragraph = document.AddParagraph();
paragraph.AddImage("logo.png", width: 160, height: 64);
```

### Hyperlinks and bookmarks

```csharp
document.AddParagraph("Jump target").AddBookmark("target-section");
document.AddParagraph()
    .AddHyperLink("Open project site", new Uri("https://github.com/EvotecIT/OfficeIMO"));
document.AddParagraph()
    .AddHyperLink("Jump inside document", "target-section", addStyle: true);
```

### Fields and table of contents

```csharp
document.AddParagraph("Chapter 1").Style = WordParagraphStyles.Heading1;
document.AddParagraph("Section 1.1").Style = WordParagraphStyles.Heading2;
document.Paragraphs[0].AddField(WordFieldType.TOC);
```

### Plain DOCX templates

Use ordinary `{{Name}}` placeholders when a Word-authored layout should bind directly to an application model. Scalar placeholders retain the formatting of the first template run, while repeated and conditional marker paragraphs can surround paragraphs, lists, or tables.

```csharp
using var document = WordDocument.Load("invoice-template.docx");

var values = new Dictionary<string, object?> {
    ["Customer"] = new Dictionary<string, object?> { ["Name"] = "Northwind Traders" },
    ["Lines"] = new object[] {
        new Dictionary<string, object?> { ["Description"] = "Assessment", ["Amount"] = 1200m },
        new Dictionary<string, object?> { ["Description"] = "Implementation", ["Amount"] = 3400m }
    },
    ["Portal"] = new WordTemplateHyperlink("Open invoice", new Uri("https://example.com/invoices/42"))
};

WordTemplateResult result = WordTemplate.Apply(document, values).EnsureComplete();
document.Save("invoice-42.docx");
```

Inside the DOCX, use `{{Customer.Name}}` for values and put block markers on their own paragraphs:

```text
{{#each Lines}}
{{Description}} — {{Amount}}
{{/each Lines}}
```

The dictionary overload is trimming and NativeAOT safe. A POCO overload is available for convenience and is annotated because it reflects over public properties. See the [template guide](https://officeimo.com/docs/word/templates/) for conditions, nested blocks, images, diagnostics, and the executable proof workflow.

### Mail merge fields

```csharp
var merge = document.AddParagraph();
merge.AddText("Customer: ");
merge.AddField(new WordFieldBuilder(WordFieldType.MergeField)
    .AddInstruction("CustomerName"));

var totalField = new WordFieldBuilder(WordFieldType.MergeField)
    .AddInstruction("OrderTotal")
    .SetFormat(WordFieldFormat.Numeric);
merge.AddField(totalField);
```

For a strict merge, use the structured result rather than assuming every field was bound:

```csharp
WordMailMergeExecutionReport report = WordMailMerge.ExecuteWithReport(
    document,
    new Dictionary<string, string> {
        ["CustomerName"] = "Ada Lovelace",
        ["OrderTotal"] = "1234.5"
    });

report.EnsureComplete();
```

The report distinguishes merged fields, missing values, and unsupported formatting. `ExecuteBatchWithReport(...)` retains the same evidence for every output record.

### OPC package signatures

```csharp
using OfficeIMO.Security;

IOfficeSecurityProvider security = OfficeSecurityProvider.Default;
WordDocument.SignPackage("report.docx", security, "CERTIFICATE-THUMBPRINT");

using WordDocument signed = WordDocument.Load("report.docx");
WordSignatureValidationReport validation = signed.ValidateSignatures(
    security,
    new WordSignatureValidationOptions());
```

OPC package signing does not sign VBA code. `WordSigningCapabilities.Package` and
`WordSigningCapabilities.MacroProject` report the two surfaces independently. `InspectSignatures()` remains available
without `OfficeIMO.Security`; it projects the shared bounded OPC inspector and does not claim digest, signature, or
certificate trust. The cross-host `InspectPackageSignatures(...)`, `ValidatePackageSignatures(...)`,
`SignPackageSignature(...)`, and `TrySignPackageSignature(...)` APIs expose the same result types used by Excel,
PowerPoint, and Visio. The established `ValidateSignatures(...)` API additionally retains Word-specific timestamp and
diagnostic evidence.

### VBA macro-project signatures

Signature parts in saved `.docm` and `.dotm` files can be inspected on every
supported platform without executing VBA:

```csharp
WordMacroProjectSignatureInfo signatures =
    WordDocument.InspectMacroProjectSignatures("automation.docm");
```

Managed VBA signing and content-binding validation work on every supported
platform. The workflow blocks existing OPC package signatures by default,
clears existing VBA signatures, creates and verifies the legacy, agile, and V3
profiles, proves that `vbaProject.bin` and the source package did not change
concurrently, and atomically replaces the package only after final validation.
When both signature kinds are needed, sign the VBA project first and the OPC
package last:

```csharp
using OfficeIMO.Security;

IOfficeSecurityProvider security = OfficeSecurityProvider.Default;
var options = new OfficeVbaSigningOptions();
options.CmsVerification.CertificateValidation.RevocationMode =
    X509RevocationMode.Online;

OfficeVbaSigningResult signing = WordDocument.SignVbaProject(
    "automation.docm",
    security,
    signingCertificate,
    options);

OfficeVbaSignatureValidationResult validation =
    WordDocument.ValidateVbaSignatures("automation.docm", security, options);
```

The caller supplies the certificate and `IOfficeSecurityProvider`; OfficeIMO
does not discover or persist private keys. Validation combines managed MS-OVBA
content binding with CMS signature, caller-controlled certificate-chain,
revocation, and RFC 3161 timestamp policy. Set
`ValidateWithWindowsSipWhenAvailable` only when a registered Microsoft Office
SIP should provide an additional differential check. Microsoft Office, SignTool,
and `offclearsig.exe` are not runtime dependencies. OfficeIMO does not execute
VBA or edit VBA source modules.

### Content controls

```csharp
document.FillContentControlValues(new Dictionary<string, object?> {
    ["Name"] = "Ada Lovelace",
    ["Approved"] = true,
    ["DueDate"] = DateTime.Today
});

Dictionary<string, object?> values = document.ExtractContentControlValues();
document.ValidateContentControlValues(values).EnsureValid();
```

### Legacy DOC files

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.LegacyDoc;

using WordDocument document = WordDocument.Load("legacy-input.doc");
document.Save("converted-output.docx");

LegacyDocWriteAssessment assessment = document.AssessLegacyDocWrite();
if (assessment.IsSupported) {
    document.Save("legacy-output.doc");
}

WordDocument.Convert("legacy-input.doc", "converted-output.docx");
WordDocument.Convert("openxml-input.docx", "legacy-output.doc");

using LegacyDocLoadResult result = WordDocument.LoadLegacyDocWithReport("legacy-input.doc");
if (result.HasDocument) {
    result.EnsureNoConversionLoss();
    result.Document.Save("converted-output.docx");
    string report = result.CreateAdvancedImportReport().ToMarkdown();
}
```

Legacy `.doc` support is first-party and dependency-free at runtime. The current
reader projects supported Word 97-2003 body paragraphs, simple zero-length,
same-paragraph, and cross-paragraph body bookmarks plus simple table-cell,
header/footer, and footnote/endnote paragraph bookmarks, simple external and
internal bookmark hyperlink fields with supported text, tab, soft/no-break
hyphen, and break display runs, simple static date/time and document-property
field display results,
common run and paragraph formatting including proofing exclusion, bidirectional
paragraph layout, mirror indents, contextual spacing, East Asian typography and
punctuation spacing flags, and automatic hyphenation suppression, built-in and
custom paragraph styles, simple tables, paragraph-boundary sections, page setup,
simple header/footer stories with tabs, text-wrapping and column breaks,
supported direct run formatting, and supported paragraph formatting, simple
footnote/endnote bodies with supported direct run and paragraph formatting and
soft/no-break hyphen runs, section note numbering and placement settings, and
document properties into the normal `WordDocument` model. Native `.doc` saving
is available for the supported simple subset: paragraphs, simple zero-length,
same-paragraph, and cross-paragraph body bookmarks plus simple table-cell,
header/footer, and footnote/endnote paragraph bookmarks, simple external and
internal bookmark hyperlinks with supported text, tab, soft/no-break hyphen,
break display runs, simple static date/time and document-property fields with
static display text and supported inline result characters including inside flattened inline content
controls, simple inline content-control display text, and simple block content
controls with nested
simple block controls plus nested inline content controls in
body/table/header/footer/footnote/endnote stories, common run and paragraph
formatting including proofing exclusion,
bidirectional paragraph layout, mirror indents, contextual spacing, East Asian
typography and punctuation spacing flags, and automatic hyphenation suppression,
tabs, soft/no-break hyphen runs, line/carriage-return/page/column breaks, simple
body tables with common formatting, including simple depth-2 nested tables,
supported table-style border, shading, layout, paragraph
formatting, run formatting, default-cell expansion, conditional table/cell
border, shading, paragraph formatting, run formatting, cell-layout expansion,
and conditional row height/header/no-split formatting, paragraph-boundary
sections, page setup, simple header/footer stories with tabs, soft/no-break
hyphen runs, text-wrapping, carriage-return, and column breaks, supported
direct run formatting, and supported paragraph formatting,
simple footnote/endnote bodies with supported direct run and paragraph
formatting and soft/no-break hyphen runs, supported section note settings, and
scalar document properties. Unsupported features such as macros, embedded OLE
objects, comments, text boxes, images, bookmark ranges outside supported
body/table-cell/header/footer/footnote/endnote paragraphs, richer
content-control children, richer visual table style effects, deeper or richer
nested table shapes,
richer note body structures, and richer header/footer or section shapes are
diagnosed or blocked rather than silently flattened. `WordDocument.Convert(...)`
uses those same load and save paths and blocks legacy sources with unsupported
or preserve-only content by default. Set `LossPolicy` to
`OfficeConversionLossPolicy.Allow` on `WordDocumentConversionOptions` or
`WordSaveOptions` only when that loss has been reviewed and is intentional.
See [DOC and DOCX compatibility](../Docs/officeimo.word.legacy-doc-compatibility.md)
for the current capability matrix and safety contract. Use the
[migration guide](../MIGRATION.md#legacy-doc-and-xls-api-changes) for canonical API replacements.

### Protection

```csharp
using DocumentFormat.OpenXml.Wordprocessing;

document.Settings.ProtectionPassword = "owner-password";
document.Settings.ProtectionType = WordDocumentProtectionType.ReadOnly;
```

### Editable equations from the shared math model

```csharp
using OfficeIMO.Drawing;

OfficeMathExpression equation = OfficeMath.Fraction(
    OfficeMath.Superscript(OfficeMath.Identifier("x"), OfficeMath.Number("2")),
    OfficeMath.Number("2"));

WordParagraph paragraph = document.AddEquation(equation);
paragraph.AddText(" is editable Word math.");
```

`WordDocument.AddEquation(...)` and `WordParagraph.AddEquation(...)` map the shared expression directly to native OMML. Existing equations expose `ToExpression()`, `SetExpression(...)`, and `ToDrawing(...)`; `WordMathMarkup` converts between OMML and `OfficeMathExpression`. The adapter covers matrices and multi-column equation arrays, left/right scripts, centered limits, skewed fractions, delimiter lists, n-ary operators, and decorations. Display-equation replacement retains `oMathParaPr` presentation metadata. Shared `Stack` and `StretchStack` nodes fail closed because OMML has no lossless equivalent; use `OfficeMath.EquationArray(...)` explicitly if that alternate layout is acceptable. The reusable AST stays in `OfficeIMO.Drawing`, while Word owns only the OMML adapter.

### Convert with adjacent packages

```csharp
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;
using OfficeIMO.Word.Pdf;

string html = document.ToHtml(new WordToHtmlOptions { IncludeDefaultCss = true });
string markdown = document.ToMarkdown(new WordToMarkdownOptions());
document.SaveAsPdf("report.pdf");
```

## Managed image export

Word page previews use the shared Drawing renderer and can be returned as PNG, JPEG, TIFF, lossless WebP, or SVG without Office automation:

```csharp
using OfficeIMO.Drawing;

byte[] webp = document.ToWebp(new WordImageExportOptions { PageIndex = 0, Scale = 1.5 });

document.ToImage()
    .Page(0)
    .FitWithin(1600, 1200)
    .AsJpeg()
    .WithRasterEncoding(raster => raster.Jpeg.Quality = 90)
    .Save("page-1.jpg");
```

The document package owns Word pagination and diagnostics; `OfficeIMO.Drawing` owns sizing, pixels, and encoding. The same fit limit applies to SVG and raster output. `SaveAsJpeg`, `SaveAsTiff`, and `SaveAsWebp` are thin convenience wrappers over the same builder.

## Content provenance

Inspect C2PA and AI-specific IPTC metadata in the package and its supported embedded images, then remove only the selected carriers:

```csharp
using OfficeIMO.Provenance;
using OfficeIMO.Word;

OfficeProvenanceReport report = WordDocument.InspectProvenance("input.docx");
OfficeProvenanceRemovalResult result = WordDocument.RemoveProvenance("input.docx", "clean.docx");
```

Mutation of a signed package is blocked by default. Set `SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures` only when removing the now-invalid package signature is intentional. Optional cryptographic C2PA verification is provided by `OfficeIMO.Security`.

## Concealed-content inspection and cleanup

`WordDocument.InspectContentSafety(...)` reports native/inherited hidden runs, deleted revisions, tiny or zero-geometry text, explicit low contrast, comments/notes, alternative text, and exact Unicode evidence. Pass reviewed finding IDs to `WordDocument.RemoveSelectedContent(...)`; the package is reopened after cleanup, and signed-document mutation fails closed by default. These findings describe ingestion risk, not AI authorship.

## Adjacent packages

`OfficeIMO.Word` owns the Word model. Conversion and export packages stay separate so consumers only take the dependencies they need:

| Package | Use it for |
| --- | --- |
| [OfficeIMO.Word.Html](../OfficeIMO.Word.Html/README.md) | Word to/from HTML conversion. |
| [OfficeIMO.Word.Markdown](../OfficeIMO.Word.Markdown/README.md) | Word to/from Markdown conversion. |
| [OfficeIMO.Word.Pdf](../OfficeIMO.Word.Pdf/README.md) | Word to PDF export through `OfficeIMO.Pdf`. |
| [OfficeIMO.Word.GoogleDocs](../OfficeIMO.Word.GoogleDocs/README.md) | Planning and exporting Word content to Google Docs. |

## Related packages

- Use [PSWriteOffice](https://github.com/EvotecIT/PSWriteOffice) for PowerShell examples and cmdlets.
- Use `OfficeIMO.Word.Pdf` for Word-to-PDF conversion and `OfficeIMO.Pdf` for direct PDF layout and manipulation.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`; `net472` is included when building on Windows.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

Runnable samples live under [OfficeIMO.Examples/Word](../OfficeIMO.Examples/Word).

## Dependency footprint

- **External:** Open XML SDK for `.docx` package mechanics. Microsoft BCL compatibility packages are used on older targets.
- **OfficeIMO:** `OfficeIMO.Core`. The fluent model, native OMML adapter, legacy `.doc` reader/writer, lifecycle, validation, and PNG/JPEG/TIFF/WebP/SVG export are first-party.
- **Optional security:** install `OfficeIMO.Security` and pass `OfficeSecurityProvider.Default` only for OPC/VBA signing or cryptographic validation. It is not a transitive Word dependency.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
