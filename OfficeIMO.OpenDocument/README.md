# OfficeIMO.OpenDocument

`OfficeIMO.OpenDocument` creates and edits ODT, ODS, and ODP files directly. Its only runtime dependency is the zero-dependency `OfficeIMO.Core` foundation used across OfficeIMO for lifecycle and result contracts. It has no third-party runtime dependencies and does not invoke LibreOffice, Microsoft Office, or UNO.

```powershell
dotnet add package OfficeIMO.OpenDocument
```

## Create documents

Create an ODT document:

```csharp
using OfficeIMO.OpenDocument;

using OdtDocument document = OdtDocument.Create();
document.AddHeading("Summary", 1);
document.AddParagraph("Created with OfficeIMO.OpenDocument.");

OdtTable table = document.AddTable(2, 2, "Results");
table.Cell(0, 0).Text = "Metric";
table.Cell(0, 1).Text = "Value";
table.Cell(1, 0).Text = "Revenue";
table.Cell(1, 1).Text = "42";

document.Save("summary.odt");
```

Add a native ODT field in paragraph order with the value currently shown in the document:

```csharp
OdtParagraph pageLine = document.AddParagraph("Page ");
pageLine.AddField(OdtFieldKind.PageNumber, "3");
pageLine.AddText(" of ");
pageLine.AddField(OdtFieldKind.PageCount, "12");
document.Save("summary.odt");
```

`OdtParagraph.Fields` reads page number, page count, date, and time fields in order. `OdtField.DisplayText` edits the cached text; an office application can refresh a dynamic field later. `OdtField.IsFixed` retains a date, time, or page number value instead of refreshing it. Page count cannot be fixed.

Add native ODT footnotes or endnotes at the current paragraph position:

```csharp
using OfficeIMO.OpenDocument;

using OdtDocument document = OdtDocument.Create();
OdtParagraph paragraph = document.AddParagraph("The result is documented");
paragraph.AddFootnote("Source and calculation details.");
paragraph.AddText(" in the appendix.");
paragraph.AddEndnote("Additional context.");
```

`OdtParagraph.Notes` and `InlineNodes` expose note bodies and reference order after reopening the file. Note-body paragraphs are separate from document-body paragraphs, and their spans, links, and images stay on the note-body paragraph. Native note citations follow document order when notes are added to earlier paragraphs later; an imported custom citation label remains the displayed `OdtNote.Citation`. Adding a note to a document with `text:notes-configuration` for that note kind throws `NotSupportedException`, preserving its configured numbering and existing citations.

Create a sparse ODS workbook:

```csharp
OdsDocument workbook = OdsDocument.Create();
OdsSheet sheet = workbook.AddSheet("Metrics");
sheet.Cell(0, 0).SetString("Name");
sheet.Cell(0, 1).SetString("Value");
sheet.Cell(1, 0).SetString("Revenue");
sheet.Cell(1, 1).SetDecimal(42.5m);

OdsCell total = sheet.Cell(2, 1);
total.Formula = "of:=SUM([.B2:.B2])";
OdsValidation positive = workbook.AddValidation(
    "PositiveAmount",
    OdsValidationConditionSyntax.Create(
        OdsValidationValueKind.DecimalNumber,
        OdsValidationComparison.GreaterThan,
        "0"));
positive.SetHelpMessage("Amount", "Enter a value greater than zero.");
positive.SetErrorMessage("Invalid amount", "The amount must be positive.");
sheet.Cell(1, 1).ValidationName = positive.Name;
OdsRecalculationReport calculation = workbook.Recalculate();
if (calculation.FailedCells > 0) {
    Console.WriteLine(calculation.Diagnostics[0].Message);
}

workbook.Save("metrics.ods");
```

For a formula-based validation, use `OdsValidationConditionSyntax.CreateFormula("[.B2]>0")` and set `OdsValidation.BaseCellAddress` to an absolute sheet-qualified address such as `$'Metrics'.$B$2`. The typed formula condition accepts a local cell compared with a number, text value, or another local cell, plus `ISBLANK`, `ISNUMBER`, or `ISTEXT` on one local cell. These predicates can be combined with `AND`, `OR`, and unary `NOT`; numbers in comparisons may have a leading sign. Other OpenFormula expressions can be stored in the raw `Condition` property. The base cell determines how relative formula references apply to the validated cells.

ODS conditional cell styles can be authored and edited through `OdfStyle.AddConditionalMap`. Create a common named table-cell style for the desired appearance, add a mapping to a base style, and assign the base style to cells. For example, `baseStyle.AddConditionalMap("cell-content()>0", highlightStyle.Name, "$'Metrics'.$B$2")` applies the highlight style when the condition is true. The applied style must be a common named style in the same family as the base style; `Validate()` reports a missing, automatic, or different-family target. The map and its relative-reference base survive save and reopen. OfficeIMO preserves these native rules; its renderer does not evaluate them.

Create an ODP presentation:

```csharp
using OdpPresentation presentation = OdpPresentation.Create();
OdpSlide slide = presentation.AddSlide("Summary");
slide.AddTextBox(OdfRect.FromCentimeters(2, 1, 28, 3), "Native ODP");
slide.AddRectangle(OdfRect.FromCentimeters(2, 5, 8, 3)).FillColor = OdfColor.Parse("#D1E9FF");
slide.GetOrCreateSpeakerNotes().AddParagraph("Explain the result.");
presentation.Save("summary.odp");
```

Nest text styles and links when their formatting changes within a sentence:

```csharp
using OfficeIMO.OpenDocument;

using OdtDocument document = OdtDocument.Create();
OdtParagraph paragraph = document.AddParagraph();
OdtSpan emphasis = paragraph.AddSpan("Read ");
emphasis.Bold = true;
emphasis.AddHyperlink("the guide", "https://example.com/guide").Italic = true;

using OdpPresentation presentation = OdpPresentation.Create();
OdpSlide slide = presentation.AddSlide("Links");
OdpParagraph slideText = slide.AddTextBox(
    OdfRect.FromCentimeters(2, 9, 18, 2)).AddParagraph();
OdpRun label = slideText.AddRun("Open ");
label.Bold = true;
label.AddHyperlink("the guide", "https://example.com/guide").Underline = true;
```

`InlineNodes` exposes nested `Children` in document order. A nested run inherits text properties from its containing span or link until its own style overrides them.

Convert explicitly between OpenDocument and OfficeIMO Word, Excel, or PowerPoint models by installing the corresponding adapter package. Every conversion returns an `OdfConversionReport` that identifies mapped, approximated, skipped, and unsupported features.

```powershell
dotnet add package OfficeIMO.Word.OpenDocument
dotnet add package OfficeIMO.Excel.OpenDocument
dotnet add package OfficeIMO.PowerPoint.OpenDocument
```

## Edit without flattening the package

Typed objects remain backed by the source XML. A targeted edit rewrites its owning XML part while untouched package entries keep their original bytes.

```csharp
using OdtDocument document = OdtDocument.Open("input.odt");
document.Paragraphs[0].Text = "Updated text";
OdfSaveResult result = document.Save("output.odt", new OdfSaveOptions {
    CompatibilityProfile = OdfCompatibilityProfile.PreserveSource
});

IReadOnlyList<string> rewritten = result.Report.RewrittenEntries;
IReadOnlyList<string> lossy = result.Report.LossyEntries;
```

New documents use ODF 1.4. Set `OdfCompatibilityProfile.Odf13` when the output needs the ODF 1.3 schema and compatibility profile.

## Encrypt and decrypt ODF packages

Password encryption is format-owned and does not require `OfficeIMO.Security`:

```csharp
using OdtDocument document = OdtDocument.Load("protected.odt", new OdfLoadOptions {
    Password = password
});

document.AddParagraph("Updated while decrypted in memory.");
document.Save("protected-updated.odt", new OdfSaveOptions {
    Encryption = new OdfEncryptionOptions {
        Password = newPassword
    }
});
```

The password is UTF-8, used only for the current load or save, and is not retained. Input accepts 10,000 through 10,000,000 PBKDF2 iterations per entry and preflights the complete manifest against `OdfLoadOptions.MaxTotalKdfIterations` (10,000,000 by default) before deriving any entry key. Output uses AES-256-CBC, a SHA-256 password start key, per-entry PBKDF2-HMAC-SHA1 with 100,000 iterations by default, and SHA-256/1K checksums. Each encrypted entry receives fresh salt and initialization-vector material.

Encrypted input fails with a classified `OdfEncryptedPackageException` when a password is missing or incorrect, the profile is unsupported, metadata is malformed, or decrypted content exceeds configured limits. Saving an encrypted source without `OdfSaveOptions.Encryption` also fails so protection is not removed accidentally. To write plaintext intentionally, set `EncryptionHandling = OdfEncryptionHandling.Remove`.

## Supported editing surface

| Area | Current support |
| --- | --- |
| Package | Bounded ZIP/XML loading, direct reading of seekable package streams, manifest updates, deterministic output, metadata, atomic path saves, flat XML projection with loss reporting, unknown-entry preservation |
| ODT | Paragraphs, headings, ordered inline text/span/link/image/bookmark syntax, page number/count and date/time fields with cached display text, whitespace controls, common text and paragraph styles, lists, tables, sections, page layout, default/first/left master-page headers and footers, page breaks, images, paragraph insertion/deletion tracking |
| ODS | Sparse repeated rows/cells, typed values, OpenFormula text and cached values, bounded formula evaluation/recalculation, styles and data formats, merges, row/column sizing and visibility, sheet order, typed named ranges, annotations, typed scalar/list validations and messages, links, print ranges |
| ODP | Slide order and visibility, page size, masters/layouts, ordered inline text/run/link syntax, common run styles, lists, rectangles, ellipses, lines, groups, transforms, images and crop, tables, speaker notes, backgrounds, transitions, basic shape animations |
| Inspection | Annotations, tracked changes, extension namespaces, scripts, event listeners, external links, embedded objects, formulas, validations, transitions, animations, encryption, and signatures |

Unknown XML, vendor extensions, scripts, embedded content, and unsupported drawing features are preserved when their owning part is not replaced. The library never executes scripts, macros, event listeners, embedded objects, or external links. Formula evaluation is a bounded, side-effect-free parser for the documented local subset; it does not execute active content or fetch data.

`OdfCapabilityCatalog.Advanced` provides stable capability IDs and distinguishes editable subsets, preserved content, inspection, and detected-but-unsupported features.

## Content provenance

`OdfDocument.InspectProvenance("input.odt")` reports C2PA and AI-specific IPTC metadata in ODF packages and supported embedded images. `OdfDocument.RemoveProvenance("input.odt", "clean.odt")` performs a bounded package rewrite while preserving the required uncompressed, first `mimetype` entry. Signed-package mutation is blocked unless removal of invalidated ODF signature entries is requested explicitly. Optional cryptographic C2PA verification remains in `OfficeIMO.Security`.

## Concealed-content inspection and cleanup

`OdfDocument.InspectContentSafety(...)` covers ODT, ODS, and ODP native hidden fields and containers, concealed stored values/formulas, resolved tiny/transparent/low-contrast styles, zero geometry, notes, annotations, alternative descriptions, and Unicode evidence. `OdfDocument.RemoveSelectedContent(...)` removes exact reviewed text segments or exact Unicode ranges inside stored attributes through the preservation-aware package writer. Encrypted-source cleanup and implicit signature invalidation are rejected.

## Explicit boundaries

- Formula evaluation covers arithmetic, comparisons, concatenation, cell/range references, and common aggregate/math functions. External data, volatile functions, matrix formulas, and the complete OpenFormula language are not included.
- Typed validation syntax covers explicit lists and scalar whole-number, decimal, and text-length comparisons. Other valid ODF conditions remain preserved text and are reported by conversions that cannot map them exactly.
- Ordered ODT/ODP inline syntax types text, nested spans/runs, and hyperlinks. ODT also types inline images and bookmark markers. Unsupported inline elements remain `Other` nodes and conversion reports their approximation.
- Tracked-change editing covers paragraph insertions and deletions. Arbitrary inline merges and conflict resolution remain preservation-oriented.
- Animation editing covers basic shape-attribute effects and fade-in timing. Advanced timing trees are preserved when untouched.
- Password-encrypted packages using the documented AES-256-CBC profile can be opened and written. Legacy Blowfish and other unsupported profiles fail before content is exposed.
- Changed signed packages fail by default because saving would invalidate signatures. An explicit save option can remove invalidated signature entries.
- The bounded OfficeIMO XML package-manifest signature profile can be created and validated through an explicit `IOfficeSecurityProvider`. Arbitrary producer-specific signature profiles remain inspection or preservation oriented.
- ODS exposes embedded chart names, types, titles, source ranges, and frame positions through `OdsSheet.Charts`. `OdsSheet.AddChart` creates column, bar, or line charts with one to sixteen series linked to existing one-dimensional ODS cell ranges of up to 4,096 points. Chart styling is preserved in package XML; editing imported charts and pivot tables is outside the current surface.
- Flat XML variants (`.fodt`, `.fods`, `.fodp`) can be opened and written, including embedded raster images. Exotic embedded objects and package-only features may not project losslessly.
- `OdsSheet.Merge` rejects merges above its default 100,000-cell materialization limit. Use the overload with an explicit lower limit when processing untrusted dimensions.
- Unknown package entries and extension XML are always preserved by package editing. Explicit format conversion and flat XML projection report content they cannot carry through `OdfConversionReport` and `OdfSaveReport.LossyEntries`.

The package targets `netstandard2.0`, `net8.0`, and `net10.0`, plus `net472` on Windows. CI checks generated ODF 1.3 and 1.4 XML against pinned OASIS Relax NG schemas, then opens and resaves the generated packages with the runner's reported LibreOffice version.

Interoperability coverage includes ODT, ODS, and ODP files from LibreOffice and Microsoft Office, plus an externally verified Google Docs ODT export. These files exercise styles, formulas, drawings, embedded content, and preservation of unknown package entries. A separate hash-pinned LibreOffice fixture covers password encryption, including OfficeIMO reading LibreOffice output and LibreOffice reading OfficeIMO output. See the [producer manifest](../OfficeIMO.OpenDocument.Tests/Fixtures/producer-manifest.json) and [encryption manifest](../OfficeIMO.OpenDocument.Tests/Fixtures/Encryption/producer-manifest.json) for exact producer versions, hashes, and evidence.

## Dependency footprint

- **External:** None; no OpenDocument SDK and no LibreOffice process.
- **OfficeIMO:** `OfficeIMO.Core`. ODT/ODS/ODP parsing, models, preservation, inspection, and writing are first-party.
- **Security:** ODF password encryption/decryption is first-party and dependency-free. Signature carriers are detected and changed signed packages fail safely without a cryptographic dependency. `OdfDocument.SignPackage(...)` and `ValidatePackageSignatures(...)` use an explicit provider for the bounded OfficeIMO XML package-manifest profile; `OfficeIMO.Security` is not pulled transitively.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 3 | 0 | 0 | 0 | 0 | 0 |
| Read | 2 | 0 | 0 | 0 | 0 | 1 |
| Edit | 2 | 0 | 0 | 1 | 0 | 0 |
| Preserve | 1 | 0 | 0 | 0 | 0 | 0 |
| Inspect | 4 | 0 | 0 | 0 | 0 | 0 |
| Validate | 3 | 0 | 0 | 0 | 0 | 1 |
| Remove | 3 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.OpenDocument` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
