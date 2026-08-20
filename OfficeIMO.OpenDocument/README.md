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

Create a sparse ODS workbook:

```csharp
using OdsDocument workbook = OdsDocument.Create();
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

Create an ODP presentation:

```csharp
using OdpPresentation presentation = OdpPresentation.Create();
OdpSlide slide = presentation.AddSlide("Summary");
slide.AddTextBox(OdfRect.FromCentimeters(2, 1, 28, 3), "Native ODP");
slide.AddRectangle(OdfRect.FromCentimeters(2, 5, 8, 3)).FillColor = OdfColor.Parse("#D1E9FF");
slide.GetOrCreateSpeakerNotes().AddParagraph("Explain the result.");
presentation.Save("summary.odp");
```

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
| ODT | Paragraphs, headings, ordered inline text/span/link/image/bookmark syntax, whitespace controls, common text and paragraph styles, lists, tables, sections, page layout, headers/footers, page breaks, images, paragraph insertion/deletion tracking |
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
- Ordered ODT/ODP inline syntax types direct text, spans/runs, hyperlinks, images, and bookmark markers. Nested inline markup remains preserved in the package and is surfaced as an untyped node so conversion loss is explicit.
- Tracked-change editing covers paragraph insertions and deletions. Arbitrary inline merges and conflict resolution remain preservation-oriented.
- Animation editing covers basic shape-attribute effects and fade-in timing. Advanced timing trees are preserved when untouched.
- Password-encrypted packages using the documented AES-256-CBC profile can be opened and written. Legacy Blowfish and other unsupported profiles fail before content is exposed.
- Changed signed packages fail by default because saving would invalidate signatures. An explicit save option can remove invalidated signature entries.
- The bounded OfficeIMO XML package-manifest signature profile can be created and validated through an explicit `IOfficeSecurityProvider`. Arbitrary producer-specific signature profiles remain inspection or preservation oriented.
- Pivot-table editing and complete chart editing are outside the current surface.
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
