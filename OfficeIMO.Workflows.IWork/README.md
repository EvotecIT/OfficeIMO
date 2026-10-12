# OfficeIMO.Workflows.IWork

Opt-in Pages-to-Word, Numbers-to-Excel, Keynote-to-PowerPoint, and complete editable iWork-to-PDF conversion through the shared OfficeIMO workflow runner.

```csharp
using OfficeIMO.Workflows;
using OfficeIMO.Workflows.IWork;

OfficeWorkflowRunner runner = IWorkWorkflow.CreateRunner();
OfficeWorkflowResult result = await runner.RunAsync(new OfficeWorkflowRequest {
    InputPath = "budget.numbers",
    OutputPath = "budget.xlsx",
    Operation = OfficeWorkflowOperation.Convert,
    ConversionRouteId = "numbers-xlsx"
});
```

The runner captures the source, bounds input and output, reopens the staged destination through its format owner, verifies the source before publication, and applies the requested conflict policy. The `SourceSnapshot` diagnostic identifies captured bytes by SHA-256, length, selected filename, and `snapshotKind`. For a directory package, the checksum describes a deterministic private ZIP transport snapshot of the captured file membership and contents; it is not a checksum of an original ZIP file. `ConversionEvidence.Facts["packageContentSha256"]` separately identifies the normalized entries read by the iWork owner, including expanded nested indexes. It matches `IWorkSourceDocument.ComputePackageContentHash()` for equivalent captured package content. Package-root replacement and changes to membership or content prevent publication. `ConversionEvidence` retains typed fidelity categories and compact producer, projection, and coverage facts after completed conversion, including when output validation or publication later fails or is cancelled. Successful publication does not establish complete source fidelity or visual equivalence. `OutputBytes` and `OutputSha256` describe the captured published artifact; result construction does not reopen the final path.

The default acceptance policy rejects explicitly partial editable reconstruction and previews without known full-document coverage. Pass `IWorkConversionOptions` to `CreateRunner` to choose editable or visual output, permit partial reconstruction, or normalize Numbers worksheet names. These options are copied when routes are registered. Each request can instead supply its own `IWorkWorkflowSettings`; the runner snapshots them before opening the source, so later caller edits cannot change a pending conversion. Destination renames, approximations, omissions, and unassessed source records remain reported.

```csharp
var settings = new IWorkWorkflowSettings();
settings.ConversionOptions.Mode = OfficeIMO.IWork.IWorkConversionMode.VisualOnly;
// Explicitly accept a preview that may cover only part of the source.
settings.ConversionOptions.RequireCompleteVisualCoverage = false;
var request = new OfficeWorkflowRequest {
    InputPath = "report.pages",
    OutputPath = "report.docx",
    Operation = OfficeWorkflowOperation.Convert,
    ConversionRouteId = "pages-docx",
    RegisteredConversionSettings = settings
};
OfficeWorkflowResult preview = await runner.RunAsync(request);
```

The `pages-pdf`, `numbers-pdf`, and `keynote-pdf` routes compose the same editable adapters with the existing Word, Excel, and PowerPoint PDF owners. They reject visual previews, explicitly partial reconstruction, content-excluding options, and reported PDF-stage omissions, clipping or failures before publication. Numbers includes every worksheet, row, and column without print-area clipping; Keynote includes hidden slides. Complete editable content does not establish exact pagination or typography. Font and rendering approximations remain in the result.

```csharp
var pdfSettings = new IWorkWorkflowSettings();
OfficeWorkflowResult pdf = await runner.RunAsync(new OfficeWorkflowRequest {
    InputPath = "presentation.key",
    OutputPath = "presentation.pdf",
    Operation = OfficeWorkflowOperation.Convert,
    ConversionRouteId = "keynote-pdf",
    RegisteredConversionSettings = pdfSettings
});
```

Use `PagesPdfOptions`, `NumbersPdfOptions`, and `KeynotePdfOptions` for destination layout, font, and resource policy; request snapshots copy these options. Registered PDF options reject encryption before opening the source; apply the PDF protection workflow to a completed conversion. `ConversionEvidence` retains both the iWork and PDF reports, including serialization/font preflight evidence, with `conversionStages = iwork-editable,pdf`. `OfficeWorkflowRequest.ConversionOptions.CompressPdfOutput` applies the shared verified lossless PDF optimizer. Failed PDF saves and postprocessing retain available stage evidence and preserve an existing destination. Cancellation after staged-output reopening also removes staging without replacing the destination.

Supply a shared `OfficeRenderingProfile` when the source needs glyphs outside the PDF base fonts. This profile uses caller-owned font files and disables installed-font discovery:

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

var fonts = new OfficeFontFaceCollection();
foreach (var (name, style) in new[] {
    ("Regular", OfficeFontStyle.Regular), ("Bold", OfficeFontStyle.Bold),
    ("Italic", OfficeFontStyle.Italic), ("BoldItalic", OfficeFontStyle.Bold | OfficeFontStyle.Italic)
}) fonts.Add("Carlito", File.ReadAllBytes($"fonts/Carlito-{name}.ttf"), style);
fonts.Add("Noto Emoji", File.ReadAllBytes("fonts/NotoEmoji-VariableFont_wght.ttf"))
    .AddFallbackFamily("Noto Emoji");
fonts.Add("Noto JP Common", File.ReadAllBytes("fonts/NotoSansJP-OfficeIMO-Common.ttf"))
    .AddFallbackFamily("Noto JP Common");
fonts.AddFallbackFamily("Carlito");
var profile = new OfficeRenderingProfile("portable-iwork", fonts);
var portablePdfSettings = new IWorkWorkflowSettings();
portablePdfSettings.PagesPdfOptions.UseRenderingProfile(profile);
portablePdfSettings.NumbersPdfOptions.UseRenderingProfile(profile);
portablePdfSettings.KeynotePdfOptions.UseRenderingProfile(profile);
portablePdfSettings.PagesPdfOptions.ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic();
portablePdfSettings.NumbersPdfOptions.ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic();
portablePdfSettings.KeynotePdfOptions.ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic();
```

Pass `portablePdfSettings` as `RegisteredConversionSettings`. Choose licensed faces covering the characters and styles in the source; the profile does not bundle fonts in OfficeIMO. The [native-content PDF reference](../OfficeIMO.TestAssets/Documents/IWorkCorpus/native-exports/iwork-composed-pdf-portable.json) pins these font inputs and licenses, complete text/value checks and every rendered page: Pages produces two pages, Numbers four and Keynote three. Font substitutions, monochrome emoji, text metrics and numeric display remain approximations. Numbers follows the XLSX adapter's table-per-worksheet layout and can split column ranges across PDF pages; this differs from Apple's positioned sheet export.

Routes accept local ZIP files, local directory packages, provider ZIP streams, and provider directory packages. Directory capture uses the bounded iWork container owner, including nested `Index.zip` packages. Output cannot be placed inside a source package or alias one of its members. For provider directories, set `InputDirectoryPackage` to an `OfficeWorkflowDirectoryPackageInput` containing the selected filename, a reopenable `OfficeWorkflowDirectoryInput`, and a permission-aware root identity/output-separation guard. Set an explicit `OutputPath` and omit `InputStream`. The shared runner captures at most `MaximumEntries` members (10,000 by default) within `Limits.MaximumInputBytes`, then invokes the format owner on private staging. Before publication it re-enumerates the original provider and verifies member identities and bytes. The provider guard must reject root replacement and destinations inside or containing the source package. Native sandbox/bookmark qualification remains separate from this provider contract. See the [iWork support matrix](../Docs/officeimo.iwork-support-matrix.md) for format limits and the [workflow guide](../OfficeIMO.Workflows/README.md) for publication and provider contracts.

Retained evidence includes identified source-unit totals and reconstructed, omitted, and unassessed counts. These totals describe units selected by the semantic projection, including explicitly omitted unsupported drawables and Numbers sheet references. Inactive templates, auxiliary table models/tiles, unresolved references, and unassessed content paths remain excluded. A visual preview does not establish individual object coverage. Full per-kind counts and native object identities are available on the owning adapter's `IWorkConversionReport`.

Conversion evidence includes source formula-expression and cache counts from `IWorkConversionReport.FormulaSummary`. These include separate unassessed expression and cache counts for supported headers declaring formulas in undecoded cells, independently of editable or preview output. Per-cell identities and assessments remain in the owning adapter report.

Declaration evidence retains `sourceDeclarationIssueCount` after destination disposal. It counts distinct selected paths with unreadable or rejected declarations, separately from reference occurrences and objects. Owner/path details and known outer field-value counts are available in `IWorkConversionReport.SourceDeclarationIssues`; unreadable nested content and preview coverage remain unassessed.

Cell evidence retains `sourceCellIssueCount` after destination disposal. It counts materialized selected cells with decoding failures; native error markers and unmaterialized storage are excluded. Table identities, coordinates and messages are available in `IWorkConversionReport.SourceCellIssues`.

Reference evidence retains `sourceReferenceIssueCount`, `sourceMissingReferenceTargetCount`, `sourceMalformedReferenceCount`, `sourceRejectedReferenceSetCount` and `sourceUnexpectedReferenceTargetTypeCount` after destination disposal. These count failed declared occurrences in the core's assessed content paths, separately from identified objects. Repeated occurrences within a source field remain separate; reuse of a shared text or template archive does not multiply its evidence. Full owner identities, field paths and target identifiers remain in `IWorkConversionReport.SourceReferenceIssues`. Missing, unreadable or unsupported target-type references have unassessed fidelity; the counts do not establish omitted-object totals or preview coverage.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Convert | 0 | 3 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Workflows.IWork` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
