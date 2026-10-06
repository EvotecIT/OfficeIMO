# OfficeIMO.Workflows

`OfficeIMO.Workflows` is the reusable local orchestration layer for OfficeIMO document jobs. It composes the existing first-party conversion, PDF, and provenance APIs behind typed requests, bounded execution, cooperative cancellation, collision policies, atomic output publication, and post-write validation.

The package does not add a second document or PDF engine. Desktop applications, command-line tools, and services can share this workflow contract while keeping their user-interface and hosting code thin.

## Opt-in Apple conversion

[OfficeIMO.Workflows.IWork](../OfficeIMO.Workflows.IWork/README.md) supplies Pages-to-Word, Numbers-to-Excel, and Keynote-to-PowerPoint routes. `IWorkWorkflow.CreateRunner()` shares this runner's source capture, limits, destination reopen validation, and publication contract. `IOfficeWorkflowRunner.ConversionRoutes` exposes the configured executable routes; the static `OfficeWorkflowCatalog` describes built-in executability. The default workflow package remains independent of iWork.

`OfficeWorkflowConversionRegistration` adds an implementation of an existing canonical route to a runner. It cannot replace built-in owners. Opt-in converters accept captured ZIP/file streams, write to a bounded caller-owned output stream, and return immutable `OfficeWorkflowConversionEvidence`. Current opt-in destination formats are DOCX, XLSX, and PPTX. `OfficeWorkflowResult.ConversionEvidence` retains fidelity categories and compact source facts; successful reopen does not establish visual equivalence.

Provider directory packages use `OfficeWorkflowRequest.InputDirectoryPackage` with a registered directory-package converter. The shared runner preserves member layout in bounded private staging and verifies original provider membership and content before publication. The host supplies permission-aware root identity and output-separation checks through `OfficeWorkflowDirectoryPackageInput.SourcePublicationGuard`. The selected filename determines routing; an explicit output is required.

## Single conversions and file batches

`OfficeWorkflowRunner` executes its configured `ConversionRoutes`. Ordinary directory and selected-file batches use those same routes, including registered adapters, profiles, renderer options, diagnostics and publication policies. Registered adapters use the options captured when the runner was created. PDF export covers DOC, DOCX, TXT, XLSX, PPTX, HTML, Markdown and RTF. Other built-in targets include PDF-to-DOCX/XLSX/PPTX/HTML and reviewed book-project-to-EPUB export. Unsupported or filtered files produce skipped outcomes; their count is separate from selected conversions.

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

var options = new OfficeWorkflowConversionOptions {
    PlainText = new PdfPlainTextOptions { TabSize = 4 }
};
var runner = new OfficeWorkflowRunner();
var single = await runner.RunAsync(new OfficeWorkflowRequest {
    Operation = OfficeWorkflowOperation.Convert,
    InputPath = "report.txt", OutputPath = "report.pdf", ConversionRouteId = "txt-pdf", ConversionOptions = options
}, cancellationToken: cancellationToken);

OfficeConversionBatchResult batch = await OfficeWorkflow.ConvertDirectory("Documents")
    .ToDirectory("PDF")
    .WithConversionOptions(options)
    .WithConcurrency(2)
    .RunAsync(runner, cancellationToken: cancellationToken);

// Selected files use the same batch API; the target need not be PDF.
var html = await OfficeWorkflow.ConvertFiles("report.pdf", "appendix.pdf")
    .ToDirectory("HTML", ".html")
    .RunAsync(runner, cancellationToken: cancellationToken);
```

`Word`, `Excel`, `PowerPoint`, `Html`, `Markdown`, `Rtf` and `PlainText` accept their owning adapter's typed options. A mixed batch selects the settings applicable to each route. Explicit renderer options take precedence over the cross-format `OutputProfile`. For ambiguous source extensions, use `.Via("html-pdf")` or `ConversionRouteId` on the request; TXT otherwise remains literal text. Encrypted Office inputs use `ConversionOptions.SourcePassword`; PDF inputs use `PdfPassword`. Passwords remain runtime inputs and are not stored in checkpoints.

Ordinary batches support the existing `Fail`, `Rename` and `Replace` conflict policies. Directory discovery is incremental and skips filesystem links. Outputs retain the full relative source filename plus the target extension, so `report.doc` and `report.docx` have distinct PDF names. Explicit files retain relative paths when `InputDirectory` supplies their common root; otherwise they use their filenames, and destination collisions follow the selected policy.

## Book publishing

`BookManuscriptImporter` composes the owning Word, Markdown, HTML and EPUB libraries.
It imports `.docx`, `.md`, `.markdown`, `.html` and `.htm` manuscripts into reflowable
books, retaining each conversion stage's fidelity categories. Word uses its static
final revision view: comments are omitted, fields use their stored visible results,
and live controls are outside the book contract. Notes and supported semantic content
remain in the publication. Markdown front matter supplies title, language and author.

```csharp
using OfficeIMO.Epub;
using OfficeIMO.Workflows;

var imported = await BookManuscriptImporter.ImportFileAsync("manuscript.md",
    new EpubManuscriptOptions { ChapterHeadingLevel = 1 });
BookProject project = BookProject.FromImport(imported);
imported.Report.RequireNoLoss();
project.RenameChapter(0, "Opening chapter");
project.SetStylesheet(EpubTypography.CreateStylesheet(EpubTypographyProfile.Prose));
await File.WriteAllBytesAsync("book.oibook", project.ToProjectBytes());
BookProject reopened = BookProject.LoadProject(await File.ReadAllBytesAsync("book.oibook"));
await File.WriteAllBytesAsync("book.epub", reopened.Export().Bytes);
```

The file route allows at most 64 MiB of manuscript input and resolves assets only
inside the manuscript's physical parent directory. It rejects executable Word
package parts. `ImportBytesAsync` consumes a host-owned snapshot and does not read
files implicitly; a host may supply a permission-aware resource resolver and base URI.
Typed `ImportWordAsync` and `ImportMarkdownAsync` reuse an already loaded source.

`BookProject` owns validated edits: metadata, chapter insertion/removal/reordering,
chapter titles and XHTML bodies, resource renaming with reference repair, a project stylesheet, and cover selection.
`RenameResource(manifestId, containerPath)` delegates to the EPUB owner and retains the same undo/redo behavior as other project edits. Its resource-inspection limits are described in the EPUB README.

`SplitChapter(manifestId, boundaryId, newManifestId, newContainerPath, title)` delegates the atomic chapter split to the EPUB owner. The resulting content, navigation and reading-order changes participate in project undo/redo and persistence. See the EPUB README for supported boundaries and reference-repair limits.

`MergeChapters(firstManifestId, secondManifestId, boundaryId)` combines consecutive compatible chapters through the same owner and undoable transaction. Both navigation entries survive, and the second chapter's links target retained content or its new boundary. Conflicting styles, identifiers and metadata require explicit resolution; see the EPUB README for the merge contract.
`ApplyContentEdits` accepts the EPUB owner's stale-checked element proposals and commits them as one undoable edit. Named revisions can capture the book before or after this transaction. See the EPUB README for element selection and batch validation.

`ApplyEdits` commits a complete editor draft atomically. Invalid or cancelled edits
retain the previous publication. Deleting a linked chapter requires repairing its
remaining links first. A blank creator retains the current creator. Package edits
have one bounded session-only undo/redo step. Named revisions are saved separately.
`PreviewChapter` renders through `OfficeIMO.Epub.Image` using retained package assets.
It selects the requested spine position and fails if that chapter was omitted by the
bounded reading policy. Navigation edits also reject incomplete reader projections,
retaining the complete publication when item, depth, or XML size limits are reached.

Capture editorial milestones explicitly with `CreateRevision(name)`. `Revisions`
returns immutable descriptors with an ID, name, UTC timestamp, SHA-256 hash and byte
count. `RestoreRevision(id)` restores publication content as an undoable edit;
`RemoveRevision(id)` removes only that snapshot. Import diagnostics and acceptance
remain project-wide. Direct edits through `Publication` are included when capturing
a revision, but are not automatically recorded as history.

```csharp
var baseline = project.CreateRevision("Before copyediting");
project.SetMetadata("Revised title", "en", "Author");
byte[] saved = project.ToProjectBytes();
var restored = BookProject.LoadProject(saved);
restored.RestoreRevision(baseline.Id);
restored.Undo(); // Return to the revised title.
```

Use `CompareRevision(id)` to compare a retained revision with the current book.
The EPUB owner reports metadata, reading-order, resource and source-text changes;
`TextChanges` contains bounded before/after excerpts. Comparison does not modify the
project, undo state or revision history. See the EPUB README for matching rules and
coverage limits.

A project retains up to 100 named revisions and 128 MiB of combined revision EPUB
bytes, in addition to its current publication (up to 128 MiB). Capture rejects an
exhausted bound without evicting existing revisions. Version-2 projects retain these
snapshots; version-1 projects remain readable. Session undo/redo is not persisted.
Revision hashes detect inconsistent stored content; they are not digital signatures.

The `.oibook` container stores `publication.epub`, named revision EPUBs and a versioned review record, with
physical ZIP validation and byte/count limits. Loading never extracts files.
Projects may retain non-fatal review findings until the author acknowledges them;
failure diagnostics cannot be accepted as export-ready. Every EPUB export still runs
the native writer's validation. Project review records are user-owned state, not an
authenticity certificate. Project instances are mutable and not thread-safe.
Hosts own destination permissions, conflict handling and safe publication; Studio
uses its existing verified storage owner for those operations.

## Batch book publication

The built-in `book-project-epub` route exports `.oibook` files through `BookProject`
and the EPUB writer. It uses the same runner, publication guards, conflict policies,
byte limits and checkpoint recovery as other conversions.

```csharp
OfficeConversionBatchResult published = await OfficeWorkflow.ConvertDirectory("BookProjects")
    .ToDirectory("PublishedBooks", ".epub")
    .SelectExtensions(true, ".oibook")
    .WithLimits(BookProject.MaximumProjectBytes, 128L * 1024 * 1024)
    .WithCheckpoint("PublishingState")
    .RunAsync(cancellationToken: cancellationToken);
```

Outputs retain source names, such as `novel.oibook.epub`. Export requires saved author
acknowledgment of non-fatal import losses; failures always block it. The batch does
not acknowledge losses automatically. Accepted import findings and writer fidelity
findings remain structured workflow diagnostics. Only the current publication is
exported; named revisions and review records stay in the source project. Each staged
EPUB is reopened and checked by its owner before publication. This native check does
not replace EPUBCheck, accessibility review or independent-reader qualification.

For an individual in-memory project, `Export(EpubWriteOptions, cancellationToken)`
applies explicit writer limits while enforcing the same import-review gate.

## Book delivery bundles

`ToDeliveryBytes` packages the current reviewed publication for a local handoff:

```csharp
byte[] delivery = project.ToDeliveryBytes(
    new EpubWriteOptions { ModifiedAt = new DateTimeOffset(2026, 10, 5, 12, 0, 0, TimeSpan.Zero) },
    cancellationToken: cancellationToken);
// The host owns destination permissions, conflict handling and safe file publication.
```

The ZIP contains exactly `publication.epub`, `package.opf` and `manifest.json`.
The OPF is copied byte-for-byte from the selected package in the exported EPUB,
including identifiers, refinements, extensions and the writer's modification time.
Its resource references remain relative to the original `PackagePath` recorded in
the manifest; the sidecar is for metadata inspection, not standalone rendering.

The version-1 `OfficeIMO.BookDelivery` JSON manifest records each payload's name,
media type, byte length and uppercase SHA-256, plus import acknowledgment and
separate import/writer diagnostics. Diagnostic loss categories are strings.
Diagnostics can include source locations and author-supplied text. Named revisions,
undo history and project review files are excluded. Hashes detect payload changes;
they do not authenticate the sender or protect a manifest that is replaced with
the payloads.

Delivery uses the same import-review and signature policies as `Export`. The native
writer check is recorded as `passed`; independent validation and retailer acceptance
are recorded as `not-performed`. The bundle performs no ONIX conversion, external
validation, upload or retailer-specific packaging.

The EPUB is limited to 128 MiB, its OPF sidecar to 4 MiB, and the manifest to 1 MiB
and 10,000 diagnostics per stage. `maximumOutputBytes` can lower the 134 MiB delivery
ZIP ceiling. Writer limits remain effective and supplied options are not modified.
Exceeded bounds or cancellation return no bundle and leave the project intact.
Unchanged content and identical writer options produce identical delivery bytes
on the same runtime. The API returns bytes and does not write destination files.

## ONIX bibliographic export

`ExportOnix` creates one ONIX 3.1 reference-tag product record for a single EPUB
digital download. It takes a selected title (the first by default) and an explicitly
selected ISBN-13 from the exported package. Supply the message identity, notification intent,
language, publisher and credits explicitly:

```csharp
// schemas is a compiled XmlSchemaSet loaded from vetted ONIX 3.1 reference XSD files.
var options = new BookOnixExportOptions {
    SenderName = "Example Press", RecordReference = "digital-edition-42",
    SentAt = new DateTimeOffset(2026, 10, 5, 12, 0, 0, TimeSpan.Zero),
    Notification = BookOnixNotification.Confirmed,
    IdentifierId = "digital-isbn", LanguageCode = "eng", PublisherName = "Example Press",
    PublicationDate = new DateOnly(2026, 10, 5),
    Contributors = [new("Alice Example", BookOnixContributorRole.Author),
                    new("Example Studio", BookOnixContributorRole.Illustrator, IsOrganization: true)]
};
BookOnixExportResult record = project.ExportOnix(options, schemas, cancellationToken: cancellationToken);
// record.Bytes is ONIX XML; record.Publication contains the exact EPUB and its writer report.
```

The selected `dc:identifier` must exist and contain a valid ISBN-13. The shared
ISBN validator checks spelling, prefix and checksum; it does not verify registration
or that an ISBN was allocated to this digital edition. Language codes use ONIX list
74 (`eng`, `pol`, `fre`), not BCP 47 tags. The supplied schema checks code membership.
Credits distinguish people from organizations and support author, editor,
translator, illustrator and other creative responsibility. Supply 1–100 credits,
or explicitly set `NoContributors = true` with an empty list. Missing EPUB credits
are not interpreted as an assertion that there are no contributors.

Early, advance and confirmed notifications are **complete-record replacements**.
Commercial and accessibility blocks are emitted only when supplied. This profile
has no series, subject, description or retailer-specific blocks.
Do not use it to update an existing richer trade record unless replacing that record
with this profile is intended. Subtitle and publication date are optional explicit
values; all other EPUB metadata stays in the EPUB and is not automatically mapped.
Block updates, deletion records and multi-product messages are outside this profile.

Add accessibility discovery information as explicit publisher assertions:

```csharp
record = project.ExportOnix(options with {
    Accessibility = new BookOnixAccessibilityMetadata {
        Summary = "Language-tagged text and labelled links; some diagrams lack extended descriptions.",
        Status = BookOnixAccessibilityStatus.Limited,
        Features = [BookOnixAccessibilityFeature.LanguageTagging,
                    BookOnixAccessibilityFeature.ClearLinkPurposes],
        AssessmentDate = new DateOnly(2026, 10, 5),
        PublisherInformationUrl = "https://example.org/accessibility/edition-42",
        PublisherContactEmail = "accessibility@example.org"
    }
}, schemas);
```

These declarations use `ProductFormFeatureType` 09 and the supported
[ONIX list 196](https://ns.editeur.org/onix/en/196) values. Feature assertions describe
the edition as a whole. The typed API restricts supported code values; the ONIX
XSD checks their XML placement but does not enforce list 196 membership or verify
the truth of an assertion. A chapter-only TOC does not establish complete TOC navigation,
and partial narration does not establish synchronized audio for substantially all
text. EPUB metadata and validator passes never populate these fields automatically.

After an appropriate assessment, callers can explicitly supply
`Conformance = new(BookOnixWcagVersion.V2_2, BookOnixWcagLevel.AA)` to declare EPUB
Accessibility 1.1 plus that WCAG version and level. This is a publisher assertion,
not certification by OfficeIMO. Unknown accessibility cannot also assert conformance.
Legal exemption claims are outside this export profile.

Record certification provenance separately from conformance:

```csharp
var assessed = new BookOnixAccessibilityMetadata {
    Certification = new("Edition certifier", "https://certifier.example.org/scheme") {
        CredentiallingOrganizationName = "Credentialling organization",
        CredentiallingOrganizationUrl = "https://credentials.example.org/"
    },
    IndependentReportUrl = "https://certifier.example.org/reports/edition-1",
    IntermediaryInformationUrl = "https://intermediary.example.org/edition-1",
    CompatibilityReportUrl = "https://publisher.example.org/compatibility/edition-1",
    IntermediaryContactEmail = "accessibility@intermediary.example.org"
};
record = project.ExportOnix(options with { Accessibility = assessed }, schemas);
```

`Certification` requires a certifier name and organization or scheme URL (90/93);
its optional credentialling organization uses 88/89. Product-specific independent
reports use 94, trusted intermediary information uses 95, publisher information uses
96, compatibility testing reports use 97, and intermediary contacts use 98. A report
can describe an assessment without certification, so report URLs do not require a
certifier. Neither certification provenance nor reports add a conformance declaration.
The example uses fictional organizations; supply only claims supported by the actual
edition's assessment.

Omit `Accessibility` when no assertions are supplied; an empty declaration is rejected.
Feature lists must be distinct and contain at most 32 supported values. Text fields
are limited to 4096 characters, URLs must be absolute HTTP(S) without credentials,
and contacts must be plain email addresses. Export does not fetch URLs, send email,
verify assessments or derive an assessment date from the transmission timestamp.

Add explicit commercial metadata when the record must describe a market offer:

```csharp
var uk = new BookOnixTerritory { Countries = ["GB"] };
record = project.ExportOnix(options with {
    Commercial = new BookOnixCommercialMetadata {
        PublishingStatus = BookOnixPublishingStatus.Active,
        SalesRights = [new(BookOnixSalesRightsKind.Exclusive, uk)],
        Supplies = [new() {
            Territory = uk, SupplierName = "Example Press",
            SupplierRole = BookOnixSupplierRole.PublisherToCustomers,
            Availability = BookOnixAvailability.Available,
            Prices = [new() { Kind = BookOnixPriceKind.RecommendedIncludingTax,
                             Amount = 9.99m, CurrencyCode = "GBP" }]
        }]
    }
}, schemas, cancellationToken: cancellationToken);
```

Territories use explicit country lists or `Worldwide = true` with optional
`ExcludedCountries`. Codes are uppercase ONIX list 91 values and the supplied schema
checks membership. Rights territories cannot overlap in this profile. Available,
forthcoming or temporarily unavailable supply must fit the union of declared
for-sale rights; no finite country list is treated as worldwide permission.
Unavailable or withdrawn supply can be reported after rights are lost. Undeclared
territories remain unstated, and OfficeIMO does not verify rights ownership.

Each supply requires prices or an explicit `Unpriced` reason: `Free`,
`ToBeAnnounced` or `ContactSupplier`. A zero price does not mean free. Prices preserve
positive decimal amounts without rounding or currency conversion and distinguish
recommended, fixed, supplier-net and publisher agency price bases and tax inclusion.
Currency codes use ONIX list 96. Price territories default to the declared supply
market and may narrow it; they cannot broaden it. Optional `ValidFrom` and `ValidUntil`
dates preserve the supplied effective period and reject a reversed interval.
Overlapping-offer resolution is not supported by this profile.
Jurisdiction-specific business rules are not calculated or qualified.

Forthcoming publishing status requires `PublicationDate`; cancelled or indefinitely
postponed status forbids it. Supplier availability has its own date: not-yet-available
or temporarily unavailable supply requires `ExpectedSupplyDate`, or the explicit
`ExpectedSupplyDateUnknown` exception when no date is known. Other availability states
do not accept that expected-date declaration. Product publishing status and supplier
availability are separate assertions.

Use `BookOnixPrice.Discounts` for up to 16 explicit business-to-business discount
declarations. The containing price's amount, currency, territory and effective dates
remain unchanged:

```csharp
var price = new BookOnixPrice {
    Kind = BookOnixPriceKind.RecommendedExcludingTax,
    Amount = 10.00m,
    CurrencyCode = "GBP",
    Discounts = [
        new() { Kind = BookOnixDiscountKind.Rising, MinimumQuantity = 1,
            MaximumQuantity = 9, Percent = 10.00m },
        new() { Kind = BookOnixDiscountKind.Rising, MinimumQuantity = 10,
            Percent = 20.00m }
    ]
};
```

Discount kinds follow [ONIX list 170](https://ns.editeur.org/onix/en/170). Each
needs a percentage from 0–100, a nonnegative amount per copy no greater than the
price, or both. Decimal scale is preserved. Optional copy quantities must be positive
integers; a maximum requires a minimum and cannot precede it. Omitting quantities
leaves them unspecified. An empty list makes no discount assertion; an explicit zero
is retained. A zero percentage cannot accompany a positive amount.

The exporter preserves declaration order. It does not resolve overlapping tiers,
select a discount, stack declarations, check percentage/amount arithmetic, calculate
net prices or recalculate taxes. Cumulative periods and purchaser eligibility remain
trading-partner agreements; price effective dates are not a cumulative-order period.

For trading-partner codes, use `BookOnixPrice.DiscountCodes`:

```csharp
var codedPrice = new BookOnixPrice {
    Kind = BookOnixPriceKind.RecommendedExcludingTax,
    Amount = 10.00m,
    CurrencyCode = "GBP",
    DiscountCodes = [new() {
        Scheme = BookOnixDiscountScheme.ProprietaryDiscount,
        SchemeName = "Example Press trade terms",
        Code = "A1"
    }]
};
```

The seven schemes follow [ONIX list 100](https://ns.editeur.org/onix/en/100).
Proprietary discount and commission schemes require a distinctive `SchemeName`;
other schemes omit it. Up to 16 codes are retained in supplied order, before numeric
discounts. Codes and scheme names use the existing 4096-character text limit.
BIC codes require five ASCII letters followed by one to three alphanumeric
characters. ISNI-based codes require 15 digits and a final digit or `X`, followed
by a hyphen and one to three alphanumeric characters. Input is preserved verbatim.
These checks do not verify ISNI checksums or allocation, BIC prefix ownership,
local terms-code membership, partner eligibility, or code-to-rate mappings.
Discount and commission codes keep their distinct scheme values; neither changes
the price or creates a numeric discount automatically.

Tax-inclusive prices can carry up to 16 explicit `BookOnixTax` components:

```csharp
var price = new BookOnixPrice {
    Kind = BookOnixPriceKind.RecommendedIncludingTax,
    Amount = 10.50m,
    CurrencyCode = "PLN",
    Taxes = [new() {
        Type = BookOnixTaxType.ValueAdded,
        RateCode = BookOnixTaxRateCode.Lower,
        RatePercent = 5.00m,
        TaxableAmount = 10.00m,
        Amount = 0.50m
    }]
};
```

These are illustrative publisher-supplied values, not a tax-rate recommendation.
Types follow [ONIX list 171](https://ns.editeur.org/onix/en/171); optional rate
classifications follow [list 62](https://ns.editeur.org/onix/en/62). Each component
needs a percentage, a tax amount, or both, and can describe the affected price part.
Values retain decimal scale and use the price's currency and territory. Percentages
are bounded to 0–100, taxable amounts must be positive, and tax amounts cannot be
negative. Each supplied taxable amount plus its tax must fit within the price;
the sum of supplied tax amounts cannot exceed it. Zero-rate assertions cannot carry
positive tax amounts. Different taxes may share a taxable base, so taxable bases
are not summed.

Use `TaxExempt = true` with an empty `Taxes` list to assert exemption. A zero-rated
tax uses `BookOnixTaxRateCode.Zero`; an empty list alone leaves tax information
unstated. Tax components require a tax-inclusive price kind. OfficeIMO does not
calculate missing values, reconcile percentage arithmetic or rounding, determine
applicable laws, or verify that a rate classification applies to the market.

Commercial metadata is bounded to 32 rights declarations, 32 supplies, 16 prices per
supply and 250 unique codes per country list. The total XML byte limit still applies.

The host supplies and owns the provenance of a compiled `XmlSchemaSet` declaring
`ONIXMessage` in `http://ns.editeur.org/onix/3.1/reference`. OfficeIMO ships no ONIX
schema files and fetches none during export. The [opt-in evidence runner](../Build/Epub/README.md#onix-fixtures)
shows loading with a resolver restricted to three local schema files. The exporter
enforces that schema and rejects validation findings; schema validation alone does
not establish business-rule completeness or recipient acceptance.

The same import-review and EPUB signature gates apply as for `Export`. Results retain
import findings, acknowledgment, writer findings and an EPUB SHA-256 captured at
export time. They perform no upload or independent EPUB validation. Options do not
change project metadata or history. Text fields are limited to 4,096 characters,
the inspected OPF to 4 MiB and ONIX XML to 1 MiB. EPUB writer limits and signature
policy can be supplied separately through `epubOptions`. Project instances, input
credit/commercial lists and schema sets must not be mutated concurrently with export.

### Title selection and subjects

`TitleId` selects a `dc:title` by its OPF identifier from the exact exported EPUB.
Omitting it retains the first-title default. Missing identifiers are rejected;
selection does not change EPUB metadata. `Subtitle` remains an explicit assertion.

`AlternativeTitles` adds up to 32 publisher-supplied product-level titles after the
selected distinctive title. Declare their meaning explicitly: original-language,
abbreviated, parallel-language, former, distributor, cover, back-cover, expanded,
widely known alternative, spine, or intermediate translation title. Serial-only
title types are outside this book profile. Multiple titles of the same type are
allowed, for example parallel titles in different languages; caller order is retained.

```csharp
var translated = existingOptions with {
    AlternativeTitles = [
        new() { Type = BookOnixAlternativeTitleType.OriginalLanguage,
            Title = "L’histoire", LanguageCode = "fre", TitleSorting = new() { Prefix = "L’" } },
        new() { Type = BookOnixAlternativeTitleType.OtherLanguage,
            Title = "Historia", LanguageCode = "pol" }
    ]
};
```

Each alternative may carry a subtitle and an ONIX list 74 language code, applied to
its title, sorting components and subtitle. Omitted language remains unspecified.
These are separate ONIX assertions: they neither change EPUB metadata nor infer
translation history or territorial applicability. The supplied schema validates
codes, while retailer acceptance requires recipient-specific qualification.

`TitleSorting` distinguishes an unknown prefix from a publisher's explicit sorting
instruction. It is available on `BookOnixExportOptions`, a simple `BookOnixCollection`,
`BookOnixAlternativeTitle`, and each `BookOnixCollectionTitleElement`:

```csharp
var prefixed = existingOptions with {
    // For a selected EPUB title such as "The history of publishing":
    TitleSorting = new() { Prefix = "The " }
};
var unprefixed = existingOptions with { TitleSorting = new() };
```

Omitting `TitleSorting` retains the unsplit `TitleText` output. `new()` asserts
`NoPrefix` and writes the full title as `TitleWithoutPrefix`. Supplying `Prefix`
writes it as `TitlePrefix` and removes that exact leading text from
`TitleWithoutPrefix`. Include any separator to remove in the prefix, such as the
space in `"The "`. Matching is ordinal and case-sensitive; text, punctuation and
remaining whitespace are preserved. Prefix plus remainder reconstructs the full
title exactly. Empty, whitespace-only, mismatched and exhaustive prefixes are
rejected. OfficeIMO does not infer sorting rules from language or strip articles
automatically.

`BookOnixExportOptions.TitleSorting` applies to the selected title from the exported
EPUB. Alternative titles use their own declared text and language. For collections,
the declared title language applies to both prefix and remainder.
Sorting cannot be attached to a part-only element. With `TitleElements`, place
sorting on each element, not on the collection's simple-title fields. The choice
is explicit because `TitleText` is deprecated but still accepted in ONIX 3.1;
existing callers retain their output until they supply a sorting assertion.

Declare discoverability metadata in `Subjects`:

```csharp
var options = existingOptions with {
    TitleId = "catalog-title",
    Subjects = [
        new() { Scheme = BookOnixSubjectScheme.Bisac, Code = "JUV000000", IsMain = true },
        new() { Scheme = BookOnixSubjectScheme.Thema, Code = "YFB", IsMain = true,
            SchemeVersion = "1.5",
            Headings = [new("Children's fiction", "eng"), new("Literatura dziecięca", "pol")] },
        new() { Scheme = BookOnixSubjectScheme.Keywords,
            Headings = [new("stories; adventure", "eng")] }
    ]
};
```

The supported [ONIX list 27](https://ns.editeur.org/onix/en/27) schemes are Dewey,
Library of Congress classification and headings, BISAC, keywords, named proprietary
schemes, Thema categories and its six qualifier schemes. Each declaration needs a
code or heading; keywords use heading text only. Use semicolon-separated keywords
in one heading per language. Proprietary schemes require `SchemeName`; standard
schemes omit it. Optional versions are preserved verbatim.

Export accepts up to 64 subjects with 16 headings each. A subject cannot repeat a
heading language, including unspecified language. Heading language codes use ONIX
list 74 rather than EPUB BCP 47 tags. At most one subject per scheme (and name for
proprietary schemes) can be main; keywords and Thema qualifiers cannot be main.
Text fields use the existing 4096-character bound and the total XML limit still
applies. Subject authority strings in the EPUB are not automatically mapped. Schema
validation checks ONIX structure and list values, not membership of individual
subject codes, supplied scheme versions, classification suitability or discoverability
in a recipient's catalog.

### Collateral text and attribution

`CollateralTexts` carries publisher-supplied descriptions, review quotes, excerpts,
cover copy and other supporting text. The recipient of this copy is separate from
the readership of the book:

```csharp
var options = existingOptions with {
    CollateralTexts = [new() {
        Type = BookOnixTextType.ReviewQuote,
        Audiences = [BookOnixContentAudience.EndCustomers],
        Texts = [new("A supplied review quotation.", "eng")],
        Authors = ["Example Reviewer"],
        SourceCorporate = "Example Journal",
        SourceTitles = [new("Review of the book", "eng")],
        SourceLinks = ["https://example.org/review"],
        PublishedOn = new DateOnly(2026, 9, 1),
        UsableFrom = new DateOnly(2026, 10, 1)
    }]
};
```

The profile supports [ONIX list 153](https://ns.editeur.org/onix/en/153) text types
02–19: short/full product descriptions, table of contents, cover copy, current or
previous-edition/work review quotes, endorsement, headline, feature, biographical
note for all contributors, publisher notice, excerpt, index, short/full collection
descriptions, new feature and version history. Each item requires explicit
[list 154 recipients](https://ns.editeur.org/onix/en/154). Recipient codes must be
distinct; `Unrestricted` cannot accompany another code. List order becomes the
collateral sequence order.

Text defaults to plain text (`textformat="06"`). Markup-like input remains
literal unless the variant explicitly selects `Format = BookOnixCollateralTextFormat.Xhtml`;
links are never fetched. Each item supports up to 16 language variants,
16 authors, 16 source-title variants and 16 source links. Language codes use ONIX
list 74, with distinct languages per variant list including unspecified. Source
links require absolute HTTP(S) URLs without credentials. Optional `Territory`
reuses the existing country/worldwide profile and describes use of the collateral,
independently of product sales rights.

There may be at most 64 items. Each text variant accepts up to 65,536 UTF-16 code
units, and texts plus source titles share a 524,288-unit export budget. Source
titles and other attribution fields retain the 4096-unit field bound. Short product
and collection descriptions additionally permit at most 350 Unicode scalar values,
so a supplementary character counts once. The complete ONIX document still has its
1 MiB serialized size limit.

`PublishedOn` and `UpdatedOn` describe the collateral. `UsableFrom` and `UsableUntil`
carry its permitted-use dates; reversed intervals are rejected. These dates,
restricted-recipient labels and territory declarations are metadata assertions,
not access controls: export includes the text and does not enforce embargoes or
filter a recipient's copy. The publisher remains responsible for accurate attribution,
permission to use the text and recipient acceptance. Review ratings, media resources and license terms are outside this collateral-text profile. No text or
attribution is inferred from EPUB content, and export does not change the book.

#### XHTML collateral variants

Use an explicit fragment when the description needs semantic formatting:

```csharp
var variant = new BookOnixCollateralTextValue(
    "<p>A <strong>formatted</strong> description.</p><ul><li>A feature</li></ul>", "eng") {
    Format = BookOnixCollateralTextFormat.Xhtml
};
```

XHTML variants emit `textformat="05"`. Input must be well-formed XML fragments,
not HTML requiring parser repair. Multiple elements and surrounding text are allowed.
Elements without an explicit namespace, in the standard XHTML namespace, or in the
ONIX reference namespace are normalized to the ONIX namespace used by its XHTML
subset schema. `xml:lang` becomes the subset's `lang` attribute; conflicting values
are rejected. The outer variant language remains an ONIX list 74 code. This is
semantic serialization, not byte-for-byte preservation of the original markup.

The authoring profile supports paragraphs/divisions, headings, ordered/unordered and
definition lists, block quotations, preformatted text, links, line breaks, rules,
common emphasis and code/phrase elements, and tables with captions, row groups and
column groups. Allowed attributes are `title`, `lang`, `dir`, link `href`/`hreflang`,
quotation `cite`, ordered-list `start`/`type`, cell `colspan`/`rowspan`, header-cell
`scope`, and column/group `span`, subject to the supplied schema's element rules.
Links and citations must be absolute HTTP(S) URLs without credentials. The profile
rejects styles/classes, identifiers/local anchors, images, active content, event
handlers, foreign elements/attributes, comments, processing instructions and DTDs.
No entities or resources are retrieved. Use numeric references or XML's built-in
entities; HTML-only entities such as `&nbsp;` are not defined.

Fragments need nonblank text and are bounded to 32 nested elements and 4096 XML
nodes before tree materialization. Existing source-text length and aggregate limits
include markup. For short descriptions, the 350-scalar limit counts decoded text,
including supplied whitespace, while excluding markup. `SourceTitles` remains plain
text only. Full-schema validation checks nesting and attribute values; recipient
rendering and acceptance are separate qualifications.

### Audience categories, ages and school grades

Use `Audience` for explicit readership assertions:

```csharp
var options = existingOptions with {
    Audience = new() {
        Categories = [new(BookOnixAudienceType.Children, IsMain: true),
                      new(BookOnixAudienceType.Teenage)],
        AgeRanges = [new(BookOnixAgeRangeType.InterestYears, Minimum: 10, Maximum: 14),
                     new(BookOnixAgeRangeType.ReadingYears, Minimum: 9, Maximum: 12)],
        GradeRanges = [new(BookOnixGradeSystem.UnitedStates, BookOnixGrade.Grade5, BookOnixGrade.Grade8)],
        Descriptions = [new("Readers of adventure and exploration", "eng")]
    }
};
```

All 13 [ONIX list 28](https://ns.editeur.org/onix/en/28) categories are supported.
Categories must be distinct, with at most one main audience. Descriptions are plain
text, not HTML, with optional ONIX list 74 language codes. Up to 16 descriptions
are allowed, with distinct languages including unspecified; text fields retain the
4096-character bound. Omit `Audience` when making no assertion.

Age ranges distinguish interest in years or months from reading age in years.
At least one nonnegative integer bound is required. Equal bounds mean an exact age;
a lone minimum means “from”, a lone maximum means “to”, and different minimum/maximum
bounds form a closed range. Each range type may appear once. Interest months and
interest years cannot coexist. Following [ONIX list 30](https://ns.editeur.org/onix/en/30),
month-based interest ages allow a first value up to 36 and a second value up to 42.
Thus 36–42 months is valid, while an exact age of 42 months or a lone upper bound of
42 months is not.

`GradeRanges` adds school and college levels for `UnitedStates` (qualifier 11),
`CanadaExcludingQuebec` (26), and `China` (29). The first two use
[ONIX list 77](https://ns.editeur.org/onix/en/77); China uses
[list 227](https://ns.editeur.org/onix/en/227). Values are `Preschool` (P),
`Kindergarten` (K), then `Grade1` through `Grade17`, in that order. Their educational
meaning depends on the selected system; grades 13–17 denote tertiary levels.
The Canadian profile does not represent Québec's grading system.

Each system may appear once, with at least one bound. Equal bounds express an exact
grade; a lone minimum means “from” and a lone maximum means “to”. Closed ranges
must follow grade order, including preschool before kindergarten before grade 1.
Age and grade ranges may coexist. No age-to-grade conversion is performed.

Audience categories, age ranges and grade ranges are independent assertions. Supply an appropriate
range for children's, teenage and school material when known; export does not guess
one from a category or inspect the book to assess suitability. This profile does not
represent proprietary/national audience codes, other national grade schemes, adult-content
ratings or reading-complexity schemes. Schema validity does not establish educational
suitability or recipient acceptance, and audience export does not change EPUB metadata.

### Collection membership

Declare collection identity and ordering explicitly. These fields do not change or
automatically project EPUB series metadata:

```csharp
var options = existingOptions with {
    Collections = [new() {
        Type = BookOnixCollectionType.Publisher,
        Title = "Collected studies",
        LanguageCode = "eng",
        Identifiers = [new(BookOnixCollectionIdentifierType.Proprietary,
            "studies", "Publisher catalog")],
        Sequences = [new(BookOnixCollectionSequenceType.Publication, "3"),
                     new(BookOnixCollectionSequenceType.Narrative, "2.1")],
        Contributors = [new("Alex Editor", BookOnixContributorRole.SeriesEditor)]
    }]
};
```

Up to 32 named collections are supported. Each can declare a subtitle, an ONIX list
74 title language, up to 16 identifiers and up to 16 sequence positions. Collection
types distinguish publisher series/sets, collections éditoriales, and ascribed
collections. An ascribed collection requires the defining party's `SourceName`.

Identifiers support named proprietary schemes, ISSN and ISBN-13. ISSN shape and
check digits are checked; an optional central hyphen is removed and a final `x` is
uppercased. ISBN normalization uses the shared publishing validator. Checks establish
neither identifier allocation nor ownership. Use a collection ISBN only when the
collection is available as a single product. Each identifier type may occur once per
hierarchy level, except that different named proprietary schemes may coexist.
Set an identifier's `Level` when it identifies a particular level of a hierarchy;
that level must be present in the title. Scoped and unscoped identifiers for the
same type and proprietary scheme cannot coexist.

Sequence types cover title, publication, narrative, original publication, suggested
reading, suggested display and named proprietary ordering. Positions retain their
text: `2.1` is hierarchical, and `3.-.8` can omit an intermediate level. Components
must be ASCII digits or a hyphen, separated by dots. Each sequence type/name may
occur once. Named proprietary identifiers and sequences require a name; standard
types omit it. All text fields retain the 4096-character bound.

Each collection can carry up to 100 ordered `Contributors`, using the same person or
organization credit model as product credits. `SeriesEditor` writes ONIX role `B09`.
Collection credit numbering starts at 1 for each membership. Credits are never copied
between the product, collections and EPUB metadata; place them according to the
recipient's requirements. An empty list makes no assertion. `NoContributors = true`
explicitly asserts no collection contributors and cannot accompany credits.

For structured collection titles, use `TitleElements` instead of the simple `Title`, `Subtitle`,
`LanguageCode` and `TitleSorting` fields:

```csharp
var collection = new BookOnixCollection {
    Type = BookOnixCollectionType.Publisher,
    TitleElements = [
        new() { Level = BookOnixCollectionLevel.Collection,
                Title = "Collected studies", LanguageCode = "eng" },
        new() { Level = BookOnixCollectionLevel.Subcollection,
                Title = "Historical studies", PartNumber = "Series II", LanguageCode = "eng" },
        new() { Level = BookOnixCollectionLevel.SubSubcollection,
                PartNumber = "Part 3" }
    ],
    Identifiers = [new(BookOnixCollectionIdentifierType.Issn, "0317-8471") {
        Level = BookOnixCollectionLevel.Subcollection
    }],
    Frequency = BookOnixCollectionFrequency.Annual
};
```

The list holds one element per represented level, up to five. Subcollections and
sub-subcollections must include their series parent levels. List order determines display order and writes consecutive
`SequenceNumber` values; it can differ from hierarchy order. Each series level needs
a title, a part designation (including its caption), or both. Its optional language
applies to its title, subtitle and part designation; languages are not inherited
between levels. The product title remains separate. Collection-level alternative
title types are outside this profile.

`MasterBrand` and `Universe` add explicit ONIX levels `05` and `07` to `TitleElements`.
They require a nonblank `Title` and can appear alone or alongside series levels.
Neither is treated as a parent of the series or of the other identity. For example:

```csharp
var branded = new BookOnixCollection {
    Type = BookOnixCollectionType.Publisher,
    TitleElements = [
        new() { Level = BookOnixCollectionLevel.MasterBrand, Title = "Voyager Tales" },
        new() { Level = BookOnixCollectionLevel.Universe, Title = "Orbital Commons" },
        new() { Level = BookOnixCollectionLevel.Collection, Title = "Early readers" }
    ]
};
```

This profile writes these identities in `Collection` composites, keeping the product
title separate. Use separate `Collections` entries for independent brands or universes
at the same level. Identifiers can select either level when it appears in that entry's title;
existing scheme and scope checks apply. These are publisher-supplied identities,
not inferred licensing, ownership, or associations from EPUB text.

`Frequency` declares the schedule of successive products in the collection. It
supports the [ONIX list 259 values](https://ns.editeur.org/onix/en/259), including
irregular, explicitly unknown and no future publications. Omitting it makes no
schedule assertion. `TwiceYearly` and `EveryTwoMonths` distinguish two from six
publications per year; `MoreOftenThanWeekly` includes daily publication. Frequency
is not inferred from publication dates and does not assert product availability.

`NoCollection = true` explicitly asserts no collection membership and cannot accompany
`Collections`. An empty list with the default `NoCollection = false` makes no assertion.
Collection claims and recipient acceptance remain the publisher's responsibility.
The [ONIX collection types](https://ns.editeur.org/onix/en/148),
[title levels](https://ns.editeur.org/onix/en/149),
[contributor roles](https://ns.editeur.org/onix/en/17),
[identifier schemes](https://ns.editeur.org/onix/en/13) and
[sequence types](https://ns.editeur.org/onix/en/197) define the trade semantics.

### Edition metadata

`Edition` describes the publication edition independently of project revision history:

```csharp
var options = existingOptions with {
    Edition = new() {
        Types = [BookOnixEditionType.Revised, BookOnixEditionType.Annotated],
        Number = 2,
        VersionNumber = "1.2",
        Statements = [new("Second revised and annotated edition", "eng"),
                      new("Drugie wydanie", "pol")]
    }
};
```

Numbered editions use a positive integer. A minor version requires that number and
is preserved as text. Statements are complete display descriptions, serialized as
plain text with optional ONIX list 74 language codes. They are not interpreted as
HTML. Up to 16 statements are allowed, with distinct languages (including unspecified),
and each text field is limited to 4096 characters.

The supported [list 21](https://ns.editeur.org/onix/en/21) characteristics are abridged,
unabridged, annotated, revised, enlarged, illustrated, critical and new. Types must
be distinct; abridged and unabridged conflict. `New` cannot accompany a more specific
type or edition number. Set `Edition = new() { NoEdition = true }` only when explicitly
asserting that no edition information applies; it cannot accompany edition details.
Omitting `Edition` makes no assertion. These values do not change EPUB metadata,
assign a new ISBN, establish the truth of an edition claim, or substitute for a
recipient's business rules.

### Multi-product ONIX messages

Compose exported records into one message with `BookOnixMessage.Create`:

```csharp
// Both options use the same SenderName and SentAt, and select distinct edition ISBNs.
var first = firstBook.ExportOnix(firstOptions, schemas);
var second = secondBook.ExportOnix(secondOptions, schemas);
var message = BookOnixMessage.Create([first, second], schemas);
File.WriteAllBytes("catalog.onix", message.Bytes);
```

The composer preserves product order and complete product XML, then validates the
combined message against the supplied schema. All serialized headers must match;
it does not choose a sender or timestamp on the publisher's behalf. Repeated record
references or ISBNs are rejected. Each result's ONIX and EPUB bytes must still match
its hashes captured at export time.

`message.Products` retains the original export results, including each EPUB,
writer report, import diagnostics and loss acknowledgment. The collection is
snapshotted; its results' byte arrays remain mutable. Do not modify inputs during
composition, and keep exported bytes unchanged when using their recorded hashes.
Composition accepts 1–1000 records, at most 16 MiB of source ONIX XML, and at most
16 MiB of combined XML. Individual exports retain their 1 MiB limit. It performs no
retailer submission, block update, deletion notification or recipient acknowledgment.

## Optional checkpoints

Add `.WithCheckpoint("PDF-State")` to the builder, or set `CheckpointDirectory` on `OfficeConversionBatchRequest`, for restartable execution. Source, output and checkpoint trees must be separate local folders. For selected HTML files and Markdown files with local resources enabled, output and checkpoint folders must also be outside each file's resource tree, including an explicit Markdown `BaseDirectory`. Checkpoint jobs require `Fail`: recorded completed artifacts are immutable and verified by source, rendering-settings, local-resource and output hashes before reuse.

Checkpoints support built-in routes. Registered adapters require an ordinary batch because their captured runtime configuration cannot be fingerprinted. An explicit registered route with checkpoints is rejected before execution; a registered input discovered in a mixed checkpoint batch reports a failed item without publishing it.

Before publication, the runner flushes validated staged output and records its hash and staging identity. Restart can finish that recorded move or verify an output moved before the final receipt was written. Changed completed sources or settings, altered/missing outputs and outputs without a bound receipt fail the item for inspection. `RetryFailed` permits retrying recorded failures, including corrected failed inputs. Completed files and recorded pending publications survive cancellation. An interruption before publication intent is recorded can leave a hidden staging file; inspect it before removing it.

Checkpoint reuse verifies a **recorded artifact**; it does not rerender it or promise that a newer renderer would produce identical bytes. Compatible engine updates do not invalidate completed receipts. The checkpoint schema, host, source/output roots, target and per-item rendering inputs define compatibility. Execution concurrency, selection and byte/file budgets may change; the new budgets still apply when verifying artifacts. New source files are discovered on each run.

HTML and enabled local Markdown resources are conservatively fingerprinted within the source root, with at most 256 regular files and an aggregate input-byte budget. Every filename is included because Markdown identifies image formats from their bytes. Resource trees cannot contain links; checkpointed Markdown also requires `RestrictLocalImagesToBaseDirectory`. A changed CSS, image or font invalidates reuse. Checkpoints exclude remote resources and runtime resource, text-shaping or cryptography callbacks because their output cannot be identified from captured settings; use an ordinary batch or the native adapter for those cases. Workflow HTML resource resolution remains scoped to the source; custom HTML resolvers belong to the native adapter.

Per-document defaults are 64 MiB input and 256 MiB output; concurrency accepts 1–32. `MaximumFiles` bounds discovered files, including skipped files, and defaults to one million. Discovery beyond this bound stops the run while preserving completed output. These are configurable resource bounds, not throughput guarantees. TXT defaults to strict BOM-aware decoding, literal markup, tab expansion and bounded wrapping; `PlainText` carries its encoding and layout limits. Legacy DOC import loss blocks output unless `LegacyDocLossPolicy = OfficeConversionLossPolicy.Allow` accepts reported reductions.

`IProgress<OfficeConversionBatchItemResult>` callbacks can arrive concurrently; consume or stream them without retaining a whole inventory. The result contains bounded counts. Checkpoints retain up to 32 non-information diagnostics and report truncation. A host can supply `publicationGuard` to protect output and checkpoint destinations. Conversion completion retains each adapter's fidelity limits and does not prove exact Microsoft Office pagination.

Studio exposes **Convert → Batch PDF export**. The CLI uses `officeimo workflow batch`; PSWriteOffice uses `Export-OfficeDocumentPdf -InputDirectory ... -OutputDirectory ...` or selected file pipelines on PowerShell 7.4 or newer.

## Email evidence and conversation dossiers

`EmailEvidenceWorkflow` produces a portable ZIP containing `report.html`, `report.md`, `manifest.json`
and an optional `report.pdf`. Reports include From/To/Cc, sent and received dates, attachment indexes,
source fingerprints, protection classification, diagnostics and explicit body clipping. The body is
semantic text from `OfficeIMO.Email.Html`, escaped for display; original formatting and embedded images
are omitted. Attachment payloads and original messages stay outside the ZIP. The workflow reads local
content without network access, signature verification, decryption or certificate discovery.

```csharp
var evidence = EmailEvidenceWorkflow.Create("message.eml");
File.WriteAllBytes("message-evidence.zip", evidence.ToZipBytes());

using var mailbox = OfficeIMO.Email.Store.EmailStoreSession.Open("archive.pst");
var selected = mailbox.EnumerateItems().First().Key;
var dossier = EmailEvidenceWorkflow.CreateConversation(mailbox, selected,
    new EmailEvidenceOptions { MaxItemsScanned = 10_000, MaxMessages = 100 });
File.WriteAllBytes("conversation.zip", dossier.ToZipBytes());
```

Conversation selection reuses the existing graph. Messages are chronological; thread links retain their
evidence and heuristic status, and missing or ambiguous parents remain visible. `GraphComplete` reports
the graph owner's coverage. Each message is projected under the body/report bounds before the next body
is read; eager store formats retain their own bounded opening behavior. Embedded attachments are classified
from available MAPI metadata even when their nested payload is not read. A file fingerprint hashes the same open source before and after parsing;
a store fingerprint uses the store's durable source contract, including its composite directory hash.
Hashes identify source bytes and resident attachment payloads; they do not certify message authenticity.
Deferred attachment streams are not opened for hashing. Report fields use bounded display values,
diagnostics retain a sample of up to 500 entries with the total count, and PDF conversion diagnostics
are included in the manifest. `EmailEvidenceOptions` controls input, body, graph, report, page and output
bounds. Outputs are created in memory; applications choose and authorize their publication destination.

## Invoice inspection, conversion and rendering

`OfficeInvoiceBufferWorkflow` composes the typed invoice model, optional standards
validator and PDF adapter. It accepts captured XML bytes and returns an operation
report without reading paths, fetching invoice links or publishing files:

```csharp
using OfficeIMO.Invoicing;
using OfficeIMO.Workflows;

var target = new InvoiceXmlOptions(
    InvoiceSpecificationRelease.En16931_1_3_16,
    InvoiceSyntax.Ubl,
    InvoiceProfile.En16931);
var request = new OfficeInvoiceWorkflowRequest(
    File.ReadAllBytes("invoice.xml"),
    OfficeInvoiceWorkflowOperation.Convert,
    target,
    inputName: "invoice.xml");
var result = await OfficeInvoiceBufferWorkflow.RunAsync(request);
foreach (var diagnostic in result.Diagnostics)
    Console.WriteLine($"{diagnostic.Location}: {diagnostic.Message}");
if (result.Succeeded)
    File.WriteAllBytes("converted.xml", result.ToOutputBytes()!);
```

Inspection completion does not establish validity. Check `ModelValidation`,
`Source.HasCompleteMapping` and target diagnostics separately. Recognized
MINIMUM and BASIC WL inputs receive aggregate model checks without inventing
invoice lines. Conversion and rendering block unmapped source data and
unsupported target fields; explicit lower-profile projection returns each
intentional reduction as a warning.

Select `Validate`, `RenderPresentationPdf` or `RenderHybridPdf` for the other
operations. Rendering requires an explicit CII contract. Pass `PdfOptions` with
the fonts your content needs and `InvoicePdfLayoutOptions` for appearance and
resource limits. The returned output XML is the same captured invoice used for
the visible PDF; hybrid output embeds those exact bytes.

For bounded header replacements that retain XML extensions, create a request with
`OfficeInvoiceWorkflowRequest.ForSourceEdit(xml, new InvoiceSourceEdits(number:
"INV-002"))`. The file equivalent is
`OfficeInvoiceFileWorkflowRequest.ForSourceEdit("invoice.xml", "edited.xml",
edits)`, which uses the same output preflight and atomic publication contract.
`EditSource` retains the original syntax/profile and accepts no conversion target.
Its `Succeeded` status means every requested edit completed; model and mapping
findings can still contain errors. `Source` and `ModelValidation` describe the
edited XML, or remain unavailable when its retained data exceeds the semantic
mapper. Passing an explicit validation release and validator requires the exact
edited XML to pass schema and business rules before any artifact is returned.
See the [source-editing contract](../OfficeIMO.Invoicing/README.md#read-and-edit-safely)
for supported fields, representations and bounds.

Standards validation requires both `validationRelease` and a configured
`InvoiceValidator` supplied to `RunAsync`. A requested validator that is missing,
or a schema/business-rule stage that does not pass, blocks output. Otherwise
`SchemaStatus` and `BusinessRulesStatus` explicitly report `NotRun`. For writing
operations, `StandardsValidation.Sha256` identifies the output XML validated
before artifact generation; it does not certify the PDF's conformance.

`RunBatchAsync` preflights requests before executing them, preserves input order,
and observes cancellation between bounded owner operations. Defaults are 256
requests, 64 MiB of combined XML input and 64 MiB of retained output artifacts.
`ContinueOnFailure` controls whether subsequent items run. An item exceeding the
output budget returns diagnostics and no artifact bytes. Cancellation throws
`OperationCanceledException`; hosts remain responsible for collision policies
and safe output publication.

`OfficeInvoiceFileWorkflow` provides that local-file adapter. Its immutable
`OfficeInvoiceFileWorkflowRequest` captures paths and the same target and render
settings. It preflights all inputs and destinations, applies the combined batch
budgets, then creates each successful artifact through the shared atomic writer.
Existing destinations and colliding batch outputs are rejected. It never
overwrites inputs or existing files; batch publication is per item rather than a
transaction. `OfficeInvoiceFileWorkflowResult` keeps publication errors separate
from the model and standards evidence in `Workflow`.

For a desktop host with local or provider-backed storage, use
`OfficeWorkflowRunner.RunInvoiceAsync` with `OfficeInvoiceStorageWorkflowRequest`:

```csharp
var storageResult = await new OfficeWorkflowRunner().RunInvoiceAsync(new() {
    InputPath = "invoice.xml",
    Operation = OfficeInvoiceWorkflowOperation.EditSource,
    SourceEdits = new InvoiceSourceEdits(number: "INV-002"),
    OutputPath = "invoice.edited.xml",
    ConflictPolicy = OfficeWorkflowConflictPolicy.Rename
});
```

The adapter captures at most 16 MiB of input, clones render settings before
acquisition, and verifies source contents and physical identity again before
publication. Local output supports fail, numbered-copy and atomic replacement
policies. It protects source aliases and asks the supplied publication guard
about the final destination. Provider inputs use reopenable
`OfficeWorkflowStreamInput`; provider output uses `OfficeWorkflowStreamOutput`
with explicit `Replace` after the host obtains direct-write consent. A durable,
hash-verified XML or PDF recovery copy precedes provider creation/writing. Failed
or unverified provider publication returns `Unconfirmed` with retained recovery;
it cannot promise atomic replacement or rollback. Read `Workflow` for invoice
evidence and `Status`, `Diagnostics` and `Recovery` for storage outcomes. The
default retained output limit is 64 MiB. Cancellation before publication returns
`Cancelled` and removes temporary staging.

## Project reports and table exchange

`ProjectReportWorkflow` exports a calculated Project view through the existing document owners:

```csharp
using OfficeIMO.Project;
using OfficeIMO.Workflows;

using var project = ProjectDocument.Load("delivery.xml");
var schedule = project.CalculateSchedule(new ProjectScheduleOptions { CalculateAssignments = true });
schedule.Report.ThrowIfErrors();
var view = project.CreateView(schedule, new ProjectViewOptions {
    Kind = ProjectViewKind.TaskUsage,
    Timescale = ProjectViewTimescale.Week
});
File.WriteAllBytes("delivery.pdf", ProjectReportWorkflow.ToPdf(view));
File.WriteAllText("delivery.html", ProjectReportWorkflow.ToHtml(view));
using var workbook = ProjectReportWorkflow.CreateExcel(view);
workbook.Save("delivery.xlsx");
```

`ToSvg` and `ToPng` return one result per page. Supply an `OfficeRenderingProfile` with the required fonts for explicit Unicode coverage. `CreateWord` and `CreatePowerPoint` include chart images followed by editable data tables. `CreateExcel` separates report values, usage, status/groups, baseline dates, and dependencies into editable worksheets. These exports do not reconstruct Microsoft Project's saved views or native styles.

Project PNG exports default to 300 DPI, rendered from the drawing at the requested resolution. Choose a shared quality preset for screen images:

```csharp
using OfficeIMO.Drawing;

var images = ProjectReportWorkflow.Images(view)
    .WithQuality(OfficeImageExportQuality.Screen)
    .As(OfficeImageExportFormat.Png)
    .Export();
```

The presets are `Preview` (96 DPI), `Screen` (192 DPI), and `Print` (300 DPI). `ExportImages` returns encoded bytes, pixel dimensions, density, and diagnostics; its consumer overload streams results under shared batch limits. Supply `ProjectImageExportOptions` to select fonts, explicit density, pixel limits, or a rendering deadline. Oversized Project images fail by default; choose `RasterOverflowBehavior.ReduceScale` only when reduced detail is acceptable. Use SVG for zoomable vector text and geometry. Enlarging an existing PNG beyond its pixel dimensions still magnifies its pixels.

For consistent typography, register both regular and bold TrueType faces in the rendering profile used for measurement and output. A regular face alone requires synthesized bold text; font substitution can change line wrapping. `ProjectOfficeReportOptions` selects chart images, data tables, and chart image quality for Word and PowerPoint. Set `IncludeCharts = false` for editable-table reports. A Table view always retains its primary editable content, including when only charts are selected.

Word tables repeat their headers and flow across pages. PowerPoint uses measured row heights to keep complete rows on each slide. Excel retains numeric and date cells, freezes the header row, and prints narrow reports in portrait and wider usage tables in landscape, with a report title and page numbers. Tables wider than eight columns print at full scale across pages; usage sheets repeat UID and name columns on horizontal continuations. Native Office exports retain fixed page dimensions; portable drawing exports can trim unused page height through `ProjectViewOptions.FitPageHeightToContent`.

`ProjectDataWorkflow` transports the Project owner's mapped tables:

```csharp
var projection = project.ExportTables(allowLossyProjection: true);
foreach (string notice in projection.Notices)
    Console.WriteLine(notice);
using var transfer = ProjectDataWorkflow.CreateExcel(projection);
transfer.Save("project-data.xlsx");
foreach (var table in projection.Tables)
    ProjectDataWorkflow.CreateCsv(table.Table).Save(table.Kind + ".csv");
```

Loss permission is explicit because tables omit dependencies, native presentation, and other semantics outside the selected exchange fields. `ReadExcel` requires a bounded worksheet rectangle; formula/error cells and numbers that cannot be represented exactly are rejected. `ReadCsv` uses the CSV owner's parsing and quoting rules. Wrap the resulting `ProjectDataTable` in a `ProjectMappedTable` with explicit field mappings before calling `ProjectDocument.ImportTables`. Table import creates a new project and validates identity, references, units, and conflict policy. See [Project support](../OfficeIMO.Project/SUPPORT.md#portable-reports-and-mapped-data-exchange) for the full boundary.

## Review OCR before publication

`MakePdfSearchableAsync` captures recognition evidence before writing an output. Set `PdfSearchableWorkflowRequest.ReviewAsync` to choose eligible words, or `ReviewCorrectionsAsync` to return original eligible word instances mapped to their reviewed text. Choose one callback. An empty review selection deliberately preserves the source copy; a recognition result with no eligible words and no native source text fails before publication.

```csharp
request.ReviewCorrectionsAsync = (review, token) => {
    token.ThrowIfCancellationRequested();
    IReadOnlyDictionary<OfficeIMO.Pdf.Ocr.PdfRecognizedWord, string> reviewed = review.Ocr.Pages
        .SelectMany(page => page.Words).ToDictionary(word => word, word => word.Text);
    return Task.FromResult(reviewed);
};
```

A host review interface can edit dictionary values and exclude entries before returning them. Correction eligibility, text limits, source-identity checks, output conflicts, and publication guards apply to local files, provider outputs, and OCR sessions. Workflow diagnostics retain recognition warnings and page numbers. A nonrecoverable provider error prevents publication, and image recognition with no usable text does not create an empty success artifact. Successful publication means the reviewed artifact was saved; it does not certify recognition accuracy.

OfficeIMO Studio shows the source region, original recognition, editable replacement, confidence, and inclusion choice. Corrections persist across page navigation. **Next uncertain word** navigates low-confidence and sub-90% words without making rejected words eligible. The selected text can be extracted without creating a PDF or saved as a searchable layer; cancellation preserves the existing destination.

## Reference from source

When working from an OfficeIMO source checkout, reference the workflow project directly:

```xml
<ProjectReference Include="..\OfficeIMO.Workflows\OfficeIMO.Workflows.csproj" />
```

## Inspect concealed text in raster images

`OfficeRasterContentSafety` combines a caller-supplied `IOcrEngine` with decoded pixel evidence. It accepts one static raster image, normalizes it to a metadata-free PNG, and assesses every OCR line, word, or character span that includes bounded pixel or normalized geometry. Image metadata and unbounded provider text do not become visibility findings.

```csharp
using OfficeIMO.ContentSafety;
using OfficeIMO.Ocr;
using OfficeIMO.Workflows;

IOcrEngine engine = GetConfiguredOcrEngine();
byte[] source = File.ReadAllBytes("review.png");

var options = new OfficeRasterContentSafetyOptions {
    EnableOpaqueRectangleRedaction = true,
    MinimumOcrConfidenceForRedaction = 0.9
};

OfficeContentSafetyReport report = await OfficeRasterContentSafety.InspectAsync(
    source,
    engine,
    options,
    cancellationToken);

string[] approvedIds = report.Findings
    .Where(finding => finding.CleanupCapability == OfficeContentCleanupCapability.RedactRegion)
    .Select(finding => finding.Id)
    .ToArray();

OfficeContentCleanupResult result = await OfficeRasterContentSafety.RedactSelectedContentAsync(
    source,
    engine,
    new OfficeContentCleanupSelection(approvedIds),
    options,
    cancellationToken);
File.WriteAllBytes("review-redacted.png", result.Output);
```

The inspector reports nearly transparent, tiny, and low-contrast OCR regions using bounded geometry and conservative pixel evidence. Aggregate OCR text must be fully represented by accepted bounded spans after verified hierarchical duplicates are removed, and malformed Unicode is rejected before finding identities are derived. Redaction is deliberately disabled by default. When enabled, only sufficiently confident, currently matching findings can be selected; the workflow covers their bounded regions with the configured opaque color, emits a single-frame PNG derivative, reopens it, verifies every output pixel, and reruns OCR inspection under the same captured engine identity and capabilities. A selected region, including its configured padding, must not overlap any independently bounded recognized span that was not also selected. When a provider emits an aggregate line or word together with finer spans carrying the same line identity, the finer spans own overlap validation only when they fully reproduce the aggregate text. Cumulative limits bound both region-pixel work and OCR-region intersection comparisons. This is destructive rectangular coverage, not semantic image editing. Multi-frame images, unsupported geometry, oversized provider output, input or analysis work, provider errors or non-recoverable diagnostics, color-rendering metadata that is neither canonical sRGB nor normalized by the managed decoder, unsupported embedded orientation, and outputs where OCR still recognizes leaf text in a changed region all fail closed.

## Save a prepared scan copy

`ScanCleanup` creates a separate PDF containing the prepared page pixels. It uses the same page selection, region crop, perspective correction, and tonal settings as `OfficeIMO.Pdf.Ocr` preview. The source is protected from replacement. Native text, forms, links, signatures, and attachments are omitted from the raster copy, so callers must acknowledge that output contract.

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Workflows;

OfficeWorkflowResult result = await new OfficeWorkflowRunner().RunAsync(new() {
    Operation = OfficeWorkflowOperation.ScanCleanup,
    InputPath = "scan.pdf",
    OutputPath = "prepared-scan.pdf",
    ScanCleanup = new() {
        AcknowledgeRasterOutput = true,
        Preparation = new() {
            Dpi = 200,
            ReadOptions = new() { PageSelection = PdfPageSelection.From(1) },
            ScanProcessing = new() { Deskew = false, StraightenDegrees = 2, Gamma = 1.1 }
        }
    }
});
```

Set `ExpectedSourceSha256` to the SHA-256 hex digest of a reviewed snapshot to reject a source that changed before export. Provider inputs and destinations use the same snapshot, confirmation, recovery, and publication guards as other workflows. To retain the visible source and add searchable text, use the searchable OCR workflow instead.

## Optimize embedded Word images

```csharp
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Workflows;

var runner = new OfficeWorkflowRunner();
OfficeWorkflowResult result = await runner.RunAsync(new OfficeWorkflowRequest {
    Operation = OfficeWorkflowOperation.OptimizeWordImages,
    InputPath = "input.docx",
    OutputPath = "optimized.docx", // .doc or .pdf also selects that output format
    ConflictPolicy = OfficeWorkflowConflictPolicy.Fail,
    WordImageOptimization = new WordImageOptimizationOptions {
        Mode = OfficeImageOptimizationMode.DownsampleAndRecompress,
        TargetDpi = 144,
        JpegQuality = 85
    }
});
```

The runner snapshots the source, optimizes through `OfficeIMO.Word`, reopens the staged output through its format owner, and publishes a separate copy atomically. Source replacement and publication over any batch source are refused. `AnalyzeWordImages` returns per-media diagnostics without publishing a file. Diagnostics include dimensions, formats, candidate metadata removal and whether each change was applied. Metadata removal produces a warning, and the inventory exposes `requiredStagedBytes` for budgeting. `RunBatchAsync` accepts up to 250 requests, snapshots options before execution, and publishes each item independently.

DOCX and supported legacy DOC inputs can produce DOCX, native DOC, or PDF. Incomplete legacy projections block output; analysis warns that its inventory covers only projected pictures. The native DOC writer preflights destination support. Word reports encoded-media savings; `InputBytes` and `OutputBytes` measure actual files. PDF generation after Word optimization retains its default image policy to avoid a second JPEG quality reduction. [Word image options and preservation rules](../OfficeIMO.Word/README.md#images) apply to every host.

## Convert a document

The executable conversion routes run in process through the OfficeIMO format and
rendering packages. They do not launch an external office suite or document
converter. Independent producer files and compatibility checks belong to
validation; they are not prerequisites for running these conversions.


```csharp
using OfficeIMO.Workflows;

var runner = new OfficeWorkflowRunner();
OfficeWorkflowResult result = await runner.RunAsync(new OfficeWorkflowRequest {
    Operation = OfficeWorkflowOperation.Convert,
    ConversionRouteId = "docx-pdf",
    InputPath = "report.docx",
    OutputPath = "report.pdf",
    ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
});

if (!result.Succeeded) {
    throw new InvalidOperationException(result.Summary);
}

Console.WriteLine(result.OutputPath);
```

The same request has a fluent form. When a source/target extension pair maps to
one locally executable route, the builder infers it; use `Via(routeId)` for an
ambiguous pair or to pin a specific route contract.

```csharp
OfficeWorkflowResult result = await OfficeWorkflow
    .Convert("report.docx")
    .To("report.pdf")
    .WithProfile(OfficeWorkflowOutputProfile.PrintReady)
    .OnConflict(OfficeWorkflowConflictPolicy.Replace)
    .RunAsync(cancellationToken: cancellationToken);
```

`OfficeWorkflowCatalog.Routes` projects the complete canonical first-party
conversion catalog for discovery. `ExecutableRoutes` is the subset this local
runner can invoke. Route metadata includes accepted extensions, the owning
package, representative API and result contract, fidelity and support evidence,
known limits, browser and agent availability, and `CanExecute`.

Use `ConversionOptions` (or the fluent `WithConversionOptions` method) for settings
specific to a route. PDF input routes accept `PageRanges`. PDF-to-Word and
PDF-to-PowerPoint expose editable or visual import modes; visual pages accept
`RasterDpi`. Excel-to-PDF supports worksheet-canvas or flowing-table layout, and
PDF-to-HTML supports semantic or positioned HTML. PDF output routes can request
verified lossless compression with `CompressPdfOutput`. The runner rejects options
and output profiles unsupported by the selected route.

```csharp
OfficeWorkflowResult result = await OfficeWorkflow.Convert("source.pdf")
    .To("visual-pages.docx")
    .WithConversionOptions(new OfficeWorkflowConversionOptions {
        PageRanges = "1-3,5",
        WordMode = OfficeIMO.Word.Pdf.PdfWordImportMode.VisualPages,
        RasterDpi = 144
    })
    .RunAsync(cancellationToken: cancellationToken);
```

`OfficeWorkflowRunner.PreviewDocument(bytes, extension, cancellationToken)` creates
an in-memory sample of the first three pages of a PDF, DOCX, XLSX, PPTX, or HTML
artifact. The sample includes rendering diagnostics and uses bounded input and
image sizes. HTML previews do not load external or sibling resources. Previewing
is a review aid; inspect the full saved document for whole-document fidelity.

`RunAsync` also exposes PDF inspection, comparison, optimization, repair planning, repair, and sanitization through typed operations. `ExportPdfPagesAsync` exports selected PDF pages as images, `AssemblePdfAsync` combines supported PDFs, images, documents, folders, and ZIP archives, and `PdfPrintPlanner.Create` produces deterministic print-sheet placement plans.

Every workflow request runs with explicit input and output limits, cancellation, staged output validation, and a caller-selected collision policy. Passwords remain request-only values and are not copied into diagnostics or results. PDF comparison accepts a separate `ComparisonPdfPassword` when the two inputs use different credentials.

`PdfPrintRenderer.Prepare(document, request)` turns an authenticated `PdfDocument` snapshot and its `PdfPrintPlanRequest` into immutable PNG sheets. Display `prepared.Sheets[i].GetPng()` for review, then pass that same `PdfPreparedPrintDocument` to `IPdfPrinterService.SubmitAsync` with `PdfPrintDeliveryOptions`. Delivery does not reopen the source. `PdfPrinterService.GetPrintersAsync` lists installed queues; Windows uses GDI and macOS/Linux use the installed CUPS `lpstat`, `lpoptions`, and `lp` tools. Copies are collated, and duplex defaults to the printer's setting.

`GetPaperSourcesAsync(printerName)` lists the queue's reported paper sources. Set `PdfPrintDeliveryOptions.PaperSourceId` to one of those identifiers, or leave it null to retain the printer default. Identifiers belong to the queried queue; delivery rechecks an explicit selection before submitting. Windows reads driver bins and checks that the driver accepts the selected bin. CUPS discovers `InputSlot` or `media-source` choices exposed by `lpoptions -l`. Queues that expose no choices retain their default source.

Print preparation enforces PDF printing permissions, page and raster limits, cancellation, and a retained-output byte budget. Rendering above 150 DPI also requires high-quality print permission when restrictions apply. The output reflects the managed renderer's diagnostics and printer margins. `PdfPrintSubmission` is a queue acceptance receipt, not confirmation of physical delivery; `PdfPrintDeliveryException` means submission began and retrying could duplicate pages. Windows file-printer paths must be new local paths and are checked before submission; the driver controls the eventual write.

Applications that keep documents open can set `PublicationGuard` on `OfficeWorkflowRequest`, `PdfAssemblyRequest`, and `PdfPageImageExportRequest`. Implement `IOfficeWorkflowPublicationGuard.CanPublishAsync` to check live ownership of the supplied absolute destination. For directory outputs, check whether publication would replace a directory containing an owned document. The runner calls the guard after validating the staged artifact and checks every numbered candidate: a denied destination fails `Fail` or `Replace`, while `Rename` tries the next name. Cancellation and guard errors prevent publication. Calls can originate on worker threads, so UI hosts must dispatch ownership inspection to their UI thread. This is an application ownership check at publication time; it does not lock paths against concurrent external filesystem changes.

## Save protected or unencrypted PDF copies

```csharp
OfficeWorkflowResult protectedCopy = await OfficeWorkflow.ProtectPdf("report.pdf",
    new OfficeIMO.Pdf.PdfStandardEncryptionOptions(documentPassword) {
        OwnerPassword = ownerPassword,
        AllowedPermissions = OfficeIMO.Pdf.PdfStandardPermissions.Print
    })
    .To("protected.pdf")
    .RunAsync(cancellationToken: cancellationToken);

OfficeWorkflowResult unencryptedCopy = await OfficeWorkflow
    .RemovePdfProtection("protected.pdf", ownerPassword)
    .To("unencrypted.pdf")
    .RunAsync(cancellationToken: cancellationToken);
```

Typed requests use `ProtectPdf` with `OutputEncryption`, or `RemovePdfProtection`, and `PdfOwnerPassword` for existing protection. The owner password takes precedence over `PdfPassword` for these operations. To replace existing protection, use `ProtectPdf(...).WithPdfOwnerPassword(currentOwnerPassword)`. Settings are copied before execution; passwords are not included in results or reports.

These operations create separate copies, preserve their sources, and use the PDF engine's authorization and rewrite-preservation policy. Existing signatures or other protected document structures may prevent a rewrite. AES-256 is the default; AES-128 and explicitly selected legacy RC4 follow the canonical encryption options. Output verification checks the document-open password, page count, encryption state, permissions, metadata protection, and preservation report before publication. A null PDF metadata-encryption flag means the standard default of encrypted metadata.

Only the `Faithful` profile applies. Input snapshots, byte limits during generation, output conflicts, provider confirmation contracts, recovery, and final publication guards follow the same runner behavior as other single-output operations. Cancellation is forwarded through security preflight, graph rewriting, preservation inspection, and publication.

## Extract selected PDF pages

```csharp
OfficeWorkflowResult result = await OfficeWorkflow.ExtractPages("report.pdf", 5, 1, 2, 5)
    .To("selected-pages.pdf")
    .OnConflict(OfficeWorkflowConflictPolicy.Rename)
    .RunAsync(cancellationToken: cancellationToken);
```

The equivalent typed request uses `Operation = OfficeWorkflowOperation.ExtractPages` and `PageNumbers = [5, 1, 2, 5]`. Page numbers are one-based; order and intentional repeats are preserved, up to 100,000 selected pages. Extraction uses the PDF engine's page-preservation policy and supports only the `Faithful` profile. It creates a separate PDF and does not permit replacing the source.

The runner snapshots local and provider inputs, checks for source changes before publication, bounds output serialization, and reopens the generated PDF before publishing. `InputStream`, `OutputStream`, `PublicationGuard`, and the result's publication and recovery states follow the same contracts as other single-output workflows. Cancellation is observed before and after synchronous page extraction and during serialization; it cannot interrupt the PDF engine while that synchronous step is running.

## Save a certificate-signed PDF copy

Supply a caller-owned `IPdfExternalSigner` and `IPdfSignatureCryptographyProvider`. The PDF engine owns signature creation and inspection; the workflow captures the settings and publishes a separate output only after checking page count, signature structure, signature math, and document digests.

```csharp
OfficeWorkflowResult signedCopy = await OfficeWorkflow.SignPdf("report.pdf", signer,
    new OfficeIMO.Pdf.PdfExternalSignatureOptions {
        FieldName = "Approval",
        Reason = "Reviewed",
        VisibleAppearance = new() { PageNumber = 1, X = 36, Y = 36, Width = 180, Height = 48 }
    }, verifier)
    .To("signed-report.pdf")
    .RunAsync(cancellationToken: cancellationToken);
```

Keep the signer and verifier alive until the task completes. A host using `OfficeIMO.Security` can provide `PdfCmsExternalSigner` and `PdfCmsSignatureCryptographyProvider` adapters. The engine enforces the source document's permissions and certification policy; its current signing plan rejects documents that already contain a signature. Rejected requests leave the source and existing output unchanged.

`SignatureReport` reports certificate-chain, revocation, and timestamp evidence separately. Successful publication does not by itself establish certificate trust. A visible appearance identifies the signature on the page; it is not a substitute for cryptographic verification. On unconfirmed provider publication, the report describes the retained prepared artifact, not the destination's contents. Signing callbacks and cryptographic validation may be synchronous; cancellation is observed around those calls and prevents later publication.

## Split a PDF into consecutive parts

```csharp
PdfSplitWorkflowResult result = await runner.SplitPdfAsync(new PdfSplitWorkflowRequest {
    InputPath = "report.pdf",
    OutputDirectory = "report-parts",
    PagesPerDocument = 10,
    ConflictPolicy = OfficeWorkflowConflictPolicy.Rename
}, cancellationToken: cancellationToken);

foreach (PdfSplitFile file in result.Files) {
    Console.WriteLine($"{file.Path}: {file.PageCount} pages starting at source page {file.FirstSourcePage}");
}
```

The runner produces `part-001.pdf`, `part-002.pdf`, and subsequent parts in source order. Hosts can call `PdfSplitPlan.Create(pageCount, pagesPerDocument)` to preview the same filenames and ranges that execution uses. To split at chosen pages instead, such as top-level bookmarks, build `PdfSplitPlan.FromStarts(pageCount, starts)` from `PdfSplitStart(firstPage, title)` values and assign it to `PdfSplitWorkflowRequest.Plan`; parts are named `001-Title.pdf`, pages before the first start form a leading `part-001.pdf`, and the runner validates that the parts cover every source page exactly once in order and have unique file names. The runner generates and reopens one part at a time, checks the aggregate output budget before continuing, and publishes a local folder as a unit. `MaximumParts` limits the output count. Cancellation is checked between parts and during file operations; the PDF engine's synchronous generation of one part must finish before cancellation can stop it.

For provider folders, supply `DirectoryOutput` and explicitly choose `Replace`. Each part is written and verified individually. Inspect `Status`, `Files`, and `OutputRecoveries`: verified parts remain available if a later write fails. Local directory recovery locations appear in diagnostic details when an interrupted replacement needs attention. `InputStream` and `PublicationGuard` use the same source verification and live ownership contracts as other workflows.

## Read provider-backed inputs

Set `InputStream` on an `OfficeWorkflowRequest` when a file picker or storage provider supplies stream access. Keep `InputPath` as the original location or absolute URI, and supply the display filename for format routing:

```csharp
var request = new OfficeWorkflowRequest {
    Operation = OfficeWorkflowOperation.Convert,
    ConversionRouteId = "docx-pdf",
    InputPath = selectedLocation,
    InputStream = new OfficeWorkflowStreamInput(selectedName, openSelectedReadStream),
    OutputPath = outputPdfPath
};
OfficeWorkflowResult result = await runner.RunAsync(request, cancellationToken: cancellationToken);
```

`openSelectedReadStream` is a `Func<CancellationToken, Task<Stream>>`. It must return a fresh readable stream with the provider's permission scope each time. The runner closes every returned stream, stages a bounded private input for the document engine, and verifies the provider's SHA-256 again after host authorization and before publication. An optional `expectedSha256` constructor argument binds execution to contents captured when the user selected the input. Revoked access, changed contents, cancellation, and exceeded limits prevent publication. This is a point-in-time content check; providers do not offer a shared filesystem lock or atomic compare-and-replace contract.

For provider selections with a local path, the runner captures file identity while the read stream's access scope is active. It reopens that access for source/output checks and host authorization, then rejects physical source replacement even when the new file has identical contents. Local provider destinations also receive a read-scope check before writing; a new file may report `FileNotFoundException`, while other access failures prevent publication.

Comparison accepts `ComparisonStream`. Assembly accepts `SourceStreams`, keyed by the exact original entries in `Sources`, and preserves input order and display names. Its provider staging shares the total input byte budget. A provider HTML stream can use embedded resources; selecting it alone does not grant access to neighboring images or stylesheets. A selected ZIP can carry relative resources through the existing bounded archive intake.

For a selected provider folder, set `PdfAssemblyRequest.SourceDirectories` with an `OfficeWorkflowDirectoryInput` keyed by its original `Sources` entry. Its enumeration factory returns `OfficeWorkflowDirectoryEntry` values with a relative path, original location, and reopenable file input; a null input denotes a directory. Enumerate parents before children, obey the supplied recursion and traversal limits, and never follow links. Keep item references available until the runner returns. The runner preserves the relative tree for HTML resources, enforces aggregate entry and byte limits, rejects unsafe or colliding portable names, and rechecks both membership and file contents before publication. Each re-enumeration must reflect current provider state, including newly returned file objects at an existing location.

Page-image export accepts `PdfPageImageExportRequest.InputStream` and applies the same bounded staging and provider-content check before publishing its filesystem output folder. For print preview, use `PdfPrintPlanner.Create(document, request)` with an already opened `PdfDocument` to plan and render from the same snapshot. That overload uses the document's existing authentication and printing permissions; it does not reopen the request's input location.

Provider operations require an explicit output destination when they produce a file. Input staging is removed before publication or on failure; cleanup failures are reported. Report-only inspection and comparison may omit a destination.

## Write provider folders

Set `DirectoryOutput` on `PdfPageImageExportRequest` to write into a selected provider folder. Its `OfficeWorkflowDirectoryOutput` resolver receives each image filename and returns an `OfficeWorkflowDirectoryOutputFile` without modifying the provider. Existing children use their actual location. For new children whose location is assigned during creation, use the selected parent as the initial location and supply `OfficeWorkflowStreamOutput.PrepareDestination`; this callback creates the child after recovery is durable and returns its actual location. The runner authorizes that location before opening the write stream. Read factories must reopen the current child, not a cached copy of its bytes.

Provider folders require `Replace` and explicit consent. Each file is verified separately; cancellation or failure stops further writes without rolling back earlier ones. `Files` and `OutputBytes` describe verified outputs even when the batch does not complete. `OutputRecoveries` contains every retained recovery copy. Use these fields with `Status` rather than treating the folder as an atomic result. A recovery record for a newly created child identifies the selected parent and requested filename; successfully published files report their actual provider locations.

## Write provider-backed outputs

Set `OutputStream` on an `OfficeWorkflowRequest` or `PdfAssemblyRequest` to publish through a selected provider. Supply an `OfficeWorkflowStreamOutput` with the display filename, fresh read and write stream factories, and an `OfficeWorkflowOutputRecoveryStore` rooted in a private local directory. Set `OutputPath` to the original provider reference and `ConflictPolicy` to `Replace`. The host must obtain explicit consent for a direct write and for the required local recovery copy.

The runner validates the complete artifact and retains a verified local copy before preparing a new provider child or opening the write stream. It closes the write stream and reads the destination back to verify its SHA-256. A verified write returns `Completed` with the provider reference in `OutputPath`. A failure after the write starts returns `Unconfirmed`, leaves `OutputPath` unset, and exposes the retained copy through `Recovery`. Cancellation after the write starts also returns `Unconfirmed`; it does not prove that the destination is unchanged. Do not automatically retry these results.

The store defaults to a 1 GiB aggregate admission limit and at most 100 records. `GetRecoveries()` restores available records after restart, `VerifyAsync()` checks a copy before use, and `Discard()` removes a copy after explicit user action. Active publications are excluded from discovery. Successful or safely rejected writes remove their copies; cleanup failures are reported and can leave a recovery record. Before admitting another output, the store removes recognized incomplete records left before metadata publication, while preserving active leases and unfamiliar contents. Retained recovery copies do not expire automatically. Keep them outside normal output locations and require Save As when opening them for editing. Provider writes cannot guarantee atomic replacement, rollback, or exclusion of concurrent writers.

## Review and apply PDF redactions

Searchable PDF generation uses `OfficeWorkflowRunner.MakePdfSearchableAsync` with a `PdfSearchableWorkflowRequest` and a caller-owned `IOcrEngine`:

```csharp
var result = await new OfficeWorkflowRunner().MakePdfSearchableAsync(new() {
    InputPath = "scan.pdf",
    OutputPath = "searchable.pdf",
    ConflictPolicy = OfficeWorkflowConflictPolicy.Fail,
    Ocr = new OfficeIMO.Pdf.Ocr.PdfOcrMergeOptions { Language = "en", Dpi = 150 }
}, engine, cancellationToken);
```

The runner captures a bounded input snapshot, adds searchable text through `OfficeIMO.Pdf.Ocr`, and reopens the staged PDF before publication. It verifies source contents and local physical identity after recognition, then applies `PublicationGuard` and the selected conflict policy. The request also accepts `InputStream` and `OutputStream` with the same provider consent and recovery requirements described above. Inspect `Status`, `OutputPath`, and `Recovery` before opening or retrying an output. The engine remains owned by the caller.

Set `ReviewAsync` to pause before creating the text layer. The callback receives a `PdfSearchableOcrReview` and returns eligible word instances selected from that review. The shared PDF owner rejects foreign, duplicate, and policy-rejected selections. The destination remains untouched while review is pending, cancellation prevents publication, and source identity is checked again after the decision. Without a callback, the runner uses all eligible words.

Use `RecognizeImageAsync` for standalone images. It reads the image through `OfficeIMO.Reader.Image`, executes the selected engine through `OfficeIMO.Reader.Ocr`, and saves UTF-8 text. Its optional review callback receives the original image, recognition evidence, and recognized text, and returns the text to save:

```csharp
var result = await runner.RecognizeImageAsync(new ImageOcrWorkflowRequest {
    InputPath = "invoice.png",
    OutputPath = "invoice.txt",
    Ocr = new OfficeIMO.Reader.OfficeDocumentOcrExecutionOptions { Language = "eng" },
    ReviewAsync = (review, token) => Task.FromResult(review.Text)
}, engine, cancellationToken);
```

The callback can present a preview and accept corrections. Until it returns, the destination is untouched. Empty recognition remains visible in diagnostics; failed or skipped recognition does not publish a partial text file. Corrected text is bounded by `Limits.MaximumOutputBytes`, reopened before publication, and protected by the same source identity, conflict, provider consent, and recovery contracts as PDF output.

`RunOcrSessionAsync` accepts an ordered collection of `OfficeOcrSessionRequest` items containing either request type. It snapshots request settings, uses one caller-owned engine sequentially, and protects every selected source from every output. Each item has a unique caller id and a distinct output destination. Progress and terminal result callbacks let a host show completed outputs while later items await review. Cancellation retains completed outputs and returns cancelled outcomes for unstarted items. A retry should contain only the explicitly selected failed or cancelled items; an `Unconfirmed` result stops the remaining items and requires checking the destination and recovery copy first. Pass previously completed output locations through `protectedOutputPaths` when retrying a subset, including when a provider resolves a different destination during publication. Also pass every retained session source as an `OfficeWorkflowProtectedSource` through `protectedInputs`, including its provider stream access. This preserves original inputs of completed items while retrying or adding work.

Redaction uses a separate versioned plan/review/apply contract. Planning produces privacy-safe candidate identifiers and geometry. Application re-plans the exact source and recipe, requires every current candidate to be explicitly approved or rejected, applies only approved candidates, and publishes only after native and configured OCR verification succeeds.

```csharp
var recipe = new PdfRedactionRecipe();
recipe.Rules.Add(new PdfRedactionRule {
    Name = "account-number",
    Kind = PdfRedactionRuleKind.Literal,
    Value = "Account: 123-45-6789",
    ContentScope = PdfRedactionContentScope.TextAndUnderlay,
    AppearanceMode = PdfRedactionAppearanceMode.QuantizedWidth
});

var runner = new OfficeWorkflowRunner();
PdfRedactionWorkflowResult plan = await runner.RunRedactionAsync(
    new PdfRedactionWorkflowRequest {
        Mode = PdfRedactionWorkflowMode.PlanOnly,
        InputPath = "contract.pdf",
        EvidencePath = "contract.plan.json",
        Recipe = recipe
    });

var decisions = new PdfRedactionDecisionManifest {
    SourceSha256 = plan.SourceSha256,
    RecipeSha256 = plan.RecipeSha256,
    ApprovedCandidateIds = plan.Candidates.Select(candidate => candidate.Id).ToList()
};

PdfRedactionWorkflowResult applied = await runner.RunRedactionAsync(
    new PdfRedactionWorkflowRequest {
        Mode = PdfRedactionWorkflowMode.ApplyAndVerify,
        InputPath = "contract.pdf",
        OutputPath = "contract-redacted.pdf",
        EvidencePath = "contract-redacted.evidence.json",
        Recipe = recipe,
        Decisions = decisions
    });
```

Rule and explicit-region names are stable, non-sensitive evidence identifiers. `ContentScope` decides whether a reviewed area removes only text or also intersecting underlay content. `AppearanceMode` independently controls the privacy of the visible mark: exact, nearby-merged, quantized-width, or full-line. Recipe, decision, and batch JSON reject unknown members so misspelled policy fields cannot silently fall back to defaults.

The schemas are `officeimo.pdf.redaction.recipe.v1`, `officeimo.pdf.redaction.plan.v1`, `officeimo.pdf.redaction.decisions.v1`, `officeimo.pdf.redaction.result.v1`, `officeimo.pdf.redaction.batch-request.v1`, and `officeimo.pdf.redaction.batch.v1`. Persisted `PdfRedactionWorkflowRecord` JSON omits matched text, extracted text, passwords, OCR payloads, provider options, host paths, and caller request identifiers. The in-memory operational result still carries paths and request correlation for host UX. Evidence retains rule names, policies, hashes, counts, complete atomic candidate geometry, stable issue codes, one-way SHA-256 fingerprints of provider/model/language values, and OCR confidence. Raw provider-returned metadata is never persisted, so document text or credentials cannot become evidence metadata even when they contain only identifier characters. Encrypted input requires an explicit reject, decrypt, or decrypt-and-reencrypt policy with runtime-only owner credentials. Zero-area verification of a re-encrypted output also requires the trusted output SHA-256 from prior apply evidence.

Signed input uses an explicit `SignaturePolicy`. The default rejects it. `CreateUnsignedDerivative` removes invalidated signature structures through a full rewrite before planning and records source/output signature counts; `CreateAndSignDerivative` additionally requires a runtime `IPdfExternalSigner` and can cryptographically validate the new signature through an optional `IPdfSignatureCryptographyProvider`. The output is always a separate artifact. Runtime `ExternalValidators` accept `IPdfRedactionCancellationAwareExternalValidator` implementations that bind independent parser, renderer, or forensic checks to the final bytes; their names are retained in evidence, cancellation stops the workflow before publication, and any rejection prevents publication.

Single-item evidence, per-output bytes, batch items, concurrency, and aggregate prepared output/evidence bytes have independent limits. Batch preparation reserves each in-flight item's configured worst-case size and fails before publication when the aggregate ceiling cannot be honored; successful items are reclassified as unpublished if any sibling fails.

The file-set overload deterministically selects PDFs and mirrors their relative directories into separate output, evidence, and decision roots:

```csharp
PdfRedactionBatchResult batch = await runner.RunRedactionBatchAsync(
    new PdfRedactionBatchRequest {
        Mode = PdfRedactionWorkflowMode.PlanOnly,
        InputRoot = "incoming",
        EvidenceRoot = "review-evidence",
        ManifestPath = "review-evidence/batch.json",
        Recipe = recipe,
        PublicationPolicy = PdfRedactionBatchPublicationPolicy.AtomicAll
    });
```

`RunRedactionBatchAsync` prepares every bounded item before atomic publication with configurable concurrency, stages every file beside its destination, and rolls back already published files if an ordinary publication failure occurs. `ContinuePerItem` instead publishes successful items independently and records failures in the consolidated manifest. Batch destinations must be portable-case unique, remain physically outside the input root, and use one fail-or-replace conflict policy. Recursive discovery does not follow reparse points, and explicit linked inputs must still resolve beneath the physical input root. This is an in-process publication transaction, not a filesystem-wide crash transaction.

## Inspect and remove provenance

The provenance workflow keeps format logic in its owning package. `OfficeIMO.Word`, `OfficeIMO.Excel`, `OfficeIMO.PowerPoint`, `OfficeIMO.Visio`, `OfficeIMO.OpenDocument`, `OfficeIMO.Epub`, `OfficeIMO.Pdf`, `OfficeIMO.Html`, and `OfficeIMO.Markdown` handle their formats; `OfficeIMO.Core` handles supported images and structured text. Consumers can discover the exact extension, structural format, owner, operation, memory-only, and browser contract through `OfficeProvenanceWorkflowCatalog.All`, `ToJson()`, or `ToMarkdown()`.

The workflow requires a registered extension and matching structural format. It does not infer ownership for unknown extensions or generic containers. Applications that already own such a format context can call the lower-level `OfficeProvenanceInspector` API directly for signature-based inspection.

```csharp
using OfficeIMO.Workflows;

var runner = new OfficeWorkflowRunner();

OfficeProvenanceWorkflowResult inspection = await runner.RunProvenanceAsync(
    new OfficeProvenanceWorkflowRequest {
        Operation = OfficeProvenanceWorkflowOperation.Inspect,
        InputPath = "report.docx"
    });

OfficeProvenanceWorkflowResult removal = await runner.RunProvenanceAsync(
    new OfficeProvenanceWorkflowRequest {
        Operation = OfficeProvenanceWorkflowOperation.Remove,
        InputPath = "report.docx",
        OutputPath = "report.cleaned.docx",
        ExpectedInputSha256 = inspection.InputSha256,
        ConflictPolicy = OfficeWorkflowConflictPolicy.Fail
    });
```

`Assess` combines the owner-specific structural report with exact Unicode findings and optional `IOfficeProvenanceVerifier` / `IOfficeProvenanceSignalDetector` services supplied to the runner. It preserves each provider's result and does not infer a universal authorship verdict.

Removal is strict by default. It removes only selected, structurally valid carriers and blocks a package-signature-invalidating save unless the caller explicitly selects `OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures`. The output is written to a sibling staging file, reopened through the same format owner, checked against the removal report, and only then published under the requested conflict policy. Generic ZIP packages and renamed package subtypes are rejected because the workflow has no matching registered format owner for them.

`OfficeProvenanceReportSerializer.Serialize(result)` produces the same `officeimo.provenance.result.v2` document used by the CLI, Studio report export, and browser provenance download. `SerializeBatch(results)` uses `officeimo.provenance.batch.v2`. Reports retain structured evidence and diagnostics, string enum values, coverage notes, input/output SHA-256 hashes, and explicit check states. Assessment reports distinguish disabled or unsupported Unicode inspection from a completed empty report and distinguish an absent provider from verification that ran.

Pass `ExpectedInputSha256` from a reviewed result when a later action must use the same source bytes. The runner compares the immutable input snapshot before any mutation. `PublicationGuard` applies the host's live ownership check to each final destination; Fail/Replace reject an owned path and Rename skips it.

`OfficeProvenanceAudit.RunAsync(new OfficeProvenanceAuditRequest { Inputs = ["documents"], Include = ["*.html"], MaximumItems = 1000 })` discovers and assesses a bounded set without modifying it. Discovery is recursive by default, excludes symbolic links and common generated/VCS directories, and fails on an empty selection or exceeded bounds. Explicit files retain ordinary workflow errors. `OfficeProvenanceAudit.HasFindings(result, carriers: false, dangerousText: true)` evaluates an evidence policy; callers must handle execution failures separately. `OfficeProvenanceSarif.Serialize(results)` exports the same evidence and failures as SARIF 2.1.0.

Use `RunProvenanceBatchAsync` for bounded sequential batches. Sequential execution keeps parser and provider resource use predictable, while per-request progress includes an overall batch fraction.


For a memory-only host, `OfficeProvenanceBufferWorkflow.Inspect(bytes, fileName, options)` and `Remove(bytes, fileName, removalOptions)` use the same catalog and format owners without opening paths or following remote references. Read the qualified extensions from `OfficeProvenanceWorkflowCatalog.MemoryOnlyExtensions`; the current families are JPEG, PNG, WebP, PDF, DOCX, XLSX, and PPTX. Removal returns a separate result and re-inspects its bytes before returning. Specify limits appropriate to the host; a browser should use tighter limits than a local batch runner.

```csharp
var inspection = OfficeProvenanceBufferWorkflow.Inspect(inputBytes, "report.docx");
var result = OfficeProvenanceBufferWorkflow.Remove(inputBytes, "report.docx");
byte[] cleanedCopy = result.ToArray();
// Inspect result.After and result.Changes before presenting the copy to the user.
```

For memory-only report export, pass the inspected bytes and report to `OfficeProvenanceReportSerializer.FromBuffer(fileName, bytes, inspection, removal)` and serialize the returned result. These factories do not read paths or verify cryptographic authenticity.

`OfficeTextIntegrityReview` in Core owns source-bound text selections and encoding-preserving export. `OfficeTextIntegrityReportSerializer.Serialize(review, review.Text, fileName, selectedIndices)` exports exact findings, selected occurrence indices, UTF-16 offset units, source hashes, encoding/BOM information, and the selected-copy digest. It does not include the full source text.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 1 | 0 | 0 | 0 | 0 | 0 |
| Read | 1 | 0 | 0 | 0 | 0 | 0 |
| Edit | 1 | 0 | 0 | 0 | 0 | 0 |
| Preserve | 1 | 0 | 0 | 0 | 0 | 0 |
| Validate | 0 | 1 | 0 | 0 | 0 | 0 |
| Convert | 0 | 3 | 0 | 0 | 0 | 0 |
| Export | 1 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Workflows` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
