# OfficeIMO.Epub - EPUB reading and authoring

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Epub)](https://www.nuget.org/packages/OfficeIMO.Epub)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Epub?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Epub)

`OfficeIMO.Epub` reads EPUB publications and creates or edits bounded EPUB 2/3 packages.
Use `EpubDocument` for extraction and `EpubPublication` when the complete package must
survive editing and saving.

## Install

```powershell
dotnet add package OfficeIMO.Epub
```

## Quick start

```csharp
using OfficeIMO.Epub;

EpubDocument book = EpubDocument.Load("book.epub", new EpubReadOptions {
    PreferSpineOrder = true,
    IncludeRawHtml = false,
    MaxChapters = 100
});

Console.WriteLine(book.Title);

foreach (EpubChapter chapter in book.Chapters) {
    Console.WriteLine($"{chapter.Order}. {chapter.Title ?? chapter.Path}");
    Console.WriteLine(chapter.Text);
}

foreach (string warning in book.Warnings) {
    Console.WriteLine(warning);
}
```

### Check chapter completeness and limits

`book.ReadSummary` counts selected reading positions, extracted chapters, and skipped
positions. `IsComplete` means every selected spine position was emitted; it does not
certify EPUB conformance, rendering fidelity, or retention of every resource.
Non-linear positions excluded by policy are not counted as requested.

```csharp
EpubDocument book = EpubDocument.Load("book.epub", new EpubReadOptions {
    MaxChapters = 500,
    MaxTotalTextCharacters = 32L * 1024L * 1024L
});

Console.WriteLine($"{book.ReadSummary.ExtractedChapterCount}/{book.ReadSummary.RequestedChapterCount}");
if (!book.ReadSummary.IsComplete) {
    foreach (EpubDiagnostic diagnostic in book.Diagnostics) {
        Console.WriteLine($"{diagnostic.Code}: {diagnostic.Message}");
    }
}
```

Chapter count, chapter byte, and total text limits produce diagnostics when selected
content is omitted. Total text is measured in UTF-16 characters and defaults to
32 Mi characters. Adjust the budget explicitly for larger publications.
Malformed character encodings are diagnosed and skipped rather than replaced silently.
Repeated positions that fail for the same reason produce one identical diagnostic;
the read summary still counts every skipped position.

Spine selection is independent of ordering: repeated references remain distinct
chapters, and `PreferSpineOrder = false` does not bypass non-linear exclusion.
Unsupported spine media types follow their manifest fallback chains to the first
supported XHTML or SVG resource. Chapters identify that extracted resource while
retaining the original spine position and selection policy. Missing or cyclic
fallback chains produce diagnostics; fallback chapters use the normal extraction budgets.
Archive recovery scanning applies only when no usable spine is declared, excludes
declared navigation documents, and sets `UsedFallbackScan`; completeness remains unknown.

SVG spine documents retain their positions, text, and structure. `IncludeRawHtml`
also retains their original SVG markup in `chapter.Html`. SVG extraction does not
reproduce fixed-layout page geometry. Inline text remains continuous across formatting
elements, while block boundaries and explicit whitespace separate text.

The synchronous three-argument `Load` overloads accept a `CancellationToken` for
package reading and parsing. `LoadAsync` accepts the token through its existing overloads.

### Inspect package signatures

```csharp
EpubDocument book = EpubDocument.Load("signed.epub");

if (book.HasSignatures) {
    Console.WriteLine($"Signature elements: {book.Signatures.XmlSignatureCount}");
    Console.WriteLine($"Well-formed signatures.xml: {book.Signatures.IsWellFormed}");
}
```

The parser reads `META-INF/signatures.xml` under the normal bounded package limits and reports malformed signature
metadata. It does not claim cryptographic validation and does not require `OfficeIMO.Security`.

To create or validate the bounded OfficeIMO XML package-manifest signature profile, pass an optional provider explicitly:

```csharp
using OfficeIMO.Security;

IOfficeSecurityProvider security = OfficeSecurityProvider.Default;
EpubDocument.SignPackage("book.epub", security, signingCertificate);
OfficeXmlPackageSignatureValidationReport validation =
    EpubDocument.ValidatePackageSignatures("book.epub", security);
```

The signed manifest covers every non-carrier ZIP entry. Validation rejects missing, changed, duplicate, and unsigned entries. General producer-specific EPUB signatures and DRM/resource decryption remain outside this API. Recognized IDPF and Adobe font obfuscation is handled separately by the reader.

### Inspect bounded manifest resources

```csharp
EpubDocument book = EpubDocument.Load("book.epub", new EpubReadOptions {
    IncludeResourceData = true,
    MaxResources = 500,
    MaxResourceBytes = 4L * 1024L * 1024L,
    MaxTotalResourceBytes = 32L * 1024L * 1024L
});

foreach (EpubResource resource in book.Resources) {
    Console.WriteLine($"{resource.Path} ({resource.MediaType}, {resource.LengthBytes} bytes)");
}
```

Manifest metadata is returned even when payload loading is disabled. Payload inclusion is opt-in and bounded per resource, in total, and by resource count; skipped payloads produce warnings.

When `IncludeResourceData` is enabled, IDPF and Adobe font-obfuscated resources are deobfuscated only when the OPF package identity provides the required key. `EpubResource.WasDeobfuscated` identifies the resulting payload. If the identity is missing or malformed, `Data` remains unavailable and a structured diagnostic is returned; the reader does not expose still-obfuscated bytes as usable font data. This reversible standards-defined obfuscation is not DRM decryption.

### Inspect and remove selected concealed HTML

`InspectContentSafety` evaluates every local manifest HTML/XHTML resource with the shared bounded HTML/CSS safety model. Linked stylesheets and recursive `@import` rules resolve only inside the EPUB package.

```csharp
using OfficeIMO.ContentSafety;
using OfficeIMO.Epub;

OfficeContentSafetyReport report = EpubDocument.InspectContentSafety("book.epub");
OfficeContentSafetyFinding finding = report.Findings.Single(item => item.TextPreview.Contains("ignore previous", StringComparison.OrdinalIgnoreCase));

OfficeContentCleanupResult cleaned = EpubDocument.RemoveSelectedContent(
    "book.epub",
    "book-clean.epub",
    new OfficeContentCleanupSelection(new[] { finding.Id }));
```

Cleanup replaces only changed content documents, preserves unrelated ZIP entries and the required leading uncompressed `mimetype` entry, then reopens and reinspects the output. XHTML inspection retains document-level comments as evidence, requires every local manifest HTML resource, accepts OPF package elements and core attributes only in their defined namespaces, and applies stylesheet links and base URLs only in the XHTML namespace. Rewritten HTML content preserves its declared character encoding. Missing, malformed, duplicate, encrypted, ambiguous, active integrity-qualified, external, or over-budget content and stylesheet dependencies fail closed, as do conflicting preferred titled stylesheet sets and documents with Content Security Policy declarations that the bounded cascade does not model. Foreign-namespace base and stylesheet elements, non-CSS stylesheets, and inactive-media stylesheets remain inert. An empty selection preserves the original bytes.

Package signatures, including ZIP central-directory signature records, block mutation by default. A caller that accepts invalidation must request removal explicitly:

```csharp
var cleanupOptions = new OfficeContentCleanupOptions {
    SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
};
```

That policy removes both supported signature carriers during an authorized rewrite: `META-INF/signatures.xml` and any ZIP central-directory digital-signature record. It does not claim that a modified package remains signed.

### Resolve chapter-relative references

```csharp
EpubReference reference = EpubReference.Resolve(
    "EPUB/text/chapter.xhtml",
    "../images/cover%20art.png?size=large#front");

if (reference.Kind == EpubReferenceKind.Container) {
    Console.WriteLine(reference.ContainerPath);     // decoded ZIP lookup path
    Console.WriteLine(reference.ContainerUrlPath);  // URL-encoded link path
    Console.WriteLine(reference.ResolvedValue);     // encoded path plus query and fragment
}
```

Use the three-argument overload with `chapter.BaseHref` when resolving URLs found in chapter markup. Results distinguish container, external, embedded data, and invalid references without performing network or file-system access. `ContainerPath` remains case-sensitive and decoded for ZIP lookup, while `ContainerUrlPath`, `Target`, and `ResolvedValue` preserve a safe URL serialization. Root-relative references are resolved safely but marked non-conforming; `file:` URLs, ambiguous encoded separators, and paths that escape the container are rejected.

## What it does

- Opens EPUB files as ZIP containers.
- Parses `META-INF/container.xml` and OPF package metadata.
- Follows OPF manifest and spine ordering.
- Reads hierarchical EPUB 3 navigation and EPUB 2 NCX labels when available.
- Extracts chapter text from XHTML and SVG content documents.
- Returns deterministic OPF manifest resources with optional bounded payloads.
- Resolves package, navigation, and content references through a shared typed URL contract.
- Emits structured diagnostics and warning messages for malformed, unsafe, encrypted, fixed-layout, or unreadable content.

## Examples

### Read metadata and spine-ordered chapters

```csharp
using OfficeIMO.Epub;

EpubDocument book = EpubDocument.Load("handbook.epub", new EpubReadOptions {
    PreferSpineOrder = true,
    IncludeNonLinearSpineItems = false,
    MaxChapters = 50
});

Console.WriteLine(book.Title);
Console.WriteLine(book.Creator);
Console.WriteLine(book.Language);

foreach (var chapter in book.Chapters) {
    Console.WriteLine($"{chapter.Order}. {chapter.Title ?? chapter.Path}");
}
```

### Keep raw chapter HTML when building a converter

```csharp
using OfficeIMO.Epub;

var book = EpubDocument.Load("book.epub", new EpubReadOptions {
    IncludeRawHtml = true,
    MaxChapterBytes = 2L * 1024L * 1024L
});

foreach (var chapter in book.Chapters) {
    File.WriteAllText(
        $"chapter-{chapter.Order:000}.txt",
        chapter.Text);

    if (chapter.Html != null) {
        File.WriteAllText($"chapter-{chapter.Order:000}.xhtml", chapter.Html);
    }
}
```

### Read from a stream and report warnings

```csharp
using OfficeIMO.Epub;

await using var stream = File.OpenRead("upload.epub");
EpubDocument book = EpubDocument.Load(stream, new EpubReadOptions {
    FallbackToHtmlScan = true,
    DeterministicOrder = true
});

foreach (string warning in book.Warnings) {
    Console.WriteLine(warning);
}
```

## Create a publication

```csharp
using OfficeIMO.Epub;

EpubPublication publication = EpubPublication.Create("A publishing example", "en");
publication.Creator = "Author";
publication.AddStylesheet("style", "EPUB/style.css",
    "body { font-family: sans-serif; } p { line-height: 1.4; }");
publication.AddChapter("opening", "EPUB/opening.xhtml", "Opening",
    "<h1 id='opening'>Opening</h1><p>A chapter with <em>formatted</em> text.</p>",
    new[] { "style" });
publication.SetNavigation(new[] {
    new EpubNavigationEntry("Opening", "EPUB/opening.xhtml#opening")
});
publication.AddMetadataProperty("schema:accessMode", "textual");
publication.AddMetadataProperty("schema:accessibilityFeature", "structuralNavigation");
EpubWriteReport report = publication.Save("book.epub");
report.RequireNoLoss();
```

`AddChapter` accepts a well-formed XHTML **body fragment**, not a complete document or
arbitrary HTML. Escape text containing `&` or `<`. Stylesheet arguments are manifest ids.
Chapter and resource paths are canonical container paths; navigation targets are
URL-encoded container paths with optional fragments. A publication needs at least one
linear reading position. To create EPUB 2.0.1 with NCX navigation, pass
`version: EpubVersion.Epub2` to `Create`.

`AddResource` copies image, font, media, or other payload bytes into a manifest resource.
`SetCoverImage` selects an existing image. Spine APIs edit positions without deduplicating
repeated imported references. New positions require distinct manifest ids; repeated
content needs a separate chapter resource. `SetNavigation` accepts hierarchical TOC entries, page-list entries,
and landmarks (EPUB 2 uses NCX and guide references, with flat page-list and guide
entries). Generated XHTML navigation links respect retained local HTML base URLs;
navigation authoring rejects an external base that cannot address container content.
Clearing an EPUB 2 page list with retained headers or extension XML requires explicit
`GetContentXml`, serialization, and `UpdateResource` because NCX cannot retain an empty page-list shell.
`SetMetadataProperty` updates a
property; `AddMetadataProperty` adds repeatable values. `AddDublinCoreMetadata` adds
contributors, languages, or other Dublin Core values with optional ids and language.
EPUB 3 vocabulary prefixes, page progression, and rendition-layout declarations are
available. `SetRenditionLayout` declares a package default; use
[`SetFixedLayoutPage`](#fixed-layout-xhtml-pages) for an XHTML or SVG page canvas.

Accessibility metadata describes supplied content and does not certify conformance.
The writer adds SVG and MathML properties discovered in rewritten XHTML and remote-resource
properties discovered in rewritten XHTML/SVG. Existing remote-resource declarations remain
intact. Dependencies reached through external stylesheets require correct manifest
declarations from the caller and independent validation.

## Import an HTML manuscript

`EpubManuscript` converts inert HTML into a reflowable EPUB 3 publication. It splits
chapters at heading level 1 by default, retains nested heading navigation, and rewrites
internal links across chapter files. Lists, tables, code, notes, image alternatives,
SVG namespaces and source styles remain semantic content rather than page snapshots.

```csharp
using OfficeIMO.Epub;
using OfficeIMO.Html;

var source = HtmlConversionDocument.Parse(
    "<title>A short book</title><h1 id='opening'>Opening</h1><p>First chapter.</p>" +
    "<h1 id='closing'>Closing</h1><p><a href='#opening'>Return</a></p>");
EpubManuscriptResult imported = EpubManuscript.ImportHtml(source,
    new EpubManuscriptOptions { Language = "en", Creator = "A writer" });
imported.RequireNoLoss().Save("book.epub");
```

Set `ChapterHeadingLevel` to 0 for one chapter, or 1–6 for the split threshold.
Title, language and creator overrides take precedence over supported source metadata.
Embedded data resources work without a resolver. `ImportHtmlAsync` accepts an explicit
`ResourceResolver` for authorized external assets; the importer does not fetch the
network or open files implicitly. Resource collection includes inactive CSS imports,
fonts, responsive image candidates, and nested SVG dependencies. References are
rebased into declared package resources, with byte, count and total-resource bounds.
The defaults allow 256 resources, 10 MiB per resource and 50 MiB in aggregate.
Import review retains at most 10,000 findings; reaching that bound adds a failure
diagnostic instead of silently treating a truncated report as successful.

Inspect `Report.FidelityDiagnostics` before publishing. Executable content and active
attributes are omitted and reported. Missing assets, unresolved internal anchors,
invalid package content and unnamed images produce failure diagnostics. Images may
provide `alt`, an accessible ARIA name, or an explicit decorative role. The importer
does not invent image descriptions. `RequireValue` rejects failed imports;
`RequireNoLoss` also rejects reported approximations and omissions. Imported HTML is
not a full EPUB schema or accessibility certification route; independently validate
the final publication and review its content in representative readers.

DOCX and Markdown composition and editable book projects belong to
[`OfficeIMO.Workflows`](../OfficeIMO.Workflows/README.md#book-publishing).

### Read-aloud narration

`AddMediaOverlay` adds an EPUB 3 SMIL overlay for one existing XHTML document.
Add the audio resource first, then supply ordered cues and the measured duration
of each audio file:

```csharp
publication.AddChapter("story", "EPUB/story.xhtml", "Story",
    "<p id='first'>The first sentence.</p><p id='second'>The second sentence.</p>");
publication.AddResource("voice", "EPUB/voice.mp3", "audio/mpeg", File.ReadAllBytes("voice.mp3"));
publication.AddMediaOverlay("story", "story-narration", "EPUB/story.smil", new EpubMediaOverlay {
    AudioDurations = new Dictionary<string, TimeSpan> { ["voice"] = TimeSpan.FromSeconds(8) },
    Cues = new[] {
        new EpubMediaOverlayCue("first", "voice", TimeSpan.Zero, TimeSpan.FromSeconds(3)),
        new EpubMediaOverlayCue("second", "voice", TimeSpan.FromSeconds(3), TimeSpan.FromSeconds(8))
    }
});
```

The operation creates the SMIL resource, manifest association, overlay duration
and publication duration atomically. Clip durations are summed without floating-point
rounding. Each clip must satisfy `0 <= begin < end <= audio duration`. Audio durations
are caller declarations; OfficeIMO does not decode the media to verify them.

This profile accepts 1–10,000 cues and local, nonempty, unencrypted `audio/mpeg`
or `audio/mp4` resources. Targets use existing unqualified XHTML body element ids,
are distinct and non-nested, and follow document order. Clips can select different
files or reuse portions of a file; their list order determines playback. Narration
may cover only part of a chapter. An existing overlay is never overwritten.
Other retained SMIL resources need one unambiguous, nonnegative duration each,
expressible exactly as a `TimeSpan`, before the publication total can be recalculated.

`ReplaceMediaOverlay("story-narration", revisedOverlay)` replaces the same bounded
single-sequence profile. It preserves cue IDs for retained text targets and reserves
removed IDs during the edit, so new cues cannot accidentally inherit their links.
Removal fails when another inspectable resource still references a removed cue.
The operation preserves manifest identity and duration metadata attributes,
recalculates the publication total, and leaves audio resources unchanged. Shared
or encrypted overlays, extra SMIL structures/attributes, and processing instructions
are rejected. Validation, cancellation and size-limit failures leave the publication
unchanged. Other retained renditions must not depend on the replaced resource.

Use `SetMetadataProperty("media:active-class", "narration-active")` with a matching
CSS class in every narrated document for active-text styling. Validate the final
EPUB with EPUBCheck and check synchronization, highlighting, seeking and pause/resume
in target reading systems. General SMIL editing, nested skippable/escapable sequences,
SVG narration, synthesized speech and playback are outside this authoring profile.

### Fixed-layout XHTML pages

`SetFixedLayoutPage` configures an existing EPUB 3 XHTML or SVG spine document.
For XHTML it creates a viewport, CSS page canvas and typed presentation overrides:

```csharp
// For an entirely fixed-layout book, declare the package default too.
publication.SetRenditionLayout(EpubRenditionLayout.PrePaginated);
publication.AddChapter("plate", "EPUB/plate.xhtml", "Illustrated plate",
    "<main id='plate-content'>" +
    "<h1>Illustrated plate</h1><p>Selectable text in reading order.</p></main>");
publication.SetFixedLayoutPage("plate", new EpubFixedLayoutPage(800, 600) {
    Orientation = EpubPageOrientation.Landscape,
    Spread = EpubPageSpread.Both,
    Side = EpubPageSide.Right,
    Regions = new[] { new EpubFixedLayoutRegion("plate-content", 40, 40, 720, 520) }
});
```

For an SVG page, add an `image/svg+xml` resource with `AddResource`, place it in
reading order with `AddSpineItem`, and include it in `SetNavigation`. The same
`SetFixedLayoutPage` call sets root `width`, `height` and `viewBox="0 0 width height"`
and applies the spine overrides. It replaces any prior SVG viewport, including its
origin; it does not translate or resize the artwork. Child elements, transforms,
labels, links and `preserveAspectRatio` remain intact. SVG pages require an empty
`Regions` collection: position artwork with SVG coordinates and transforms.
Validate scaling, cropping and navigation in the intended reading systems.

Dimensions are positive integer CSS pixels. The method replaces the document's
single viewport declaration with `width` and `height`, and maintains a dedicated
canvas stylesheet setting HTML/body dimensions, zero margin/padding and a relative
body positioning context. Other content and styles remain intact. The method does
not fit overflowing text, paginate prose or rasterize text. Existing CSS and reader
styles can override the generated rules. Overflow is not hidden automatically.

`Regions` places existing top-level XHTML body elements by their HTML `id`. Each
rectangle supplies left, top, width and height as decimal CSS pixels, with a top-left
origin even in RTL books. Generated rules use absolute positioning, zero margins and
`border-box` sizing, so padding and borders belong inside the supplied dimensions.
Nested content retains its semantic structure and normal layout within each region.
Use a top-level section, figure or other appropriate XHTML container for grouped
content; nested targets and standalone SVG elements require a different layout policy.

Every region must fit inside the canvas with positive width/height and nonnegative
coordinates. Rejection identifies the target and canvas dimensions, and leaves the
whole publication unchanged. This checks declared boxes, not actual glyph, image,
transform or shadow overflow. Inspect rendered content at the intended fonts and
reader settings. Intentional region overlap is allowed; DOM order remains the
logical reading order and determines normal painting order.

A page accepts at most 1024 regions, each with an identifier of at most 1024 UTF-16
code units. Duplicate or missing targets are rejected. Each call replaces the complete
set of generated region rules; the default empty collection removes prior generated
placement while preserving authored CSS. Existing inline geometry or more specific
styles may take precedence, so remove competing declarations when the generated
boxes should control placement. IDs, content, links, ARIA relationships and DOM order
are not rewritten; CSS selectors safely quote punctuation and Unicode identifiers.

The selected spine position receives `rendition:layout-pre-paginated`, orientation
and spread overrides. `Auto` resets that aspect to the reading system's default;
`Center` requests a single centered page. The method replaces existing overrides
for these aspects, including vocabulary aliases, while preserving other properties
and package defaults. Set `PageProgressionDirection` separately for LTR/RTL books.
EPUB permits mixed reflowable and fixed-layout books using item overrides, but reader
support varies. Set the package default to `PrePaginated` for an entirely fixed-layout
book. In the Apple Books macOS check, item-only declarations clipped the landscape
fixture; the package declaration rendered the whole canvas in its right-hand spread
slot. The portrait page still appeared beside the reader's end-of-book panel instead
of centered alone. Treat mixed layouts and per-page spread overrides as reader-specific
qualification requirements.

Content and spine edits commit together after reference, byte-budget and cancellation
checks. The API rejects EPUB 2, resources other than XHTML/SVG, and missing or
repeated spine positions. For XHTML it also rejects multiple viewport declarations
and a conflicting use of its reserved `officeimo-fixed-layout-canvas` style identifier.
Existing XHTML viewport options beyond width/height are intentionally replaced.
Repeated XHTML calls update one canvas stylesheet.

DOM order, identifiers, links and accessibility relationships remain unchanged.
Supply content in logical reading order; visual coordinates do not establish that
order. [Executable fixtures](../Build/Epub/Fixtures/FixedLayoutFixture.cs) cover
landscape/portrait pages and LTR/RTL progression. EPUBCheck and browser canvas
inspection are separate from native reader spread, rotation, scaling and
assistive-technology qualification. Those reader checks remain open.

### Reflowable typography

Choose a reusable baseline when importing a manuscript:

```csharp
var options = new EpubManuscriptOptions {
    TypographyProfile = EpubTypographyProfile.Prose
};
```

`Basic` preserves the existing minimal manuscript stylesheet. `Prose` adds paragraph
spacing and indentation, heading break hints, and wrapping table cells. `Technical`
adds table and code borders that inherit the text color, wrapping code, and relative
monospace sizing. Both richer profiles use logical spacing and leave body font, body
font size, text color, and background to the reader or publisher. They do not add
fonts, fixed page widths, or `!important` declarations. Source CSS follows the baseline
and remains subject to the normal cascade. `IncludeDefaultStyles = false` omits the
baseline entirely.

For authored chapters or a book project, reuse the same CSS through
`EpubTypography.CreateStylesheet(EpubTypographyProfile.Technical)` with
`AddStylesheet` or `BookProject.SetStylesheet`. These profiles are starting points,
not guaranteed pagination or reader compatibility. Long table cells can overflow
with the minimal `Basic` profile; use a richer profile or publisher CSS for those
documents. Font embedding, vertical writing, and native reader behavior require
separate qualification. The [opt-in fixture generator](../Build/Epub/README.md#typography-fixtures)
produces publication bytes and browser previews for checking those boundaries.

## Edit and preserve a package

```csharp
using OfficeIMO.Epub;
using System.Xml.Linq;

EpubPublication publication = EpubPublication.Load("book.epub");
publication.Title = "Revised title";
string chapterId = publication.Spine[0].ManifestId;
XDocument chapter = publication.GetContentXml(chapterId);
XNamespace xhtml = "http://www.w3.org/1999/xhtml";
chapter.Root!.Element(xhtml + "body")!.Add(new XElement(xhtml + "p", "An added paragraph."));
publication.SetContentXml(chapterId, chapter);
EpubWriteResult result = publication.Write();
Console.WriteLine($"{result.Report.PreservedEntries.Count} unchanged entry payloads");
publication.Save("revised.epub");
```

Loaded publications retain all bounded file entries, including unmanifested extension
payloads, other rootfiles, unknown OPF nodes, and declaration attributes. An unedited
write with default options returns the exact original compressed package. Selecting
`CompressEntries = false` rebuilds the archive with stored entries while retaining
unchanged payloads. Edited output rewrites the selected
OPF and changed content; unchanged entry payloads retain their bytes. ZIP order,
timestamps, compression, XML formatting, and lexical prefixes may differ after editing.
`GetPackageXml` returns an inspection copy; typed metadata, manifest, and spine APIs
edit the retained package. `GetContentXml` / `SetContentXml` support targeted XHTML/SVG
editing. Instances are mutable and are not thread-safe.

Use the dictionary overload to apply coordinated content edits atomically. For example,
when changing an anchor ID, include both its document and documents containing incoming
links in the same batch:

```csharp
var target = publication.GetContentXml("chapter-1");
var source = publication.GetContentXml("chapter-2");
// Edit the target ID and its incoming links in these independent XML copies.
publication.SetContentXml(new Dictionary<string, XDocument> {
    ["chapter-1"] = target,
    ["chapter-2"] = source
});
```

The batch validates content identifiers, local ID references and publication links before
replacing any retained bytes. Duplicate IDs, dangling links, invalid content, retention
limits or cancellation leave the publication unchanged. Caller-owned XML remains independent.
This overload requires a structurally valid resulting publication; the single-document
setter remains available while assembling a book. Export still applies signature,
encryption and output policies, and independent accessibility and reader checks remain separate.

For editorial changes to individual elements, use `ApplyContentEdits` with immutable
`EpubContentEdit` proposals. Each proposal selects a manifest ID and `id`/`xml:id`,
retains the expected element, and supplies a replacement (or `null` to delete it).

```csharp
var document = publication.GetContentXml("chapter-1");
var expected = document.Descendants().Single(e => (string?)e.Attribute("id") == "paragraph-1");
var replacement = new XElement(expected);
replacement.Value = "Revised paragraph text.";
publication.ApplyContentEdits(new[] {
    new EpubContentEdit("chapter-1", "paragraph-1", expected, replacement)
});
```

The operation compares expected XML with current content and rejects stale proposals.
Targets must be inside an XHTML body or below an SVG root. Up to 10,000 independent,
non-overlapping targets can be edited together; surrounding content and document
scaffolding remain intact. Namespace-aware XML equality includes attributes and
whitespace. Proposals copy their XML. Identifier changes and deletions require all
remaining references to be valid in the combined batch; links are not guessed or
silently removed. The same validation and atomicity rules as the document batch apply.

Compare editions with `baseline.CompareTo(revised)`. The read-only result separates
package metadata, spine order/attributes, package scaffolding and resource changes.
Resources match by manifest ID; retained entries outside the manifest match by path.
Resource flags distinguish additions/removals, resolved location, declarations, XML
structure, text, serialization-only differences and uninterpreted binary changes.
The unrefined `dcterms:modified` write timestamp is excluded from metadata comparison.

`TextChanges` reports changed XHTML leaf text blocks (paragraphs, headings, list and
definition items, quotations, preformatted text, table cells and figure captions),
plus a `$body` aggregate to cover loose text. Blocks match by unique `id`/`xml:id`,
or by namespace-aware element position when no unique ID exists. Position matching
can report shifted blocks after insertion; it does not infer editorial intent.
Before/after excerpts are limited to 4096 UTF-16 characters and carry `IsTruncated`;
comparison uses full text. More than 10,000 changed text blocks rejects the operation
instead of returning an apparently complete partial report.

Comparison does not render content, normalize whitespace, interpret CSS, decrypt
assets, fetch remote resources or infer renamed manifest IDs. Added/removed resources
have resource flags; text excerpts compare resources present in both editions. XML
parse failures remain explicit. A report with no changes describes these source
contracts, not visual, accessibility or reader equivalence.

`RenameResource(manifestId, containerPath)` moves a local resource without changing its
manifest identifier or spine position. It repairs incoming links and rebases relative
references inside a moved document. Standard OPF links, XHTML/SVG links and resource
carriers, responsive image candidates, inline/external CSS and XML stylesheet
processing instructions, NCX navigation and SMIL
text/audio references use the same operation. Queries and fragments are retained.
For example:

```csharp
publication.RenameResource("chapter-1", "EPUB/parts/introduction.xhtml");
```

Renaming requires one rendition, no encryption/font obfuscation or scripting, no
`xml:base`, SVG animation or refresh navigation, and inspectable resource types. XHTML, SVG, NCX, SMIL, CSS and supported
raster/font/audio/video resources are inspected or retained as appropriate; unknown
resource formats and unknown XML processing instructions are rejected rather than
assumed to contain no references. Container
controls and rootfiles cannot be renamed. All changes are staged and validated before
commit; rejection leaves the publication unchanged. SMIL reference repair does not
qualify timing or playback. Export signature policy still applies.

`SplitChapter` divides a reflowable XHTML chapter before an identified block. The
selected block and following content become the next spine item, with a new manifest
ID, resource path, title, and sibling TOC entry:

```csharp
publication.SplitChapter("chapter-1", "second-section", "chapter-2",
    "EPUB/parts/second.xhtml", "Second section");
```

The operation clones the head and surrounding `body`, `div`, `section`, `article`,
and `main` containers. Incoming fragment links follow moved IDs; whole-document
links continue to target the first chapter. IDs on copied containers remain local
to each chapter, while incoming links to those shared IDs retain the first target.
Relative assets and links are rebased. Existing TOC nesting, page-list and landmark
entries are retained with repaired targets; the operation does not reorganize their
hierarchy. A primary TOC link that targeted moved content is reset to the first
chapter, and the new sibling targets the split boundary.

Both chapters must retain content. Cuts inside inline, table or list structures,
and cuts separating local accessibility, form-control, microdata or image-map
references, are rejected. Splitting
requires exactly one spine position for the source and the same inspectable-resource
profile as renaming. Shared manifest resources, fallbacks, media-overlay chapters,
and fixed-layout chapters require separate policies and are rejected. Validation,
retention limits and cancellation apply before any retained state changes. The
operation preserves markup, but changed chapter boundaries can affect pagination,
CSS counters and layout; assess the result in the intended reading systems.

`MergeChapters` combines consecutive reflowable reading positions into the first
resource. Supply an unused content ID for the second chapter's boundary:

```csharp
publication.MergeChapters("chapter-1", "chapter-2", "second-chapter-start");
```

Both TOC entries and their nesting remain. Whole-document links to the second chapter
target the new boundary; fragment links follow their retained elements. Relative assets,
HTML base URLs, page-list entries and package links are repaired before the second
manifest/spine entry and payload are removed. Its title is retained on the boundary
marker; the merged document keeps the first title. Identical structural containers at
the join, with matching IDs and attributes, are recombined, including containers cloned
by `SplitChapter`.

The default merge requires matching root/body attributes and equivalent heads after URL
rebasing, apart from title text. It rejects conflicting styles, metadata, processing
instructions, remaining duplicate IDs or image-map names, different reading-position
attributes, and package refinements that would lose their target. Document-local
relationships cannot be redirected to an empty boundary marker. Resolve these conflicts
explicitly before merging. To intentionally share both chapters' styles in one cascade:

```csharp
publication.MergeChapters("chapter-1", "chapter-2", "second-chapter-start",
    new EpubChapterMergeOptions { StylePolicy = EpubChapterMergeStylePolicy.AppendSecondStyles });
```

This policy retains the first head and appends every second-head `style` and stylesheet
`link` in source order after URL rebasing. It preserves duplicates because repeating a
stylesheet can affect the cascade. Media, title and other stylesheet attributes remain;
all other head content and attributes must still match apart from title text. Duplicate
content IDs, scaffold differences and package-refinement conflicts still fail atomically.
Later rules can restyle **both** chapters. This is an explicit cascade choice, not CSS
isolation or a promise to preserve each chapter's original appearance. Assess the merged
result in the intended readers. There is no automatic identifier-renaming policy. The same reflowable, resource-inspection, retention and atomicity limits as
splitting apply. Reader layout and accessibility assessment remain separate checks.

`EpubWriteReport` identifies preserved, regenerated, and removed entries. Its
`RenamedEntries` map distinguishes original paths relocated by the rename API from
content omissions. `MergedEntries` identifies original resources consolidated into
retained chapters; later renames follow that identity, and deleting the retained
chapter reports the corresponding original resources as omitted. Reports describe
operations relative to the current instance's load baseline; this history is not
embedded in the EPUB, and reopening establishes a new baseline. A rename or supported
merge alone does not fail `RequireNoLoss`. This is resource-identity evidence, not a
semantic comparison of arbitrary content edits. Explicit
resource or signature removal produces omission diagnostics; `RequireNoLoss` rejects
them. `RemoveResource` blocks structural and cover references, declared rootfiles,
and payloads referenced by retained alternate packages or their XHTML/SVG/NCX content.
Removal fails before mutation if an alternate package or its XML content cannot be
inspected safely. Update content and navigation links before removing their targets.
Raw resource APIs reject replacement or removal of `mimetype`, declared rootfiles,
and `META-INF` controls. A manifested signature may be removed from the model, but
writing still requires the explicit signature-removal policy described below.

Imported duplicate spine references retain their reading positions and produce
`EPUB_WRITE_RETAINED_DUPLICATE_SPINE`. This preserves the source's semantics but retains
its nonconformance; `RequireNoLoss` checks omissions, not EPUB validity.

MIME type comparisons accept equivalent casing while preserving declared spelling.
EPUBCheck 5.4.0 flags uppercase EPUB 3 navigation and cover media types. For these
imports, assign lowercase `MediaType` values explicitly when validator compatibility is required.

Editing a package with `signatures.xml` or a ZIP central-directory signature fails unless
`RemoveInvalidatedSignatures = true` is supplied in `EpubWriteOptions`; removal is reported.
IDPF/Adobe-obfuscated font bytes remain unchanged when editing
other content; replacing them or changing their package identity is rejected.
Unsupported encrypted-resource edits are rejected. Unedited supported imports retain
their protection metadata and ciphertext; the writer neither decrypts nor re-keys them.
Editable loading rejects unreadable, over-budget, or ambiguous encryption declarations
so protection guards cannot be bypassed by incomplete classification.

## Footnotes and endnotes

`AddNote` connects a labelled XHTML anchor to a note and generates its return link.
The reference marker has an `id` and no `href`; its text and other attributes are retained.
The source and notes may be in the same chapter or separate chapters.

```csharp
publication.AddChapter("chapter", "EPUB/text/chapter.xhtml", "Chapter",
    "<h1>Chapter</h1><p>A statement <a id='note-ref'>1</a>.</p>");
publication.AddChapter("notes", "EPUB/back/notes.xhtml", "Endnotes",
    "<section><h1>Endnotes</h1><ol id='notes-list'/></section>");
publication.AddNote(new EpubNoteOptions {
    SourceManifestId = "chapter", ReferenceId = "note-ref",
    NotesManifestId = "notes", ContainerId = "notes-list",
    NoteId = "note-1", Kind = EpubNoteKind.Endnote,
    BodyXhtml = "<p>Supporting detail.</p>",
    BacklinkText = "Return to reference"
});
```

Footnotes use an `aside` with `epub:type="footnote"` and `role="doc-footnote"`,
appended to the selected `section`, `div`, or `body`. Endnotes use native list items
inside an `ol` or `ul`; their enclosing section receives endnotes semantics.
The reference and backlink receive `doc-noteref` and `doc-backlink` roles. Supply
localized reference and backlink text appropriate to the book.

Both document edits are atomic: duplicate identifiers, unresolved local ARIA references,
scripts, conflicting roles, encrypted content, and combined retention-limit failures
leave the publication unchanged. Generated links respect local HTML base URLs;
external bases are rejected. EPUB 2 note authoring is unsupported. Full save and
independent accessibility/reader checks still apply; note popup behavior depends on the reader.

## Publishing metadata

`AddCreator` and `AddContributor` append EPUB 3 records with ordered MARC roles,
sorting names, and name languages. `AddCollection` records membership in a series
or set. Each operation commits the record and its refinements together and rejects
an identifier already used anywhere in the package.

```csharp
publication.AddCreator("author-alice", new EpubContributorMetadata {
    Name = "Alice Example", FileAs = "Example, Alice", Language = "en",
    MarcRoles = new[] { "aut", "ill" }
});
publication.AddContributor("translator-jan", new EpubContributorMetadata {
    Name = "Jan Kowalski", Language = "pl", MarcRoles = new[] { "trl" }
});
publication.AddCollection("chronicles", new EpubCollectionMetadata {
    Name = "The Example Chronicles", Kind = EpubCollectionKind.Series,
    FileAs = "Example Chronicles, The", Position = new uint[] { 2, 1 }
});
```

Positions are hierarchical: `2, 1` produces `2.1`; an empty position list leaves
the order unspecified. MARC roles retain their supplied priority, with repeated
codes collapsed. The API checks the three-letter code shape; independent validation
checks the vocabulary. Omitting roles leaves the contribution unspecified.

Existing metadata and refinements remain intact, including other creators and
collections. These append operations do not replace the primary `Creator` property.
Titles, identifiers, subjects, and primary publication details have typed operations:

```csharp
publication.SetPrimaryTitle("main-title", new EpubTitleMetadata {
    Text = "The Example Book", Kind = EpubTitleKind.Main,
    FileAs = "Example Book, The", DisplaySequence = 1, Language = "en"
});
publication.AddTitle("subtitle", new EpubTitleMetadata {
    Text = "An illustrated introduction", Kind = EpubTitleKind.Subtitle,
    DisplaySequence = 2
});
publication.AddIdentifier("isbn", new EpubIdentifierMetadata {
    Value = "978-0-306-40615-7", Kind = EpubIdentifierKind.Isbn13
});
publication.AddSubject("fiction", new EpubSubjectMetadata {
    Text = "FICTION / General", Authority = "BISAC", Code = "FIC000000"
});
publication.SetPublicationDetails(new EpubPublicationDetails {
    Publisher = "Example Press", Description = "An illustrated introduction.",
    Rights = "Publisher-supplied rights statement.",
    PublicationDate = new DateTime(2026, 10, 5)
});
```

`SetPrimaryTitle` updates the first title; when it already has an identifier, supply
that same identifier to preserve inbound refinements. Null optional fields retain
existing values. `AddTitle` appends a variant without replacing the first title;
display sequence refines presentation order without reordering XML records.

ISBN authoring validates length, characters, prefix for ISBN-13, and checksum, then
writes an ISBN URN. ISBN-10 supports historical records. DOI input is a DOI name,
such as `10.1000/182`, rather than a resolver URL. `Unspecified` retains an opaque
identifier without a type refinement. Adding an identifier leaves the selected
package identity unchanged. Neither registration nor ISBN range allocation is
checked, and subject authority/code pairs are publisher-supplied classifications.

`SetPublicationDetails` changes only supplied primary values in one atomic edit;
null fields retain existing values and additional publishers/declarations remain.
Publication dates use the supplied calendar date without a timezone conversion.
All typed publishing operations above require EPUB 3. Use `AddDublinCoreMetadata`
and `SetMetadataProperty` for other declarations; typed records do not infer rights,
publisher identities, or accessibility claims.

## Bibliographies and citation links

`AddBibliographyEntry(manifestId, listId, entryId, entryXhtml)` appends a formatted
XHTML entry to an existing `ol` or `ul` directly inside a body `section`. It marks
the section with `epub:type="bibliography"` and `role="doc-bibliography"`, preserving
its heading and existing entries. Entries use native `li` semantics and retain
caller order. Formatting, citation numbering, sorting, and disambiguation remain
with the publisher or the existing `OfficeIMO.Bibliography` CSL renderer.

```csharp
publication.AddChapter("references", "EPUB/back/references.xhtml", "References",
    "<section><h1>References</h1><ol id='entries'/></section>");
publication.AddBibliographyEntry("references", "entries", "work-one",
    "Example Author. <em>An illustrative source</em>. 2026.");
// The chapter already contains a labelled <a id="citation-one">[1]</a>.
publication.LinkBibliographyEntry("chapter", "citation-one", "references", "work-one",
    "Return to the citation");
```

Citation links receive `doc-biblioref` and `epub:type="biblioref"`. Each occurrence
can link to the same entry with its own optional localized return link. Omit the
return label to preserve a separate bibliography document's bytes. Links respect
both documents' effective HTML bases. Existing links, conflicting roles, invalid
targets, duplicate identifiers and over-budget edits are rejected; cross-document
changes commit atomically. Full schema, accessibility and reading-system checks
remain separate from these native content checks.

To consume the managed CSL renderer, request `CslOutputFormat.Html` from
`OfficeIMO.Bibliography`, then pass each nonempty `rendered.Bibliography` entry's
`Content` to `AddBibliographyEntry` in the renderer's returned order. Map its citation
keys to unique XML-compatible content identifiers; arbitrary source keys need not
be valid EPUB IDs. For single-item citations, retain the rendered citation label
inside the source anchor before calling `LinkBibliographyEntry`. A citation cluster
covering several sources needs separate per-source links or publisher-authored
navigation; one anchor cannot target multiple entries.

The [executable fixture](../Build/Epub/Fixtures/BibliographyFixture.cs) exercises
escaped text, italic titles, title sorting and repeated single-item citations.
CSL hanging indents, spacing and second-field alignment still require publisher CSS
derived from the renderer's layout result and independent presentation checks.
`OfficeIMO.Epub` does not take a runtime dependency on the citation renderer.

## Indexes

`AddIndexEntry` appends a plain-text term and labelled links to an existing `ul`
directly inside a body `section`. It marks that section with `epub:type="index"`
and `role="doc-index"`. Existing headings and entries remain intact, and publisher
order is retained.

```csharp
publication.AddChapter("index", "EPUB/back/index.xhtml", "Index",
    "<section><h1>Index</h1><ul id='entries'/></section>");
publication.AddIndexEntry("index", "entries", "publishing", "Publishing",
    Array.Empty<EpubIndexLocator>(), subentriesId: "publishing-entries");
publication.AddIndexEntry("index", "publishing-entries", "reading-order", "reading order",
    new[] { new EpubIndexLocator {
        ManifestId = "chapter", FragmentId = "reading-order-heading", Label = "Reading order"
    } });
```

The optional `subentriesId` creates a nested list for subsequent entries; fill it
before publishing. Each entry requires locators or a subentry list. Locators target
existing XHTML spine documents, optionally selecting one unambiguous body-content
identifier. They can also target an existing index term for a cross-reference.
Omit `FragmentId` for a whole-document link. Labels can be source-page labels or
section titles; this API does not infer screen pages, sort terms, or extract an index
from prose. Each call accepts at most 1024 locators and observes cancellation.

Links respect the index document's effective HTML base. Target content stays
unchanged, while invalid targets, conflicting roles, duplicate identifiers and
retention-limit failures leave the index unchanged. Native reader interaction and
human accessibility assessment remain separate qualification steps.

## Glossaries

Append entries to an existing definition list, then link any number of occurrences
to a term. Labels and definitions remain publisher-authored; entries retain insertion
order rather than applying an implicit language-dependent sort.

```csharp
publication.AddChapter("glossary", "EPUB/back/glossary.xhtml", "Glossary",
    "<section aria-labelledby='glossary-title'><h1 id='glossary-title'>Glossary</h1>" +
    "<dl id='terms'/></section>");
publication.AddGlossaryEntry("glossary", "terms", "reflowable", "Reflowable book",
    "<p>A book whose text adapts to the reading area and reader settings.</p>");
// The chapter already contains <a id="term-reference">reflowable book</a>.
publication.LinkGlossaryTerm("chapter", "term-reference", "glossary", "reflowable",
    "Return to the passage");
```

`AddGlossaryEntry` escapes plain term text into `dt/dfn`, appends the well-formed
XHTML definition in `dd`, marks the containing section as an EPUB
glossary, and assigns `doc-glossary` to the section. The supported container is a
`dl` directly inside a body `section`, with complete direct `dt/dd` groups. Existing
headings, entries, attributes and unrelated content are retained; conflicting
section roles are rejected. Definition links use the receiving document's effective
HTML base. This API does not generate a dictionary, translate terms, or infer definitions.

`LinkGlossaryTerm` requires a labelled body anchor without an existing link and a
target `dt` followed by exactly one `dd` in that semantic glossary. It adds
`doc-glossref` and `epub:type="glossref"`. The optional localized return label adds
a `doc-backlink` in the definition; omit it to leave a separate glossary document
byte-for-byte unchanged. Each linked occurrence can have its own return link.
Both documents commit atomically with identifier, content, cancellation and retention
checks. These links use standard navigation; reader-specific popup behavior and
assistive-technology interaction require independent qualification.

## Book matter and print-page navigation

`SetDocumentMatter` marks an XHTML document as front, body, or back matter while
retaining its other semantic tokens. It does not change spine order or the table of contents.
`AddPrintPageMarker` marks an existing empty span as a source-page boundary and adds
its labelled link to the EPUB 3 page list.

```csharp
publication.AddChapter("preface", "EPUB/front/preface.xhtml", "Preface",
    "<h1>Preface</h1><p><span id='page-iv'/>Opening text.</p>");
publication.SetDocumentMatter("preface", EpubDocumentMatter.FrontMatter);
publication.AddPrintPageMarker("preface", "page-iv", "iv", "Reference pages");
publication.SetMetadataProperty("pageBreakSource", "urn:isbn:9781234567897");
```

Add markers in the source edition's page order and identify that edition with
`SetMetadataProperty("pageBreakSource", sourceIdentifier)`, using its actual identifier
or a description that uniquely identifies the source. Page labels and
the new page-list heading are caller-supplied text and can be localized. Existing
page-list headings, attributes, and entries are retained; adding a duplicate target
is rejected. Content and navigation edits commit together, including retention-limit
and cancellation checks. Generated links respect local HTML base URLs.

These APIs require EPUB 3. Page markers belong to spine documents and receive
`epub:type="pagebreak"`, `role="doc-pagebreak"`, and an accessible page label.
They reference an existing edition's pages; they do not paginate text or create
fixed-layout geometry. Independent reader and accessibility qualification still applies.

## Publication preflight

`Preflight` returns structured native checks without writing a destination or invoking
external tools. It uses the supplied save policy, inspects retained XHTML/SVG identifiers
and document-local ARIA/table-header references, and checks HTML image alternative
presence and EPUB 3 accessibility discovery declarations. The `media-overlays` check
verifies retained SMIL text/sequence targets, manifest associations, audio declarations,
explicit clip intervals and duration sums. It also inspects unchanged imported content, even when saving preserves that
content byte-for-byte.

```csharp
EpubPreflightReport review = publication.Preflight();
foreach (EpubPreflightCheck check in review.Checks) {
    Console.WriteLine($"{check.Code}: {check.Status}");
    foreach (EpubDiagnostic finding in check.Diagnostics) {
        Console.WriteLine($"{finding.Path}: {finding.Code}: {finding.Message}");
    }
}
```

`HasErrors` describes the executed native checks. `HasUncheckedItems` remains true
because full EPUB schema conformance, comprehensive accessibility assessment, and
reading-system presentation need independent validation and review. Image alternative
presence does not establish whether the descriptions are useful. Each native check
retains at most 10,000 findings and reports an error when that bound is reached.
The report is a snapshot; rerun it after editing. `RequireNoLoss` remains a separate
conversion/preservation check.

Narration checks accept nested sequences and text-only cues. An omitted `clipEnd`
or encrypted target leaves the relevant inspection explicitly `NotChecked` unless
another finding makes it fail. Duration sums tolerate up to one second of rounding.
The separate `media-overlay-audio-decoding` check remains `NotChecked`: valid clock
values and matching totals do not prove that offsets fit the encoded audio or that
speech aligns with text. Full SMIL schema validation and reader playback remain
independent checks. `Write` preserves its existing save policy; callers use `Preflight`
and inspect its findings before delivery.

Supply publisher-reviewed discovery claims together with `SetAccessibilityMetadata`:

```csharp
publication.SetAccessibilityMetadata(new EpubAccessibilityMetadata {
    AccessModes = new[] { "textual" },
    SufficientAccessModes = new IReadOnlyList<string>[] { new[] { "textual" } },
    Features = new[] { "structuralNavigation", "tableOfContents" },
    Hazards = new[] { "unknown" },
    Summary = "Text and chapter navigation are available. Hazard review is incomplete."
});
```

Each sufficient-mode list describes one combination; multiple lists express alternatives.
The operation replaces these five publication-level properties atomically, preserves
unrelated metadata and refinements, and rejects removal of a declaration referenced by
a refinement. Null `Summary` and empty `SufficientAccessModes` remove those optional
properties. Vocabulary values are publisher-supplied tokens; review them against the
discovery vocabulary and the actual content. The API does not infer claims or certify them.
Preflight reports missing access modes, features, and hazards as errors; missing sufficient
modes and summary are warnings. EPUB 2 discovery metadata is explicitly unchecked.

Newly authored or rewritten content rejects duplicate IDs and unresolved document-local
ARIA, table-header, form-control and microdata ID references, plus missing local image-map
names. IDs may repeat in separate content documents.
Manuscript splitting reports a failure when a relationship crosses the resulting
chapter boundary; keep the related content in one chapter or repair the source.
`AddDublinCoreMetadata` accepts the standard Dublin Core element names, with exact
lowercase spelling; custom metadata belongs in the EPUB property vocabulary APIs.

For repeatable external validation, see the [EPUBCheck evidence runner](../Build/Epub/README.md).

The [independent publishing fixture](../OfficeIMO.TestAssets/Documents/Epub/IdpfPublishing/README.md)
tests creator refinements, nested navigation, 92 print-page references, and editorial
metadata changes against an unchanged EPUB 3 Samples publication. Its non-package
payloads remain byte-for-byte intact after metadata editing. This supplements the
[W3C spine fixtures](../OfficeIMO.TestAssets/Documents/Epub/W3cSpine/README.md);
neither corpus establishes full accessibility or reader-presentation conformance.

## Save validation and limits

Save preflight checks required metadata, unique package ids, manifest targets and
relationships, acyclic fallbacks, linear spine positions, navigation, direct content
URLs, responsive image targets, and XHTML/SVG fragment ids. Imported packages need
canonical, unique ZIP paths and a physically leading stored `mimetype` entry.
Cover declarations are checked against the final manifest, including their image
type and the EPUB 3 single-cover property. `SetCoverImage` also updates retained
legacy cover metadata, so replacing a cover keeps both declarations consistent.
Media-overlay associations require an
EPUB 3 content document and a SMIL target. `AddMediaOverlay` validates its authored
cues and calculates durations. `Preflight` additionally checks retained overlay
references, explicit timing and duration sums after low-level edits. Saving alone
does not run those additional checks or decode audio.
Spine items resolve to XHTML in EPUB 2, or XHTML/SVG in EPUB 3, through any fallback
chain. EPUB 2 image and stylesheet resources belong inside content documents;
direct spine references to them are rejected, including SVG with a fallback.
Navigation targets require explicit spine entries. A fallback can be listed as a
non-linear entry to make it navigable without repeating it in the primary reading order.
Remote-resource
declarations and URL-policy checks include retained media/source alternatives and
stylesheet links, even when the renderer selects a different alternative.
New or rewritten XHTML/SVG requires manifest declarations for embedded container
resources and remote audio, video, or fonts. EPUB 3 remote images, stylesheets, and
embedded documents are rejected. Inline CSS references are checked across media
conditions. New or changed linked CSS is checked recursively for declared imports,
fonts and images, including inactive rules; cycles terminate without discarding CSS.
Removing resources cannot leave retained stylesheets with dangling dependencies.
Data URLs are limited to inert raster image, audio/video, and font contexts. Embedded
data documents, SVG data URLs, and data hyperlinks are rejected under the writer's
non-scripted contract. Referenced imported scripted resources cannot be newly embedded.
Newly authored or rewritten forms are rejected under the scripted-content restriction;
unchanged imported scripted content remains eligible for preservation.
Language authoring requires well-formed BCP 47 syntax, including private-use and
grandfathered tags. Edited packages validate every Dublin Core language value.
Unedited imports retain their original language declarations.
Custom vocabulary declarations reject EPUB-prohibited mappings. Declare custom
prefixes before assigning metadata, manifest, or spine properties. Newly assigned
property tokens require valid declared or reserved prefixes and nonempty references;
unrelated imported extension declarations remain intact.
External content is never downloaded or executed. CSS semantics, full schema
validation, accessibility certification, and EPUB conformance need independent validation,
such as [EPUBCheck](https://github.com/w3c/epubcheck).

| Boundary | Default |
| --- | ---: |
| Compressed input / output | 128 MiB each |
| Expanded retained / output bytes | 256 MiB each |
| Individual retained entry | 64 MiB |
| Package/container XML at load | 4 MiB |
| Archive entries | 10,000 |
| Authored navigation depth | 64 nested levels |

Use `EpubPublicationLoadOptions` for retained input and creation resource limits;
use `EpubWriteOptions` for output limits, modification time, and compression.
Retained entry limits include mandatory package entries and apply when adding resources.
Unedited writes count every physical ZIP record, including directories, against output
limits. XHTML/SVG/NCX inspection, editing, and write preflight use the configured
retained entry-byte bound; package XML uses the configured metadata bound.
Rewritten XHTML/SVG checks XML against that bound before shared HTML analysis,
allowing for canonical XML escaping while retaining the shared DOM, stylesheet,
and responsive-resource complexity safeguards.
Ordinary entries are deflated by default; `mimetype` is always stored first without
extra fields. Repeated writes of unchanged model state use stable modification metadata
and deterministic ZIP ordering. Supply `ModifiedAt` for a reproducible timestamp
across separately created publications with identical identity and content.

File saves use staged atomic replacement: validation, bounds, or cancellation failure
before commit preserves an existing destination. Seekable caller streams are replaced
and rewound; all caller streams remain open. Stream I/O failure or cancellation during
the final copy can leave partial bytes. Async saves stage serialization synchronously
before asynchronous destination I/O. Load and save APIs accept cancellation tokens.

`publication.Read(options)` projects the current publication through `EpubDocument`
for existing Reader and image-conversion consumers. It applies the writer's default
save policy first, including signature and output-limit checks.

## Content provenance

`EpubDocument.InspectProvenance("book.epub")` reports C2PA and AI-specific IPTC metadata in the EPUB package and supported embedded images. `EpubDocument.RemoveProvenance("book.epub", "clean.epub")` performs a targeted bounded rewrite while preserving the required uncompressed, first `mimetype` entry. Signed-package mutation is blocked unless removal of invalidated `META-INF/signatures.xml` is requested explicitly. Optional cryptographic C2PA verification remains in `OfficeIMO.Security`.

## Boundaries

- This package owns EPUB parsing, native authoring, package-preserving editing, and bounded HTML manuscript import.
- Reader integration belongs in `OfficeIMO.Reader.Epub`.
- `EpubDocument` remains a read-only extraction model; use `EpubPublication` for writing. XHTML fixed-page canvases and sequential SMIL narration have the bounded authoring profiles above. Browser layout, scripting, DRM, general encrypted-resource editing and audio playback are outside the writer contract. IDPF and Adobe font deobfuscation remains bounded reader behavior.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`, and `net472` on Windows.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** No third-party EPUB engine. Concealed-content inspection uses `AngleSharp` and `AngleSharp.Css` transitively through the shared `OfficeIMO.Html` owner.
- **OfficeIMO:** Direct dependencies are `OfficeIMO.Core` and `OfficeIMO.Html`; the HTML dependency also brings `OfficeIMO.Html.Core` and `OfficeIMO.Html.AngleSharp` (including its `System.Text.Encoding.CodePages` runtime dependency). Container, OPF, spine, navigation, chapter, and resource parsing remain first-party; the shared HTML owner supplies the bounded concealed-content and stylesheet model.
- **Security:** `META-INF/signatures.xml` discovery is structural and provider-free. Creation and validation of the bounded OfficeIMO XML package-manifest profile accept an explicit `IOfficeSecurityProvider`; `OfficeIMO.Security` is not pulled transitively.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 2 | 0 | 0 | 0 | 1 | 0 |
| Read | 2 | 0 | 0 | 0 | 0 | 1 |
| Edit | 1 | 0 | 0 | 1 | 0 | 1 |
| Preserve | 1 | 0 | 0 | 0 | 0 | 0 |
| Inspect | 4 | 0 | 0 | 0 | 0 | 0 |
| Validate | 2 | 1 | 0 | 0 | 0 | 1 |
| Remove | 2 | 0 | 0 | 0 | 0 | 1 |
| Convert | 0 | 1 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Epub` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
