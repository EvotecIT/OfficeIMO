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
available; declaring fixed layout does not generate page geometry.

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

`EpubWriteReport` identifies preserved, regenerated, and removed entries. Explicit
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

## Save validation and limits

Save preflight checks required metadata, unique package ids, manifest targets and
relationships, acyclic fallbacks, linear spine positions, navigation, direct content
URLs, responsive image targets, and XHTML/SVG fragment ids. Imported packages need
canonical, unique ZIP paths and a physically leading stored `mimetype` entry.
Cover declarations are checked against the final manifest, including their image
type and the EPUB 3 single-cover property. `SetCoverImage` also updates retained
legacy cover metadata, so replacing a cover keeps both declarations consistent.
Media-overlay associations require an
EPUB 3 content document and a SMIL target; SMIL timing is outside this validation.
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

- This package owns EPUB parsing, native authoring, and package-preserving editing.
- Reader integration belongs in `OfficeIMO.Reader.Epub`.
- `EpubDocument` remains a read-only extraction model; use `EpubPublication` for writing. Browser layout, scripting, DRM, general encrypted-resource editing, media-overlay authoring, manuscript import, and fixed-page geometry creation are outside the writer contract. IDPF and Adobe font deobfuscation remains bounded reader behavior.

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
