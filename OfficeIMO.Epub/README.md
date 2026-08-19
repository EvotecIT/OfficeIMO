# OfficeIMO.Epub - EPUB extraction primitives

[![nuget version](https://img.shields.io/nuget/v/OfficeIMO.Epub)](https://www.nuget.org/packages/OfficeIMO.Epub)
[![nuget downloads](https://img.shields.io/nuget/dt/OfficeIMO.Epub?label=nuget%20downloads)](https://www.nuget.org/packages/OfficeIMO.Epub)

`OfficeIMO.Epub` provides reusable EPUB extraction primitives for modular OfficeIMO ingestion pipelines.

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

### Inspect package signatures

```csharp
EpubDocument book = EpubDocument.Load("signed.epub");

if (book.HasSignatures) {
    Console.WriteLine($"Signature elements: {book.Signatures.SignatureCount}");
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
- Extracts chapter text from XHTML/XML ASTs.
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

## Content provenance

`EpubDocument.InspectProvenance("book.epub")` reports C2PA and AI-specific IPTC metadata in the EPUB package and supported embedded images. `EpubDocument.RemoveProvenance("book.epub", "clean.epub")` performs a targeted bounded rewrite while preserving the required uncompressed, first `mimetype` entry. Signed-package mutation is blocked unless removal of invalidated `META-INF/signatures.xml` is requested explicitly. Optional cryptographic C2PA verification remains in `OfficeIMO.Security`.

## Boundaries

- This package owns reusable EPUB parsing primitives.
- Reader integration belongs in `OfficeIMO.Reader.Epub`.
- The content model is read-only. The provenance API provides only targeted carrier removal; it does not attempt CSS layout, scripting, DRM, or general package editing. IDPF and Adobe font deobfuscation is bounded reader behavior, not a general encryption API.

## Targets and license

- Targets: `netstandard2.0`, `net8.0`, `net10.0`.
- License: MIT.
- Repository: [EvotecIT/OfficeIMO](https://github.com/EvotecIT/OfficeIMO)

## Dependency footprint

- **External:** None; no third-party EPUB engine.
- **OfficeIMO:** `OfficeIMO.Core`. Container, OPF, spine, navigation, chapter, and resource parsing are first-party.
- **Security:** `META-INF/signatures.xml` discovery is structural and provider-free. Creation and validation of the bounded OfficeIMO XML package-manifest profile accept an explicit `IOfficeSecurityProvider`; `OfficeIMO.Security` is not pulled transitively.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
