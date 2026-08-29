# OfficeIMO.DocBook

`OfficeIMO.DocBook` creates, reads, edits, validates, converts, and writes a bounded common-structure profile of DocBook 4 and 5. It prioritizes articles, books, metadata, sections, paragraphs, lists, CALS-shaped tables, code, links, cross-references, admonitions, media, and indexes. It does not claim a typed object for every DocBook or extension element.

## Install

```shell
dotnet add package OfficeIMO.DocBook
```

```csharp
using OfficeIMO.DocBook;

DocBookDocument document = DocBookDocument.CreateArticle(DocBookProfile.DocBook52);
document.Title = "Deployment guide";

DocBookNode section = document.AddSection("Install");
section.AddParagraph("Install the package from NuGet.");
section.AddProgramListing("dotnet add package OfficeIMO.DocBook", "shell");
section.AddAdmonition(DocBookNodeKind.Note, "Validate before publishing.");

DocBookValidationResult validation = document.Validate();
document.Save("guide.docbook");
```

The typed API is a view over the backing `XDocument`. Unknown elements, attributes, comments, processing instructions, namespace declarations, and document types remain present. An unchanged file or stream is written byte-for-byte; after an edit, preserved XML is serialized around changed values. `AddExtension` and `Xml` provide deliberate access outside the typed common structure.

## Exact profiles and validation scope

`DocBookSchemaProfiles` exposes the exact official identifiers associated with each OfficeIMO profile:

- `DocBook45`: public ID `-//OASIS//DTD DocBook XML V4.5//EN` and system ID `http://www.oasis-open.org/docbook/xml/4.5/docbookx.dtd`.
- `DocBook52`: namespace `http://docbook.org/ns/docbook`, the OASIS 5.2 non-XInclude `docbook.rng`, and its Schematron rules.

`Validate()` performs `DocBookValidationScope.OfficeIMOCommonStructure`. Its result sets `IsOfficialSchemaValidated` to `false`: it checks the bounded authoring contract but does not pretend to run the complete external DTD, RELAX NG, Schematron, XInclude, assembly, ITS, or arbitrary extension schemas. Applications requiring formal conformance can run the exposed official artifacts in their chosen validator.

Reading accepts no-namespace DocBook 4 and namespaced DocBook 5 documents; canonical creation writes exactly 4.5 or 5.2. External entity resolution is disabled, and internal subsets containing external or parameter entity declarations are rejected; bounded internal general entities remain supported. Default limits are 32 MiB encoded input, 32 million XML characters, 256 levels, 250,000 elements, one million attributes, and 4,096 entity-expanded characters.

`ToOfficeDocumentModel()` and `FromOfficeDocumentModel()` use `OfficeDocumentModel.Structure` and deterministic loss diagnostics. The shared projection also publishes authors, links, and bounded CALS tables. Independent flat blocks, tables, assets, or links are appended at the document root when recursive structure does not already represent them, with a placement-loss diagnostic. Table projection defaults to at most 1,024 columns, 100,000 header rows, 100,000 body rows, and 1,000,000 rectangular cell slots per table; `DocBookConversionOptions` can lower those limits, and diagnostics report truncation or layout flattening while the native XML remains preserved. Validation and conversion retain at most 100 detailed diagnostics per code by default, followed by one occurrence summary; `DocBookValidationOptions.MaxDetailedDiagnosticsPerCode` and `DocBookConversionOptions.MaxDetailedDiagnosticsPerCode` can lower or raise that budget. `OfficeIMO.Reader.DocBook` adds dedicated `.dbk` and `.docbook` ingestion without taking over generic `.xml` routing.

Targets: `netstandard2.0`, `net8.0`, `net10.0`, and `net472` on Windows.

## Dependency footprint

- **External:** None; schemas are identified but not downloaded at runtime.
- **OfficeIMO:** `OfficeIMO.Core`.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
