# OfficeIMO.Bibliography

`OfficeIMO.Bibliography` is the citation-data owner for OfficeIMO. It provides one editable model and deterministic codecs for BibTeX, BibLaTeX, CSL JSON, RIS, PubMed NBIB/MEDLINE, and EndNote XML.

Install the package from NuGet:

```shell
dotnet add package OfficeIMO.Bibliography
```

## Read, edit, write, and reopen

```csharp
using OfficeIMO.Bibliography;

BibliographyReadResult read = BibliographyDocument.Load(
    "library.bib",
    BibliographyFormat.BibLatex);

BibliographyItem item = read.Document.Items[0];
item.Title = "A corrected title";
item.SetIdentifier("DOI", "10.1000/example");

BibliographyWriteResult saved = read.Document.Save(
    "library-edited.bib",
    new BibliographyWriteOptions {
        Format = BibliographyFormat.BibLatex,
        Mode = BibliographyWriterMode.Canonical
    });

BibliographyReadResult reopened = BibliographyDocument.Load(
    "library-edited.bib",
    BibliographyFormat.BibLatex);
```

The model includes citation keys, typed item kinds, personal and corporate contributors, partial, literal, and ranged dates, identifiers, titles, publication fields, pagination, URLs, keywords, notes, and ordered native fields. Unknown source fields remain available in item, name, and date `NativeFields`; BibTeX directives and safe document-level EndNote XML elements remain available in `NativeEntries`.

CSL JSON maps all 45 standard item types, all 26 contributor roles, and all six date roles into the typed model. Roles include directors, performers, original and reviewed authors, and container authors; `BibliographyDateRole.Available` represents `available-date`. Unknown and incorrectly shaped values remain native evidence. Other formats retain their own supported vocabulary and report unsupported typed roles, dates, and types through conversion diagnostics.

## Preserve the source or normalize it

Preserve mode is the default. An unchanged document loaded from bytes returns the original bytes exactly, including its BOM and line endings. An unchanged document parsed from text returns the original text exactly.

After an edit, the writer produces deterministic canonical syntax for the selected format. It retains unknown native fields when the destination is the same format and the field can be written safely. This is source-backed round-trip editing, but it does not promise to retain whitespace and comments inside a modified record at their original positions.

```csharp
BibliographyWriteResult exact = read.Document.Write();
bool reusedOriginalBytes = exact.UsedOriginalSource;

BibliographyWriteResult normalized = read.Document.Write(
    new BibliographyWriteOptions {
        Mode = BibliographyWriterMode.Canonical,
        LineEnding = "\n"
    });
```

## Convert with explicit fidelity evidence

Every write returns a `BibliographyConversionReport`. Native fields that cannot be represented by another format are reported as `Approximated` or `Omitted`; they are never silently discarded.

```csharp
BibliographyWriteResult csl = read.Document.Write(
    new BibliographyWriteOptions {
        Format = BibliographyFormat.CslJson,
        Mode = BibliographyWriterMode.Canonical,
        RequireNoLoss = true
    });
```

`RequireNoLoss` throws `BibliographyConversionLossException` when the destination would approximate, omit, or fail to preserve data. For permissive conversion, leave it disabled and inspect `csl.Report.Diagnostics`. The report also implements `IOfficeConversionReport`; use `FidelityDiagnostics` when composing bibliography output with another OfficeIMO conversion stage and the policy must retain the exact `None`, `Approximation`, `Omission`, or `Failure` category.

## File, stream, text, and async APIs

`BibliographyDocument` supports:

- `Parse` from text, with explicit format or bounded content detection
- `Load` and `LoadAsync` from paths and streams
- `Write` to text and bytes
- `Save` and `SaveAsync` to paths and caller-owned streams

Path loading recognizes `.bib`, `.json`, `.ris`, `.nbib`, `.medline`, and `.xml`. Unknown extensions use bounded content detection. Parsing observes item, value, input-size, nesting, and cancellation limits through `BibliographyReadOptions`.

Path saves stage output beside the destination and atomically publish it only after writing completes and cancellation is checked. If the filesystem cannot atomically replace an existing file, the save fails and preserves that file. A failed or cancelled save removes its staging file where filesystem permissions allow. Writes to caller-owned streams can be partial when cancelled.

## Resolve local bibliography references

```csharp
BibliographyReferenceResult resolved = read.Document.ResolveReferences(
    new BibliographyReferenceOptions { MaximumDepth = 64 }, cancellationToken);
BibliographyItem child = resolved.Document.Items[0];
IReadOnlyList<BibliographyFieldProvenance> origins = resolved.Provenance;
IReadOnlyList<BibliographyDiagnostic> referenceDiagnostics = resolved.Diagnostics;
```

Resolution is explicit and returns an independently editable snapshot. It does not change the source document or fetch records. Keys are case-sensitive. Duplicate keys, missing parents, invalid xdata targets, repeated reference fields, cycles, and excessive chain depth have stable `BIBREF` diagnostics; inspect `IsComplete` before relying on a fully resolved result.

Crossref fills missing fields and maps common book/proceedings/periodical titles to the child's container title. Existing child dates and contributor roles remain whole groups. Whole-entry xdata is applied in listed order before crossref: later containers replace earlier values and child values. Set `XDataOverridesExistingFields = false` for missing-field fallback instead. The resolver does not interpret granular `xdata=key-field-index` expressions or custom Biber inheritance rules.

`Document` retains source-order records, data containers, and reference fields for native writing. `CitationItems` excludes `@xdata` containers. `Provenance` identifies the ultimate source record/field and the complete reference path for each inherited field or role group; it describes the operation snapshot and does not change after edits. Item, edge, depth, copied-value, expanded-character, diagnostic, and cancellation limits bound the operation.

## Boundaries

The package does not execute TeX, fetch DOI or PubMed metadata, resolve remote resources, manage attachments, remove DRM, or decrypt resources.

`OfficeIMO.Word` does not depend on this package.

See the [bibliography support matrix](../Docs/officeimo.bibliography-support-matrix.md) for exact field, preservation, conversion, and security behavior.

## Dependencies

`OfficeIMO.Bibliography` depends on the zero-dependency `OfficeIMO.Core` package for the shared typed conversion-report contract. It uses `System.Text.Json` for CSL JSON and `System.Text.Encoding.CodePages` for declared legacy XML encodings on compatibility targets. It has no dependency on `OfficeIMO.Word`, the Open XML SDK, a TeX runtime, EndNote, or a network client.

<!-- officeimo-operation-catalog:start -->
## Generated capability summary

This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.

| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| Create | 1 | 0 | 0 | 0 | 0 | 0 |
| Read | 1 | 0 | 0 | 0 | 0 | 0 |
| Edit | 1 | 0 | 0 | 0 | 0 | 0 |
| Preserve | 1 | 0 | 0 | 0 | 0 | 0 |
| Inspect | 1 | 0 | 0 | 0 | 0 | 0 |
| Validate | 1 | 0 | 0 | 0 | 0 | 0 |
| Convert | 30 | 0 | 0 | 0 | 0 | 0 |

The complete rows for `OfficeIMO.Bibliography` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).
<!-- officeimo-operation-catalog:end -->
