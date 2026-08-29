# OfficeIMO.Bibliography

`OfficeIMO.Bibliography` is the citation-data owner for OfficeIMO. It provides one editable model and deterministic codecs for BibTeX, BibLaTeX, CSL JSON, RIS, PubMed NBIB/MEDLINE, and EndNote XML without depending on Word or a citation-style engine.

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

`RequireNoLoss` throws `BibliographyConversionLossException` when the destination would approximate or omit data. For permissive conversion, leave it disabled and inspect `csl.Report.Diagnostics`.

## File, stream, text, and async APIs

`BibliographyDocument` supports:

- `Parse` from text, with explicit format or bounded content detection
- `Load` and `LoadAsync` from paths and streams
- `Write` to text and bytes
- `Save` and `SaveAsync` to paths and caller-owned streams

Path loading recognizes `.bib`, `.json`, `.ris`, `.nbib`, `.medline`, and `.xml`. Unknown extensions use bounded content detection. Parsing observes item, value, input-size, nesting, and cancellation limits through `BibliographyReadOptions`.

## Boundaries

The package parses data only. It does not execute TeX, fetch DOI or PubMed metadata, resolve remote resources, render citations or bibliographies, manage attachments, remove DRM, or decrypt resources. Citation-style rendering is a separate product boundary.

`OfficeIMO.Word` does not depend on this package. A future optional Word bridge can map Word bibliography XML to this model without putting Word/Open XML types into the citation-data owner.

See the [bibliography support matrix](../Docs/officeimo.bibliography-support-matrix.md) for exact field, preservation, conversion, and security behavior.

## Dependencies

`OfficeIMO.Bibliography` has no dependency on another OfficeIMO package. It uses `System.Text.Json` for CSL JSON and has no dependency on `OfficeIMO.Word`, the Open XML SDK, a TeX runtime, EndNote, or a network client.
