# OfficeIMO.Bibliography

`OfficeIMO.Bibliography` is the citation-data owner for OfficeIMO. It provides one editable model and deterministic codecs for BibTeX, BibLaTeX, CSL JSON, RIS, PubMed NBIB/MEDLINE, and EndNote XML. Its managed CSL renderer accepts caller-supplied styles without an external citation engine or executable.

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

## Render citations with a local CSL style

```csharp
CslStyle style = CslStyle.Parse(styleXml);
var processor = new CslProcessor(read.Document, style,
    new CslRenderOptions { OutputFormat = CslOutputFormat.Html });

var citation = new CslCitation("citation-1");
citation.Items.Add(new CslCitationItem(item.Key) {
    Locator = "12–14",
    LocatorType = "page"
});
CslRenderResult rendered = processor.Render(new[] { citation });
string citationHtml = rendered.Citations[0].Content;
IReadOnlyList<CslRenderedEntry> entries = rendered.Bibliography;
var visibleEntries = entries.Where(entry => !entry.IsEmpty).ToArray();
```

The renderer supports plain text and escaped HTML, style macros, names and dates, sorting, locale terms, citation positions, disambiguation, and collapse rules. It accepts independent styles and resolves dependent styles through a caller-supplied `CslStyleLoadOptions.IndependentStyleResolver`. Style resolution performs no network or filesystem access.

Citation and bibliography snapshots retain a keyed entry when a style produces no text. `CslRenderedEntry.IsEmpty` identifies these entries in both output formats, including HTML entries with an empty wrapper. Hosts can omit them from display while retaining their keys and citation numbers. Text consisting of explicit style whitespace remains content.

CSL-JSON `shortTitle` and `journalAbbreviation` supply `title-short` and `container-title-short` when the corresponding canonical field is absent. Rendering, variable conditions, sorting, and name substitution use the same values. An explicit canonical field takes precedence, including an empty value, and writing retains the supplied field names. Put metadata in its CSL-JSON fields; the `note` field remains annotation text.

`CslRenderOptions` controls locale overrides, abbreviations, output and intermediate sizes, citation count, and rendering-operation budgets. Style XML is bounded and rejects DTDs, external entities, undefined macros, and recursive macro references. Embedded CSL locales have their own attribution and license notices in the package.

Set `CslCitationItem.AuthorOnly` for a narrative citation. It uses the style's first names expression, including its name formatting and substitutions; a numeric style without names uses default long author formatting. Citation-layout affixes are omitted when every item in a cluster is narrative. `AuthorOnly` and `SuppressAuthor` cannot both be enabled on one item. Initials retain supported input emphasis. Sorting and disambiguation share the rendering work budget.

Name substitutions suppress selected variables, including their short forms, as they render. Variable-presence and numeric conditions still inspect the source values after substitution suppresses repeated output. An empty candidate releases its variable, name-comparison, and alignment state before the next fallback. Bibliography author replacement applies to name and text fallbacks, including an empty replacement that hides repeated names. Contributor labels retain their position and formatting beside replacement names. Punctuation cleanup resolves collisions between fields, delimiters, and affixes while retaining punctuation inside each field and its HTML formatting. Locale quote rules move adjoining commas and periods across inline formatting while preserving emphasis and identifier links. Adjacent single and double input quotes retain their nested quotation levels, including when formatting divides the quote characters. Apostrophes remain distinct from closing quotation marks, unmatched quotes retain their input treatment, and punctuation does not move between display containers.

Use `form="count"` on `cs:name` to count the contributors selected by the style's abbreviation rules, including a retained last name and combined editor/translator lists. Abbreviation requires both effective `et-al-min` and `et-al-use-first` settings; inheritance, subsequent-citation settings and macro sort overrides can supply either value. Identical editor and translator lists combine when those are the two selected roles. Expressions selecting additional roles render each role independently, including when an earlier substitution suppresses one role. Missing contributors can use the same substitutions as ordinary names. A present list selected down to zero names produces `0`. Macro sort keys honor `names-min`, `names-use-first`, and `names-use-last` overrides. Name and title sorting ignores punctuation while preserving internal word spaces, numeric runs, and contributor boundaries; displayed punctuation remains visible.

Use `BibliographyName.DroppingParticle` and `NonDroppingParticle`, or the matching CSL-JSON fields, to control particle placement and sorting. The family name retains the supplied text, including particles that belong in its primary sort key. For example, `{"family":"Gogh","given":"Vincent","non-dropping-particle":"van"}` lets a style demote `van`, while `{"family":"de Gaulle","given":"Charles"}` keeps the full family name together. Quotation marks in name fields remain literal text. Supply an existing initial such as `Ts.` in `given` when a name uses a multiletter abbreviation.

Initialization retains emphasis around each compound given-name component, including lowercase continuations such as `Guo-ping`. The style's `initialize-with-hyphen` option controls the separator between initialized components. Standalone lowercase name particles retain their text, and `initialize="false"` retains full compound names.

Name comparison treats single-letter given-name initials such as `J.J.`, `J. J.`, and `J J` as equivalent. This prevents typographical spacing from triggering disambiguation and lets equivalent editor/translator lists share a role label. Full given names and corporate literal names retain their distinct identities; displayed names retain the style's formatting.

Rich-text fields retain supported emphasis beside HTML entities. Escaped tags remain literal text. Small-caps spans accept whitespace in their CSS declaration; the renderer emits only supported formatting and discards other input declarations and attributes.

Set `page-range-format` on the style to expand or abbreviate page ranges. The renderer supports `expanded`, `minimal`, `minimal-two`, `chicago`/`chicago-15`, and `chicago-16`, including matching page prefixes and Roman range delimiters. Endpoints with different digit widths retain all significant digits. Distinct prefixes or suffixes retain their identifier text. Without this option, `cs:text` preserves page-range text; citation locators follow their own range rules.

Page expansion resolves abbreviated endpoints before `cs:number` converts them to Roman or ordinal forms. Decimal page digits retain their original glyphs, including fullwidth and Persian digits.

`cs:number` transforms bare numeric units individually and preserves generated locale text. Contextual labels recognize Arabic and Roman numeral lists; page and volume counts use their count value. A dotted version such as `4.2` is one value; a list or range such as `4.2 & 5.3` is plural. Nonnumeric identifiers such as `ES-22-8` retain their hyphens. Ordinals use the accompanying noun's gender, including the selected locator type. Partial locale suffix sets retain their matching rules, and an explicitly empty long ordinal remains empty. Numeric classification, numeral extraction, separators, and page ranges use linear scans with cancellation checkpoints. Numeral and connector expansion observes `MaximumIntermediateCharacters`.

Name sort keys use the CSL priority of family, particles, given name, and suffix. Short-name macros omit dropping particles and suffixes; given-only names sort by their visible name. Institutional sort keys omit initial English articles, including those followed by Unicode whitespace. Displayed names retain their text and formatting.

Note disambiguation compares citation text across full, subsequent, and near-note forms. Subsequent name options and first-note variables participate even when the style has no position condition. Short notes receive the names or titles needed to identify a work, including when they collide with another work's full note. First-note references participate in that comparison, and overlapping ambiguity sets share one bibliography-ordered year-suffix assignment.

Conditional disambiguation selects the detail needed to distinguish works. A title can resolve a collision without adding an edition; works that still match can receive further detail. Locator-dependent branches retain the detail needed to identify a work when page numbers are added. Repeated macro calls and unique full notes avoid redundant additions. Nested and sibling conditions can contribute together, and explicit conditional year-suffix fields are reconsidered after suffix assignment. Trials share the rendering work limit and are recalculated for each document sequence.

Date-only citations can collapse repeated years and year-suffix ranges. Layouts without a names expression and expressions whose names are empty share the same visible-name group. Citation prefixes and suffixes keep affixed cites separate from adjacent collapse operations.

XML indentation in an empty locale term does not render. To make whitespace itself a term, set `xml:space="preserve"` on that term or an enclosing element. Nonempty term values retain their spaces, including text separated by XML comments or CDATA. Output attributes such as `prefix`, `suffix`, and `delimiter` retain their values.

Supply `citation-label` values as record data when a style uses them; the renderer does not generate labels from author names and years. Metadata written as field declarations inside an annotation `note` remains annotation text. Supply those values in their CSL fields, such as `reviewed-title` or `container-title`. Name conjunctions use the style's generic joining rules; use an explicit name delimiter when a script requires different spacing.

`Render` recalculates a complete document sequence. Pass the revised sequence after inserting, removing, replacing, or moving citations; successive operations do not retain the previous document's numbering or disambiguation state. Results contain the current citations and bibliography. Hosts decide which displayed regions to refresh.

Set `CslCitation.NoteIndex` to the footnote or endnote number, or zero for a body citation. Body and note citations keep separate position histories. Across notes, `ibid` requires consecutive note numbers and an unambiguous previous note containing one work; within a cluster, it follows the rendered item order. Empty and absent locators both mean no locator. First-note references remain empty in body citations and in the work's first note.

Note styles capitalize the opening of the first rendered citation in a note, preserving formatting and case-protected text. Set `NoteHasPrecedingText = true` when inserting a citation into existing note prose. An item prefix containing text also prevents automatic capitalization. Later citations in the same note retain their casing.

`rendered.BibliographyLayout` carries hanging-indent, line-spacing, entry-spacing, and second-field alignment settings for a document or HTML host. HTML entries use `csl-left-margin` and `csl-right-inline` blocks for automatic second-field alignment and retain explicit CSL display blocks. Layout, group, and macro-wrapper affixes stay with their own first and last fields; formatting preserves emphasis and links. Mixed inline fields and display blocks retain their order. Apply the returned settings in the host's layout system; `MaximumLeftMarginCharacters` counts the final visible UTF-16 text after layout transformations, so use actual text measurement when selecting a column width.

HTML bibliography entries link rendered `URL`, `DOI`, `PMID`, and `PMCID` values to absolute HTTP or HTTPS targets. URI prefixes stay inside the anchor; descriptive affixes stay outside. Set `LinkBibliographyIdentifiers = false` to disable links. Citation clusters and plain text output remain text, and rendering never fetches a link.

Title casing uses the versioned CSL English stop-word list, including phrases and hyphenated words. The item language controls Unicode casing and whether English title casing applies; without an item language, title casing follows the style's default language. Output-locale overrides control locale terms separately. Case-protected input retains its original text, and ordinal day formatting uses the month term's grammatical gender.

Inspect `processor.DataConversionReport` when rendering a bibliography imported from another format; set `RequireNoDataLoss` to reject lossy CSL data projection. This report describes data conversion, not full style conformance. The renderer implements a qualified subset of CSL 1.0.2 behavior; the [support matrix](../Docs/officeimo.bibliography-support-matrix.md) records its current limits.

## Boundaries

The package does not execute TeX, fetch DOI or PubMed metadata, resolve remote resources, manage attachments, remove DRM, or decrypt resources. Citation rendering uses local style and locale data supplied by the caller or embedded in the package.

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
