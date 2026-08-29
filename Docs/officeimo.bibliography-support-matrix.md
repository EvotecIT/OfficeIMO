# OfficeIMO.Bibliography support matrix

This matrix defines the current `OfficeIMO.Bibliography` data, preservation, and conversion contract. Citation rendering, remote metadata lookup, library synchronization, attachments, Word integration, scripting, DRM, and encrypted-resource handling are outside this package.

## Format lifecycle

| Format | Create | Read | Edit | Deterministic write | Reopen output | Exact unchanged source |
| --- | --- | --- | --- | --- | --- | --- |
| BibTeX | Yes | Yes | Yes | Yes | Yes | Text and loaded bytes |
| BibLaTeX | Yes | Yes | Yes | Yes | Yes | Text and loaded bytes |
| CSL JSON | Yes | Yes | Yes | Yes | Yes | Text and loaded bytes |
| RIS | Yes | Yes | Yes | Yes | Yes | Text and loaded bytes |
| NBIB/MEDLINE | Yes | Yes | Yes | Yes | Yes | Text and loaded bytes |
| EndNote XML | Yes | Yes | Yes | Yes | Yes | Text and loaded bytes |

`.bib` paths default to BibLaTeX because the extension cannot distinguish the two profiles. Callers can select `BibliographyFormat.BibTex` explicitly when classic BibTeX output is required.

## Typed model

| Capability | Contract |
| --- | --- |
| Identity | Stable citation key plus original native item type; deterministic generated keys avoid collisions with supplied keys |
| Item kind | Journal, magazine and newspaper articles; books and chapters; conference papers and proceedings; reports; theses; web pages; datasets; software; patents; legal cases; manuscripts; personal communications; generic documents |
| Contributors | Ordered personal or literal organization names with author, editor, translator, recipient, interviewer, composer, collection-editor, and other roles |
| Names | Given, family, literal, suffix, dropping-particle, and non-dropping-particle components |
| Dates | Ordered partial, literal, or ranged issued, accessed, submitted, original, event, and other dates |
| Identifiers | Ordered scheme/value pairs, including DOI, ISBN, ISSN, PMID, PMCID, and source-specific schemes |
| Publication data | Title, container and collection titles, publisher and place, edition, volume, issue, pages, abstract, language, and URL |
| Repeatable values | Contributors, dates, identifiers, keywords, notes, and native fields preserve order |
| Extensions | Unknown record, CSL name, and CSL date fields remain in their owning `NativeFields`; document directives remain in `BibliographyDocument.NativeEntries` where supported |

## Native parsing and canonical writing

| Format | Native behavior and known limits |
| --- | --- |
| BibTeX/BibLaTeX | Reads braced, quoted, numeric, identifier, nested-brace, and `#`-concatenated values. Retains unknown fields, `@string`, `@preamble`, `@comment`, and top-level `%` comments. Empty keyword entries survive canonical output. String expansion is bounded and non-executing. Structured name components containing BibTeX separators such as top-level commas or `and` are diagnosed as lossy before strict output. Canonical writing does not evaluate TeX macros. |
| CSL JSON | Reads a single item object or item array. Known scalar properties become typed values only when their JSON shape is a string; numeric, Boolean, null, object, and array shapes remain raw native JSON for exact same-format output. Unknown properties are likewise preserved as raw JSON values. JSON comments and trailing commas are accepted but normalized. CSL JSON is treated as an interchange shape; the external CSL schema remains the authority for a consumer's chosen profile. |
| RIS | Reads ordered tags, repeated tags, records delimited by `TY`/`ER`, and continuation lines. Unknown safe two-to-five-character tags survive same-format canonical output. |
| NBIB/MEDLINE | Reads four-character-style tags, repeated fields, continuation lines, blank-record boundaries, PMID identity, and bracket-qualified identifiers such as DOI. Unknown safe tags survive same-format canonical output. |
| EndNote XML | Reads record numbers and types, contributors, titles, dates, identifiers, publication fields, URLs, keywords, notes, and unknown record elements. Safe unknown root elements and direct record elements survive same-format canonical output; retained records-container extensions cannot introduce reserved direct `record` elements. Unsupported content nested inside a known container remains available as raw native XML and produces a loss diagnostic if a canonical edit cannot merge it safely. DTD processing and external resolution are prohibited. |

## Preservation and conversion

- Preserve mode returns the original source only when the destination format matches and the semantic model is unchanged.
- Loaded bytes are retained exactly. Parsed text is retained exactly and is encoded with the selected writer encoding when bytes are requested.
- Any edit switches output to deterministic canonical syntax. Untouched trivia inside a modified record is not position-preserved.
- Unknown native fields remain in the model. Same-format writers preserve safe extensions; cross-format writers report fields they cannot carry.
- `BibliographyConversionReport` records stable codes, severity, action, item key, and field. `HasLoss` covers approximated, omitted, and error decisions.
- `BibliographyWriteOptions.RequireNoLoss` rejects a write before bytes are returned when conversion would lose data.

## Input and security limits

| Control | Default |
| --- | ---: |
| Encoded bytes | 64 MiB |
| Decoded UTF-16 characters | 64 MiB |
| Items | 250,000 |
| Values across records | 2,000,000 |
| One decoded value | 4 MiB |
| Parser diagnostics | 10,000 |
| BibTeX/JSON/XML nesting | 128 |

Parsing and writing observe cancellation. The library does not execute BibTeX/TeX, fetch URLs, load XML external entities, invoke native citation tools, or inspect local files named by citation data.

## Fixture provenance

The checked-in codec fixtures are repository-authored minimal interoperability records containing fictitious citation data. They do not copy third-party library exports or personal bibliography data. Contract tests reopen every canonical output through its owning codec.

## Integration boundary

`OfficeIMO.Bibliography` is format-neutral and has no Word or Open XML dependency. `OfficeIMO.Word` remains unchanged. The roadmap retains a future decision about an optional `OfficeIMO.Word.Bibliography` bridge that would translate Word bibliography sources to and from this model.
