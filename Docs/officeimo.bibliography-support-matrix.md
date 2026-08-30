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
| Names | Given, family, literal, suffix, dropping-particle, and non-dropping-particle components; family-only multiword names and literal names ending in commas reopen without changing shape; surrounding whitespace and empty or entirely unset names that a non-CSL destination would normalize are diagnosed before strict output |
| Dates | Ordered partial, literal, or ranged issued, accessed, submitted, original, event, and other dates; incomplete numeric sequences, null-valued empty dates, and literal-date whitespace are diagnosed when a destination cannot preserve their distinction |
| Identifiers | Ordered scheme/value pairs, including DOI, ISBN, ISSN, PMID, PMCID, and source-specific schemes; destination-specific scheme spelling changes are diagnosed before strict output |
| Publication data | Title, container and collection titles, publisher and place, edition, volume, issue, pages, abstract, language, and URL |
| Repeatable values | Contributors, dates, identifiers, keywords, notes, and native fields preserve order |
| Extensions | Unknown record, CSL name, and CSL date fields remain in their owning `NativeFields`; document directives remain in `BibliographyDocument.NativeEntries` where supported |

## Native parsing and canonical writing

| Format | Native behavior and known limits |
| --- | --- |
| BibTeX/BibLaTeX | Reads braced, quoted, numeric, identifier, nested-brace, and `#`-concatenated values. Retains unknown fields, `@string`, `@preamble`, `@comment`, and top-level `%` comments. Native item-type, identifier-scheme field casing, unknown-field spelling, parsed empty keyword entries and positions, classic BibTeX month-only dates, and the relative order of supported BibLaTeX date roles survive canonical output, and malformed fields recover at the next safe field boundary with diagnostics. Diagnostic line and column locations recognize LF, CRLF, and CR-only sources. Retained duplicate typed fields are diagnosed if an edit would promote them into the typed model. Empty issued dates that classic BibTeX cannot express are diagnosed before strict output. Odd terminal backslash runs are normalized with a loss diagnostic so they cannot escape writer-added delimiters, and unsafe post-escape values are omitted rather than producing malformed records. Per-value and cumulative string expansion are bounded and non-executing. Structured name components containing BibTeX separators such as top-level commas or `and` are diagnosed as lossy before strict output. Canonical writing does not evaluate TeX macros. |
| CSL JSON | Reads a single item object or item array. Known scalar properties become typed values only when their JSON shape is a string; numeric, Boolean, null, object, and array shapes remain raw native JSON for exact same-format output. Recognized native type spelling, typed identifier value whitespace, blank identifier properties, empty recognized contributor arrays, and signed integer date parts survive same-format canonical edits. Incomplete numeric date sequences whose month, day, or range components lack their required owner are diagnosed before strict output. Blank keyword values remain distinct from an absent keyword list. Escaped string values are bounded by decoded UTF-16 length rather than their JSON escape syntax, limit diagnostics use zero-based UTF-16 source offsets, and token scans observe cancellation between bounded input segments. Invalid UTF-16 input is rejected instead of being replaced during UTF-8 conversion. Unknown properties are likewise preserved as raw JSON values. Retained duplicates are diagnosed if an edit would move them from native storage into a typed scalar, contributor, or date owner. JSON comments and trailing commas are accepted but normalized. CSL JSON is treated as an interchange shape; the external CSL schema remains the authority for a consumer's chosen profile. |
| RIS | Reads ordered tags, repeated tags, records delimited by `TY`/`ER`, and continuation lines. Recognized native type spelling, blank identifier tags, and issued or accessed date ranges survive same-format canonical edits. Blank source date literals remain distinct from absent dates; model-created null-valued empty dates are diagnosed before strict output. Nonempty `ER` payloads are diagnosed as recovered source loss. Single start-page and end-page tags retain their distinct roles through canonical edits. Retained repeated scalar tags are diagnosed if clearing their typed owner would promote the retained value. Explicitly empty unsupported model properties are diagnosed before strict output rather than being silently normalized to null. Per-value and accumulated continuation limits are enforced before materialization. Unknown safe two-to-five-character tags survive same-format canonical output. |
| NBIB/MEDLINE | Reads four-character-style tags, repeated fields, continuation lines, blank-record boundaries, PMID identity, and bracket-qualified identifiers such as DOI. Qualified identifier-scheme casing, blank `IS` serial-identifier tags, and issued date ranges survive same-format canonical edits. Continued publication types rebind the typed item from the complete value; recognized retained types that conflict with an edited typed kind are diagnosed before omission. Compact/full author pairing is cancellation-aware, Unicode-scalar-safe, and indexed by normalized compact name. Retained repeated scalar tags are diagnosed if clearing their typed owner would promote the retained value. Explicitly empty unsupported model properties are diagnosed before strict output rather than being silently normalized to null. Per-value and accumulated continuation limits are enforced before materialization. Unknown safe tags survive same-format canonical output. |
| EndNote XML | Reads record numbers and types, contributors, titles, dates, identifiers, publication fields, URLs, keywords, notes, and unknown record elements. Recognized native type aliases, accepted root and records-container element names, identifier-scheme casing, empty recognized containers, empty issued dates, distinct nonnumeric or empty year/publication-date components, blank notes, empty primary URLs, empty identifier elements, and redundant secondary/periodical title representations survive same-format canonical edits. Conflicting recognized type names and numeric type codes are diagnosed as recovered source loss before strict output. PMID scheme identity is not representable by the accession element and is diagnosed before strict output rather than silently narrowed. Safe unknown root elements and direct record elements survive same-format canonical output; retained root extensions cannot become structural `records` containers, and records-container extensions cannot introduce reserved direct `record` elements. Unsupported shapes inside known containers, such as an `author` directly beneath `contributors`, remain raw native XML and are preserved when no typed owner must share that container; conflicting canonical edits produce a loss diagnostic. Direct mixed text in structural root, records, or record elements is bounded before DOM materialization and diagnosed as recovered source loss. XML comment and processing-instruction values are bounded before DOM materialization. DTD processing and external resolution are prohibited. |

## Preservation and conversion

- Preserve mode returns the original source only when the destination format matches and the semantic model is unchanged.
- Loaded bytes are retained exactly. Parsed text is retained exactly and is encoded with the selected writer encoding when bytes are requested, except that unchanged EndNote XML text with an encoding declaration uses the declared encoding so the XML declaration and returned bytes remain consistent.
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

Stream decoding enforces the decoded-character limit incrementally before a full source string is materialized. BOM-less UTF-16 and UTF-32 EndNote XML is detected from leading XML markup with or without a declaration and after XML whitespace. EndNote line offsets use exact zero-based UTF-16 positions resolved by a constant-memory, cancellation-aware scan rather than a per-line index. Parsing, automatic content detection, baseline and preserve-mode fingerprinting, conversion inspection, and writing observe cancellation, including EndNote post-DOM projection, per-record contributor, identifier, keyword, note, native-field, and native-entry collections, final output encoding or preserved-byte cloning, and bounded synchronous or asynchronous stream writes. Save operations check cancellation before mutating a destination. The library does not execute BibTeX/TeX, fetch URLs, load XML external entities, invoke native citation tools, or inspect local files named by citation data.

## Fixture provenance

The checked-in codec fixtures are repository-authored minimal interoperability records containing fictitious citation data. They do not copy third-party library exports or personal bibliography data. Contract tests reopen every canonical output through its owning codec.

## Integration boundary

`OfficeIMO.Bibliography` is format-neutral and has no Word or Open XML dependency. `OfficeIMO.Word` remains unchanged. The roadmap retains a future decision about an optional `OfficeIMO.Word.Bibliography` bridge that would translate Word bibliography sources to and from this model.
