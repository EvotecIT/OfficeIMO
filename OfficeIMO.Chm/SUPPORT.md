# CHM support and qualification

| Contract | Support | Boundary |
| --- | --- | --- |
| Archive container | ITSF 2/3; ITSP 1; PMGL listings and PMGI lookup chunks | Unknown containers, sections, transforms, and malformed or ambiguous directory paths fail explicitly. |
| Storage | Section zero and standard MSCompressed LZX | LZXC control versions 1/2, 32 KiB frames, reset-table version 2, 32 KiB–2 MiB power-of-two windows. Declared/padded expansion is checked before allocation. |
| Raw extraction | Every entry, including opaque internal streams | Exact decoded bytes are available in memory; there is no automatic filesystem extraction or CHM writer. |
| Book metadata | #SYSTEM 2/3 title, default topic, contents/index paths, compiler, locale | Unknown metadata records remain available in the raw stream. |
| Contents | Compiled #TOCIDX with topic/string/URL tables; HHC fallback | Empty compiled streams use the HTML fallback. Cycles, invalid references, depth, item/reference counts and aggregate text limits reject malformed navigation. |
| Keyword index | Compiled keyword BTree; HHK fallback | Hierarchy, multiple targets and See Also labels are exposed. The compiled full-text search database is retained as bytes, not executed. |
| Text decoding | Explicit override, BOM/meta declaration, then locale code page | Undeclared text uses the registered charset provider; unknown/unavailable locale encodings report a fallback. |
| Linked HTML / Markdown | Selected topic order, unique anchors, internal links, embedded images | HTML retains per-topic language/direction and inert container attributes. Reflow is semantic. Combined CSS shares one cascade; CSS ID selectors are not rewritten when anchors are prefixed. Markdown has its own presentation limits. |
| EPUB | Reflowable manuscript, contents hierarchy/fragments and embedded resources | Existing EPUB importer/writer checks apply. Keyword index, See Also and help-viewer behavior are not reconstructed. |
| PDF | Per-topic HTML rendering, page-boundary preservation, searchable text, archive resources | Renderer support governs CSS, fonts and images. Multi-topic output is untagged; cross-topic destinations and keyword index are omitted and reported. Single-topic output uses the renderer's tag policy. |
| Reader | Rich HTML projection, topic citations, archive-byte source hashes, tables, links and assets | Topic citations identify `archive.chm!/topic.html`; they are not physical page numbers. |
| Help application features | Raw preservation only | ActiveX, script execution, compiled search, external/merged books and Windows viewer window behavior are outside the document-conversion contract. |

## Evidence

The pinned Microsoft-compiled *Version Control with Subversion* fixture contains 181 HTML topics, compiled contents/index streams, and LZX storage. All 226 directory entries match independent CHMLib extraction by length and SHA-256, including opaque internal streams. The checked-in manifest and attribution live in [the test fixtures](../OfficeIMO.Chm.Tests/Fixtures/README.md).

Deterministic fixtures qualify uncompressed ITSF 2/3, HTML sitemap nesting, multiple keyword targets, See Also, charset precedence and locale decoding, case/escape handling, caller stream ownership, cancellation, and read/conversion limits. Mutated independently compiled metadata qualifies reset-table bounds, contents cycles and invalid topic references. Conversion tests reopen searchable PDF and EPUB, exercise linked Markdown and active-content omission, and validate Reader transport identity.

This evidence qualifies the described profile; it does not establish complete Windows HTML Help compatibility or pixel identity with Internet Explorer. Archive byte equivalence does not establish rendered-page equivalence. EPUB validation and PDF presentation depend on the output engines and the selected document. Inspect every operation's fidelity report before publishing converted material.
