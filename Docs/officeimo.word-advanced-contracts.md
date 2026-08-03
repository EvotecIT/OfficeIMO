# OfficeIMO.Word advanced editing and evidence contracts

OfficeIMO.Word exposes advanced editing where the package can preserve Word meaning and reports the boundary where it cannot. These contracts are narrower than “anything Word can draw or lay out,” but they are executable and do not depend on silent approximation.

The boundaries below are current compatibility contracts, not omitted roadmap items. A profile is complete when the supported shape is editable and tested and every outside shape is preserved or rejected with a stable diagnostic. Broader arbitrary-topology or layout-engine claims require a new public contract and evidence before they become backlog work.

## Drawings, charts, groups, and SmartArt

- Imported charts support category replacement, series value replacement, and series-name replacement for the cached chart shapes OfficeIMO understands. Mutating a cache detaches that cache from worksheet references without deleting unrelated workbook content.
- `WordShapeGroup` creates and reloads bounded DrawingML groups containing two or more preset shapes. Child position, size, fill, and stroke remain editable data rather than a flattened picture.
- Imported SmartArt exposes logical node order derived from connection topology. `MoveNode` changes that structural order; node removal and clearing also remove incident connections and resequence the remaining topology.
- Charts, shapes, SmartArt, and shape groups expose `TryGetLayoutSnapshot`. The snapshot reports persisted inline/anchor placement, wrapping, extent, relative origins, offsets or alignment, z-order, and group state. It is package geometry, not Microsoft Word's final line-breaking or pagination result.
- New DrawingML shapes and groups allocate non-visual IDs above the current package maximum across body, headers, footers, notes, and comments. Reloading a producer document cannot reuse process-global IDs.

The opt-in `DrawingLayout_AnchoredGroupMatchesDesktopWordGeometryWhenRequested` test opens a generated anchored group in desktop Word and compares Word's reported position and extent with the package contract. Set `OFFICEIMO_RUN_WORD_LAYOUT_COM_VALIDATION=1` to run that oracle. It proves the documented group shape; it does not make COM a runtime dependency or generalize to arbitrary drawing topology.

## Structured comparison and redline output

`WordDocumentComparer.CompareStructure` can include the bounded Wordprocessing DrawingML shape/group scope through `WordComparisonOptions.CompareShapes`. It reports shape insertion, deletion, content mutation, layout mutation, and movement between body, header, footer, footnote, and endnote parts.

Report redlines include the shape evidence. In-place redlines use a tracked textual fallback and record `Redline.Shape.InPlaceTextFallback`; they do not claim a native drawing replacement or move revision. Effective formatting remains distinct from direct formatting, and detected relocation remains distinct from a native Word move revision until those contracts can be implemented without guessing.

## Field evaluation profiles

The managed updater implements native `LISTNUM` evaluation for Word's built-in `NumberDefault`, `OutlineDefault`, and `LegalDefault` profiles, levels 1 through 9, paragraph-level fallback, resets, and start overrides. Counters remain scoped by story part. Unknown custom list templates, nested instructions that cannot be evaluated safely, and overflow are reported rather than converted to a plausible-looking number.

Custom LISTNUM templates are document-defined numbering programs rather than another native built-in profile. They remain outside managed evaluation unless a typed profile contract is added; `FieldListNumberingProfileUnsupported` is the stable result, and existing field text is preserved.

Locale-sensitive currency pictures, layout-dependent page/index/table-of-contents results, complex tables whose result text would otherwise be concatenated incorrectly, and unsupported nested instructions have stable diagnostic reasons. Numeric format switches that are not fill directives are no longer stripped as if they were formatting noise.

## Evidence ownership

[`Word/EvidenceCorpus/corpus-manifest.json`](../OfficeIMO.TestAssets/Documents/Word/EvidenceCorpus/corpus-manifest.json) is the executable provenance index for review, redline, template, mail-merge, legacy DOC, Word/HTML, rendering, and performance evidence. Every file entry has a SHA-256, producer, oracle set, source test, and loss policy; generated entries name their deterministic generator and executable oracle.

The legacy DOC entry requires its approved import report and guarded loss policy. OfficeIMO can write the documented native DOC subset, but the corpus and public API do not claim arbitrary DOC authoring. Use `AssessLegacyDocWrite`, `WordDocument.Convert` reports, and an explicit `WordConversionLossPolicy` decision for legacy output.

For reciprocal HTML behavior, see the [Word/HTML support matrix](officeimo.word-html-support-matrix.md). For template boundaries, see the [template and mail-merge scenario matrix](officeimo.word-template-mail-merge-scenarios.md). For the native DOC subset, see [DOC and DOCX compatibility](officeimo.word.legacy-doc-compatibility.md).
