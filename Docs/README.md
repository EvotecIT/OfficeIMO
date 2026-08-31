# OfficeIMO documentation

Use this index to find package guides, cross-package contracts, generated evidence, and open product work. Package READMEs are the best starting point for installation and public API examples; support matrices provide detailed coverage and known limits.

## Start here

- [Repository overview](../README.md) — package map, dependency model, supported formats, and platform coverage.
- [Migration guide](../MIGRATION.md) — version-to-version package, API, and behavior changes.
- [Examples](../OfficeIMO.Examples/README.md) — runnable examples across the package family.
- [Roadmap](ROADMAP.md) — the single backlog for open cross-package and product work.
- [GitHub Releases](https://github.com/EvotecIT/OfficeIMO/releases) — release notes, version history, and downloadable artifacts.
- [Website guide](officeimo.website.md) — local build and publication workflow.
- [Website search and answer-engine operations](officeimo.website-seo-geo-operations.md) — publication and discovery checks.
- [CI and test strategy](officeimo.ci-test-strategy.md) — test ownership and validation lanes.
- [Real-world and PDF quality corpus evidence](officeimo.real-world-corpus-evidence.md) — bounded external discovery, strict provenance-bound PDF scorecards, deterministic sampling, isolation, and interpretation limits.
- [Excel benchmark notes](officeimo.excel.benchmark-notes.md) — reproducible comparison-suite execution and provenance.
- [Image engine benchmarks](../OfficeIMO.Drawing.Benchmarks/README.md) — validated identification, decode, encode, resize, and placement-optimization workloads, with isolated opt-in library comparisons.

## Current product contracts

The package README is the primary usage guide for its public API. These repository-level documents add cross-package contracts or evidence that would be awkward to repeat in every package.

- [Provenance support matrix](officeimo.provenance-support-matrix.md) — exact carrier, package-owner, removal, verification, and resource-limit contracts.
- [Content safety support matrix](officeimo.content-safety-support-matrix.md) — concealed-content, Unicode, cleanup, and format-specific ingestion-security boundaries.
- [Text formatting support matrix](officeimo.text-formatting-support-matrix.md) — native font, decoration, script, casing, rendering, and compatibility contracts across formats.

### Office formats and interoperability

- [Security and protected-content capabilities](officeimo.security-capabilities.md)
- [Word and Excel interoperability](officeimo.word-excel-interoperability.md)
- [Word advanced editing and evidence contracts](officeimo.word-advanced-contracts.md)
- [Word template and mail-merge scenarios](officeimo.word-template-mail-merge-scenarios.md)
- [Word and HTML support](officeimo.word-html-support-matrix.md)
- [Legacy Word DOC compatibility](officeimo.word.legacy-doc-compatibility.md)
- [Excel large-workbook guidance](officeimo.excel.large-workbook-guidance.md)
- [Excel remote loading](officeimo.excel.remote-loading.md)
- [Legacy Excel XLS compatibility](officeimo.excel.legacy-xls-compatibility.md)
- [Apple iWork source-reader support](officeimo.iwork-support-matrix.md)
- [Reader package family](officeimo.reader.md)
- [Google Workspace package family](../OfficeIMO.GoogleWorkspace/README.md)
- [OpenDocument package family](../OfficeIMO.OpenDocument/README.md)
- [PowerPoint package guide](../OfficeIMO.PowerPoint/README.md)
- [Visio package guide](../OfficeIMO.Visio/README.md)

### Document and conversion formats

- [PDF current state](officeimo.pdf.current-state.md)
- [PDF conversion support](officeimo.pdf-conversion-support-matrix.md)
- [PDF reverse-conversion scorecard](pdf-reverse-conversion-scorecard.json) — executable 56-case producer/target matrix plus scanned, mixed, encrypted, and malformed stress evidence.
- [HTML rendering support](officeimo.html-support-matrix.md)
- [RTF support](officeimo.rtf-support-matrix.md)
- [Email and Outlook artifact support](officeimo.email-support-matrix.md)
- [Email performance evidence](officeimo.email-performance.md)
- [AsciiDoc support](officeimo.asciidoc-support-matrix.md)
- [LaTeX support](officeimo.latex-support-matrix.md)
- [Bibliography support](officeimo.bibliography-support-matrix.md)
- [Markdown compatibility](officeimo.markdown.compatibility-matrix.md)
- [Markdown extension authoring](officeimo.markdown.extension-authoring.md)
- [Markdown lossless round-trip design](officeimo.markdown.lossless-roundtrip-design.md)
- [OneNote current state](officeimo.onenote.current-state.md)

### Rendering and image export

- [Image export contract](officeimo.image-export.md)
- [Image export capability matrix](officeimo.image-export-capability-matrix.md)
- [Browser-local conversion performance and limits](officeimo.blazor-wasm-conversion-proof.md)

## Generated evidence

These reports are generated from repository catalogs, manifests, and test suites:

- `Compatibility/generated/` is produced by the compatibility generators described in [Compatibility/README.md](Compatibility/README.md).
- [Email, stores, and cloud acceptance](Compatibility/generated/email-cloud-acceptance.md) is generated from the machine-readable roadmap acceptance manifest.
- [HTML support](officeimo.html-support-matrix.md) is generated from the HTML capability and diagnostic catalogs.
- [PDF conversion support](officeimo.pdf-conversion-support-matrix.md) is generated from `pdf-conversion-scenarios.json`.
- [CommonMark inventory](officeimo.markdown.commonmark-inventory.md) and [GFM inventory](officeimo.markdown.gfm-inventory.md) are generated by the Markdown test suites.
- [Benchmark evidence](benchmarks/README.md) records reproducible performance runs and their provenance.
