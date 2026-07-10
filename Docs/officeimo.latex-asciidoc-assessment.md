# LaTeX and AsciiDoc support assessment

- Status: design assessment with experimental Phases 0-3 implementation
- Date: 2026-07-10
- Constraint: no new third-party packages, runtimes, executables, or downloaded engines

## Decision

AsciiDoc is a sensible addition to OfficeIMO if it is implemented as a first-party document engine: a source-aware parser, a native syntax and semantic model, a writer, and explicit adapters. It fits the repository's existing Markdown and RTF direction and gives technical-document users a useful path to Word, Markdown, HTML, PDF, and `OfficeIMO.Reader`.

LaTeX is a conditional fit. OfficeIMO can credibly support a documented, source-preserving LaTeX document profile, including structured text, common environments, references, citations, and preserved mathematics. It should not claim to be a TeX distribution or to compile arbitrary `.tex` projects. Under the no-new-dependencies constraint, exact TeX expansion, the package ecosystem, and TeX-quality PDF typesetting are outside a sensible OfficeIMO scope.

The recommended sequence is:

1. Build AsciiDoc first and prove the native parser/writer architecture.
2. Strengthen the existing Markdown semantic bridge where AsciiDoc exposes real gaps, especially inline math, anchors, cross-references, admonitions, and conversion diagnostics.
3. Add LaTeX only behind a named support profile with explicit boundaries.
4. Do not run either format through `OfficeIMO.Markup`, regex substitutions, HTML, or flattened text as its primary model.

This is not a small pair of import/export switches. A credible AsciiDoc implementation is a medium-sized document engine. Credible LaTeX support is larger and must be deliberately bounded.

## Implementation status

Phases 0-3 are implemented on `codex/latex-asciidoc-design` as experimental, source-first engines and adapters:

- `OfficeIMO.AsciiDoc` provides the dependency-free parser, source map, lossless syntax tree, typed block and inline models, preserve/canonical writers, safe explicit preprocessing, tables, substitutions, and registered .NET directive extensions.
- `OfficeIMO.AsciiDoc.Markdown` converts in both directions. `OfficeIMO.Reader.AsciiDoc` registers `.adoc`, `.asciidoc`, and `.asc` with source locations, hierarchy, tables, Markdown projections, and warnings.
- `OfficeIMO.Latex` provides TeX-aware tokenization, a lossless nested syntax tree, recovery, the bounded `LatexDocumentProfile.OfficeIMO` semantic model, source-span edits, math transport, and opt-in safe-simple macro expansion.
- `OfficeIMO.Latex.Markdown` converts the bounded profile in both directions. `OfficeIMO.Reader.Latex` registers `.tex`, reports an unrecognized profile, and retains plain/unsupported TeX as visible source.
- Every production project has no `PackageReference`, regular-expression parser, process invocation, document-controlled code loading, or external runtime.
- The release-readiness review hardened profile-aware LaTeX argument binding, macro allow-listing, conversion fallbacks and canonical generation, variable-length AsciiDoc delimiters, include tag selection, logical table spans, and whole-document Reader projection. The new suites run on .NET 8, .NET 10, and Windows .NET Framework 4.7.2 in normal CI.

The exact implemented boundaries are tracked in the [AsciiDoc support matrix](officeimo.asciidoc-support-matrix.md) and [LaTeX support matrix](officeimo.latex-support-matrix.md). Neither matrix claims ecosystem parity.

Losslessness currently applies to decoded character source, including mixed line endings. The convenience file APIs do not yet retain the original file encoding or BOM bytes.

## What "no dependencies" means here

The proposed implementation may use:

- the .NET base class libraries;
- existing OfficeIMO projects and the packages they already require;
- new OfficeIMO-owned projects that contain the format engines and adapters.

It must not add:

- a new NuGet parser or renderer;
- Pandoc, Asciidoctor, Ruby, Java, Node.js, a browser, or a helper service;
- a TeX distribution, Tectonic, `latex`, `pdflatex`, `xelatex`, or `lualatex`;
- runtime downloads of syntax definitions, packages, fonts, or conversion engines;
- a hidden optional dependency that is required for normal behavior.

External implementations remain useful as research references and optional developer-side comparison tools. They must not participate in the product build, tests, runtime, or normal user workflow.

## Why these formats fit differently

### AsciiDoc

AsciiDoc is semantic markup for technical documents. Its document structure maps naturally to capabilities OfficeIMO already owns: headings, paragraphs, lists, tables, images, links, code, admonitions, metadata, and generated document outputs. It also has features that make a simplistic line parser incorrect: document attributes, ordered substitutions, block styles, macros, includes, conditionals, passthrough content, callouts, and extension points.

This makes AsciiDoc a strong product fit and a substantial parser project. The official language material and current pre-standardization specification are useful measures of the surface area, not implementation dependencies:

- [AsciiDoc language documentation](https://docs.asciidoctor.org/asciidoc/latest/)
- [include directive semantics](https://docs.asciidoctor.org/asciidoc/latest/directives/include/)
- [substitution groups and ordering](https://docs.asciidoctor.org/asciidoc/latest/subs/)
- [document attributes](https://docs.asciidoctor.org/asciidoc/latest/attributes/document-attributes/)
- [AsciiDoc language pre-specification](https://docs.asciidoctor.org/asciidoc/latest/_exports/asciidoc-pre-spec.pdf)

### LaTeX

LaTeX is both a document language and a programmable macro system layered over TeX. Commands can change how later input is tokenized and interpreted. Packages can introduce new environments, commands, counters, references, bibliography behavior, and layout rules. A parser that recognizes backslashes and braces is not therefore a LaTeX implementation.

OfficeIMO is a good home for interoperable LaTeX document content, especially scientific and technical text moving between `.tex`, Word, Markdown, and OfficeIMO PDF. It is not a good home for another TeX engine or a reimplementation of the complete package ecosystem. The [LaTeX Project](https://www.latex-project.org/) and its [documentation](https://www.latex-project.org/help/documentation/) illustrate the size and evolving nature of that ecosystem.

The useful product is consequently "LaTeX document interoperability," not "compile every TeX project."

## Existing OfficeIMO architecture

OfficeIMO already has most of the surrounding shape needed for a good implementation:

- `OfficeIMO.Markdown` owns a dependency-free typed Markdown model, source spans, parsing, writing, and extension seams; `OfficeIMO.Markdown.Html` owns HTML rendering.
- `OfficeIMO.Word.Markdown` and `OfficeIMO.Rtf.Markdown` demonstrate direct format adapters and conversion diagnostics.
- `OfficeIMO.Markdown.Pdf` demonstrates semantic PDF generation without an external converter.
- `OfficeIMO.Reader` has a modular handler registry, input kinds, capabilities, path and stream entry points, and dedicated format adapter packages.
- `OfficeIMO.Word` already represents Office Math Markup Language equations, although its Markdown bridge does not yet provide a complete equation conversion path.

Two existing types must not be mistaken for a universal document model:

- `OfficeIMO.Markup` is an Office-oriented authoring surface. It intentionally flattens some content and is not a lossless interchange representation.
- `MarkdownDoc` is the native semantic model for Markdown. It is a useful conversion bridge, but native AsciiDoc or LaTeX source must not be forced into Markdown when that would discard format-specific structure.

The repository's own Markdown correctness work already establishes the right rules: syntax-aware parsing, source spans, a lossless parse representation, typed transforms, explicit writers and renderers, and no regex-first semantic engine. The new formats should follow those rules rather than create a second architectural style.

## Target architecture

Each source format needs two representations:

1. A lossless syntax tree that retains tokens, trivia, delimiters, line endings, comments, source spans, unknown constructs, and recoverable malformed input.
2. A typed semantic document that makes supported concepts convenient to inspect, edit, create, and convert.

The original syntax remains authoritative for exact round trips. The semantic model is authoritative for new or intentionally modified content.

```text
source text
    |
    v
lexer / block scanner
    |
    v
lossless syntax tree <----> native semantic document
    |                              |
    | exact/canonical writer       | explicit adapters + diagnostics
    v                              v
source text                  Markdown / Word / PDF / Reader
```

This provides three distinct promises:

- **Source fidelity:** unchanged documents can be emitted without losing unsupported or unknown syntax.
- **Semantic editing:** callers can use typed nodes instead of manipulating strings.
- **Honest conversion:** adapters report when target formats preserve, simplify, render, or omit a construct.

### Package ownership

The initial package shape should follow existing repository conventions and avoid introducing a speculative universal document-model package:

```text
OfficeIMO.AsciiDoc
  Native AsciiDoc lexer, parser, syntax tree, semantic model, writer
  BCL only

OfficeIMO.AsciiDoc.Markdown
  AsciiDoc <-> MarkdownDoc mapping and diagnostics
  References OfficeIMO.AsciiDoc and OfficeIMO.Markdown

OfficeIMO.Reader.AsciiDoc
  Thin Reader handler
  References OfficeIMO.Reader and OfficeIMO.AsciiDoc

OfficeIMO.Latex
  Native LaTeX tokenizer, parser, syntax tree, semantic model, writer
  BCL only

OfficeIMO.Latex.Markdown
  LaTeX <-> MarkdownDoc mapping and diagnostics
  References OfficeIMO.Latex and OfficeIMO.Markdown

OfficeIMO.Reader.Latex
  Thin Reader handler
  References OfficeIMO.Reader and OfficeIMO.Latex
```

Word and PDF should initially reuse the existing Markdown adapters after the shared Markdown model gains the missing semantic nodes. A direct `OfficeIMO.Word.AsciiDoc` or `OfficeIMO.Word.Latex` adapter should be added only when a measured fidelity requirement cannot pass through `MarkdownDoc`. That decision keeps the first implementation small enough to finish without turning Markdown into the native model.

If three or more native document engines later need the same non-Markdown semantic concepts, a small OfficeIMO-owned interchange model may become justified. It should be extracted from working adapters at that point, not designed speculatively before the first vertical slice.

### Shared parser infrastructure

AsciiDoc and LaTeX should copy the architectural lessons of `OfficeIMO.Markdown`, not make `OfficeIMO.Markdown` their parser dependency. Low-level code may be extracted into an existing shared location only after two implementations prove it is genuinely format-neutral. Likely candidates are:

- immutable source ranges and line maps;
- cursor and character-scanning primitives;
- newline and trivia preservation;
- bounded nesting and input-budget helpers;
- diagnostic location formatting.

Block recognition, delimiter rules, substitution order, command parsing, environment parsing, and semantic binding remain format-owned.

## AsciiDoc design

### Processing pipeline

AsciiDoc needs explicit stages because attributes, directives, and substitutions affect later interpretation:

1. Decode text and build a line map while retaining the original encoding marker and newline style when known.
2. Scan document header, attributes, comments, block boundaries, continuation lines, and delimited regions.
3. Parse structural blocks into a lossless syntax tree.
4. Resolve configured preprocessing operations through a restricted resource context.
5. Apply the substitution group appropriate to each block, in a defined order.
6. Parse inline constructs and bind typed semantic nodes.
7. Run validation and collect diagnostics without discarding unknown constructs.

The public result should always contain the best recoverable document plus diagnostics. Strict mode may fail on errors; normal mode should preserve and continue.

### Core document coverage

The first production profile should include more than headings and paragraphs:

- document title, author, revision, and document attributes;
- sections with IDs, roles, options, and custom attributes;
- paragraphs, literal blocks, listing blocks, source blocks, examples, open blocks, sidebars, quotes, verses, and passthrough blocks;
- ordered, unordered, description, and checklist lists with nesting and continuation;
- tables with column specifications, spans, header/footer options, cell styles, and AsciiDoc cell content;
- inline emphasis, strong text, monospace, superscript, subscript, marks, constrained and unconstrained formatting;
- links, cross-references, anchors, footnotes, images, icons, keyboard/menu/button UI macros, and mail links;
- admonitions, callouts, block titles, captions, and generated lists;
- source-language and block-option metadata;
- attribute references, character replacements, special-character substitution, quotes, macros, post replacements, and explicit substitution overrides;
- conditionals and includes through safe, configurable resolvers;
- unknown block and inline macros preserved as typed raw nodes.

The API should expose an extension registry for block processors, block macros, inline macros, postprocessors, and syntax-to-semantic binders. Extension execution must be ordinary in-process .NET code registered by the caller; source documents must never load assemblies or scripts.

### Includes and resources

Includes are both a language feature and a trust boundary. Defaults should be safe:

- string and stream parsing disable includes unless a resolver is supplied;
- path parsing may use a file resolver rooted at the source document's directory;
- callers can restrict the root, extensions, maximum depth, total bytes, and number of resources;
- `..` traversal and symbolic-link escape are rejected after canonicalization;
- URL includes are unsupported by the built-in resolver;
- include cycles produce diagnostics and a preserved directive node;
- tag and line selection are applied only after the resource passes policy checks.

Conditionals and attribute evaluation need similar depth and expansion budgets.

### Writer behavior

The AsciiDoc writer needs two modes:

- `Preserve`: reuse original syntax for unchanged nodes and synthesize only modified subtrees.
- `Canonical`: emit stable, documented OfficeIMO formatting for generated or normalized documents.

Neither mode should silently replace an unknown macro or block with its visible text. A caller may request a lossy simplification policy explicitly.

## LaTeX design

### Supported product profile

The public contract should be named, for example, `LatexDocumentProfile.OfficeIMO`. It is a LaTeX2e-oriented interoperability profile, not a claim of complete TeX compatibility.

The initial profile should cover:

- comments, whitespace, groups, control words, control symbols, optional arguments, required arguments, and environments;
- document class and package declarations as metadata;
- title, author, date, abstract, parts, chapters, sections, subsections, and starred variants;
- paragraphs, line breaks, page breaks, quotations, verbatim, and common text formatting commands;
- itemize, enumerate, description, and nested lists;
- tabular and common table alignment/span commands within documented limits;
- figures, captions, labels, references, URLs, hyperlinks, and graphics resource references;
- footnotes, citations, bibliography resource declarations, and bibliography entries that can be represented without running a bibliography processor;
- inline and display math as losslessly preserved LaTeX math source with delimiter and environment metadata;
- common theorem-like environments as typed named blocks;
- counters and command definitions as preserved syntax and metadata;
- unknown commands and environments as source-preserving nodes.

Package declarations do not activate package code. A known-feature registry may bind selected package syntax, such as common `hyperref`, `graphicx`, `booktabs`, or `amsmath` constructs, when OfficeIMO implements that behavior itself. Unknown packages remain metadata and generate a diagnostic only when their constructs cannot be mapped semantically.

### Tokenization and macro boundaries

The tokenizer must preserve TeX-relevant distinctions: control words versus control symbols, comments that consume a line ending, groups, parameter markers, math shifts, alignment tabs, superscript/subscript tokens, and verbatim regions. It also needs recovery for unmatched groups and unterminated environments.

The semantic binder must not pretend to execute arbitrary TeX. The initial rules are:

- retain `\newcommand`, `\renewcommand`, `\def`, and related definitions in the syntax tree;
- recognize calls to statically known built-in commands;
- optionally support bounded expansion for a deliberately tiny safe subset of simple, argument-only definitions;
- never execute file I/O, shell escape, dynamic control-sequence construction, or package code;
- preserve unresolved invocations and report their conversion outcome.

This boundary makes parsing deterministic and safe, while still allowing unchanged source to round-trip.

### Mathematics

Mathematics should be treated as a first-class document concept, not flattened text. A math node should retain:

- original LaTeX source;
- inline or display mode;
- original delimiter or environment;
- optional equation label and numbering metadata;
- source span;
- an optional target-specific representation, such as OMML, when an adapter can create it.

The first version should preserve and transport math, not implement TeX math layout. Word conversion can initially use a source-preserving fallback or a deliberately supported LaTeX-to-OMML subset. PDF conversion can render a documented semantic/text fallback using `OfficeIMO.Pdf`; it cannot promise TeX-identical glyph shaping or layout.

### What LaTeX support will not do

Under this design, OfficeIMO does not:

- compile arbitrary `.tex` projects;
- download or interpret CTAN packages;
- execute TeX macros generally;
- run bibliography, index, glossary, or code-generation tools;
- reproduce a TeX engine's line breaking and page composition;
- accept shell escape or source-controlled process execution;
- claim compatibility based only on accepting the file without throwing.

These exclusions should appear in package documentation and API remarks, not only in an internal roadmap.

## Semantic conversion through Markdown

`MarkdownDoc` is the shortest route to the repository's current converters, but it needs several reusable additions before it can carry these formats honestly:

- typed inline and display math nodes;
- stable anchors and typed internal cross-references;
- block roles, options, and format-neutral attributes;
- admonition and named-container semantics;
- citation and bibliography-reference nodes;
- raw foreign block and inline nodes with source format, payload, and safety classification;
- resource references that distinguish embedded data from paths and unresolved logical targets;
- a shared conversion diagnostic shape.

Those additions benefit Markdown itself and `OfficeIMO.Word.Markdown`; they should not expose AsciiDoc- or LaTeX-specific syntax in otherwise generic APIs.

Every adapter reports an outcome for unsupported or transformed content. The implementation currently uses format-specific equivalents of this proposed shared shape:

```csharp
public enum DocumentConversionOutcome {
    Preserved,
    Converted,
    Simplified,
    VisualFallback,
    SourceFallback,
    Omitted
}
```

A diagnostic includes a stable code, source span where available, feature, outcome, and human-readable explanation. A format-neutral shared result and summary count remain a possible extraction after more adapters prove the common contract.

### Conversion promises

| Path | Intended promise |
|---|---|
| AsciiDoc -> AsciiDoc | Exact unchanged round trip; stable canonical generation |
| LaTeX -> LaTeX | Exact unchanged round trip within retained source; stable canonical generation for the supported profile |
| AsciiDoc/LaTeX -> Markdown | Semantic conversion with source fallbacks and diagnostics |
| Markdown -> AsciiDoc/LaTeX | Canonical generated source for representable Markdown semantics |
| AsciiDoc/LaTeX -> Word | Semantic conversion through the strengthened Markdown bridge initially; explicit math/resource diagnostics |
| Word -> AsciiDoc/LaTeX | Canonical source generation, not recovery of an original source document |
| AsciiDoc/LaTeX -> PDF | OfficeIMO semantic rendering, not Asciidoctor or TeX visual equivalence |
| AsciiDoc/LaTeX -> Reader | Structured extraction with headings, tables, code, links, and diagnostics |

## Public API

The implemented API separates parsing, optional processing, editing, writing, and conversion:

```csharp
AsciiDocParseResult result = AsciiDocDocument.Parse(
    source,
    new AsciiDocParseOptions {
        MaximumInlineNestingDepth = 64,
        MaximumInlineNodeCount = 1_000_000
    });

AsciiDocDocument document = result.Document;
document.BlocksOfType<AsciiDocHeading>().First().Title = "Updated title";

string updated = document.ToAsciiDoc(
    new AsciiDocWriterOptions { Mode = AsciiDocWriterMode.Preserve });

AsciiDocMarkdownConversionResult markdown = document.ToMarkdownDocument();

AsciiDocProcessingResult processed = AsciiDocProcessor.Process(
    source,
    new AsciiDocProcessorOptions {
        IncludeResolver = null // Includes stay disabled until a resolver is supplied.
    });
```

```csharp
LatexParseResult result = LatexDocument.Parse(
    source,
    new LatexParseOptions {
        Profile = LatexDocumentProfile.OfficeIMO,
        MacroExpansion = LatexMacroExpansion.SafeSimpleDefinitions,
        MaximumExpansionDepth = 16
    });

LatexDocument document = result.Document;
LatexMath math = document.Math.First();

string normalized = document.ToLatex(
    new LatexWriterOptions { Mode = LatexWriterMode.Canonical });
```

`Load` and `Save` delegate to the same parser and writer. Conversion extensions belong to adapter packages so the native engines remain dependency-free.

## Reader integration

`OfficeIMO.Reader` now has `ReaderInputKind.AsciiDoc` and `ReaderInputKind.Latex`. The format handlers live in separate modular packages and advertise precise capabilities.

Implemented chunk metadata includes:

- source format and source line range;
- section/heading path and block anchor;
- block kind and language for source blocks;
- conversion or recovery diagnostics affecting the chunk.

Per-cell table coordinates, explicit math-mode fields, and resolved-resource status remain potential Reader model extensions. Table, math, and resource content is currently carried in the chunk text/Markdown plus block kind and warnings.

Detection must be conservative. `.adoc`, `.asciidoc`, and `.asc` are reasonable explicit extensions. `.tex` is explicit but may contain plain TeX rather than LaTeX; the handler should diagnose an unrecognized profile instead of fabricating structure. Content sniffing must not steal generic text files based on a single heading-like line or backslash command.

## Testing and evidence

No external runtime should be needed to build or test either engine. Tests should use original, repository-owned fixtures derived from documented language behavior.

### Parser and writer contracts

- byte-for-byte or character-for-character unchanged round trips, including line endings and comments;
- syntax tree spans that reproduce the exact source slice;
- stable canonical output and parse-write-parse semantic equivalence;
- malformed input recovery with deterministic diagnostics;
- deep nesting, long delimiters, large tables, and adversarial expansion budgets;
- unknown command, macro, block, environment, and attribute preservation;
- multi-target build coverage matching the repository's supported frameworks.

### Feature matrices

Each engine needs a checked-in support matrix with one of these states:

- parsed and semantically editable;
- parsed and source-preserved;
- generated only;
- converted with a documented fallback;
- unsupported and diagnosed.

"Parsed" must not mean that a regex found the visible text. Each supported row should point to focused contract tests and, where relevant, Word/PDF/Reader artifact tests.

### Conversion tests

- semantic node-by-node tests for Markdown mappings;
- round trips where both formats share the same concept;
- diagnostics for every intentionally lossy branch;
- generated Word package inspection for headings, lists, tables, links, images, anchors, and equations;
- PDF text/layout assertions appropriate to OfficeIMO PDF, without pixel claims against external renderers;
- Reader chunk ordering, hierarchy, metadata, and source-span tests.

Developer-only comparisons against Asciidoctor, Pandoc, or a TeX engine may help investigate behavior. They must be optional, excluded from normal CI, and unable to change the supported contract silently.

## Delivery plan and gates

These phases describe the architectural gates used by the experimental implementation. Passing a phase gate means the bounded contract works and is tested; it does not mean full language or ecosystem compatibility.

### Phase 0: architecture proof, implemented

- implement source text, spans, diagnostics, and a minimal lossless tree in `OfficeIMO.AsciiDoc`;
- parse and preserve headings, paragraphs, lists, delimited blocks, attributes, and unknown macros;
- prove exact unchanged output and one edited subtree;
- convert a representative document to `MarkdownDoc`, Word, and Reader with diagnostics.

Gate: stop if the implementation starts flattening source, depends on regex ordering for structure, or requires a new external component.

The implementation passed this gate: source remains authoritative, parsing is scanner-based, and all production projects use only the BCL and existing OfficeIMO project references.

### Phase 1: credible AsciiDoc core, implemented experimentally

- implement the block/inline coverage, substitutions, attributes, tables, macros, and safe resource model described above;
- publish the support matrix and a meaningful fixture corpus;
- add Markdown, Word/PDF-through-Markdown, and Reader paths;
- document unsupported extensions and visual differences.

Gate: call it supported only when common technical documents retain structure and every unsupported construct is preserved or diagnosed.

The experimental implementation passes the bounded Phase 1 gate: block metadata, typed inline syntax, attributes, substitution plans, conditionals, root-confined includes, compound/description lists, admonitions, structured tables, STEM, registered directives, reverse Markdown generation, and Reader projections have focused contracts. The support matrix still marks ecosystem features that remain source-preserved or unsupported.

### Phase 2: LaTeX profile proof, implemented experimentally

- implement TeX-aware tokenization, lossless groups/commands/environments, and recovery;
- bind a small document profile and math nodes;
- prove exact unchanged output and explicit behavior for unknown commands and macro definitions;
- test a scientific article, technical report, and book-like document.

Gate: stop or keep the package experimental if real documents require general macro execution to expose basic structure.

The proof stays deliberately experimental. It retains complete source, recognizes article/report/book structure without general expansion, recovers malformed groups/math/environments, and exposes unknown commands as typed source-backed commands.

### Phase 3: bounded LaTeX interoperability, implemented experimentally

- complete the documented OfficeIMO profile;
- add tables, figures, citations, cross-references, theorem blocks, math transport, and safe bounded simple macros;
- add Markdown, Word/PDF-through-Markdown, and Reader paths;
- publish package-specific and command-specific support matrices.

The implementation includes the named semantics and paths, plus an explicit LaTeX support matrix. The bounded profile has completed focused regression, cross-target, solution, packaging, and repository review. It remains experimental because broader real-world corpus feedback may refine the named profile; full AsciiDoc ecosystem parity or arbitrary TeX execution remains a substantially larger, open-ended effort and is not the product goal.

## Risks

| Risk | Consequence | Design response |
|---|---|---|
| Marketing "LaTeX support" as full compatibility | Users expect arbitrary packages and TeX-identical PDF | Name and publish the OfficeIMO profile and exclusions |
| Treating Markdown as the native model | Attributes, macros, comments, math, and source fidelity disappear | Keep a native lossless tree and use Markdown only as an adapter target |
| Regex-first parsing | Incorrect nesting, delimiter, verbatim, and recovery behavior | Use stateful scanners, recursive structure, and explicit budgets |
| Include or macro expansion on untrusted input | File disclosure, denial of service, or process execution | Resolver capability model, safe defaults, no shell execution, hard limits |
| A generic interchange model created too early | Another incomplete "brain" competes with Markdown and Markup | Extract only after multiple working adapters prove shared concepts |
| Pairwise adapter growth | Duplicated mapping and inconsistent fidelity | Reuse strengthened Markdown adapters first; add direct adapters only from evidence |
| Source-preserving nodes hide conversion loss | A parse succeeds but target output silently changes meaning | Per-feature outcomes, diagnostics, and enforceable fidelity summaries |
| Scope expands to typesetting engines | Multi-year project unrelated to OfficeIMO's strongest capabilities | Keep PDF generation semantic and explicitly non-TeX-equivalent |

## Recommendation

Keep AsciiDoc as a first-party OfficeIMO engine. The implementation validates the architectural fit without new dependencies; the next decision should be driven by real documents and the remaining rows in its support matrix, not by a blanket compatibility label.

Keep LaTeX explicitly bounded to interoperability and validate demand for these workflows:

- importing scientific or technical `.tex` content into Word;
- generating maintainable `.tex` from OfficeIMO or Markdown;
- extracting structured `.tex` content through `OfficeIMO.Reader`;
- carrying equations, references, and citations between formats.

The experimental `OfficeIMO` profile now covers those paths at a bounded semantic level. If the actual requirement is "compile arbitrary LaTeX and reproduce its PDF," the answer under the no-dependencies constraint remains no: that does not make sense for this repository.

## Research references, not proposed dependencies

These projects were inspected only to understand feature scope and architectural trade-offs:

- [Pandoc's reader-AST-writer architecture](https://pandoc.org/using-the-pandoc-api.html) demonstrates why format-specific readers and writers benefit from an explicit semantic boundary. Pandoc itself is GPL-licensed and an external executable/library, so it is not part of this design.
- [Pandoc's supported formats](https://www.pandoc.org/demo/example2.html) and [release history](https://pandoc.org/releases.html) show that even a mature converter added its AsciiDoc reader comparatively recently. That is useful scope evidence, not a compatibility target.
- [Asciidoctor](https://docs.asciidoctor.org/asciidoctor/latest/) is the principal AsciiDoc implementation and a useful behavior reference. Its Ruby, JVM, and JavaScript runtimes are excluded.
- [Tectonic](https://github.com/tectonic-typesetting/tectonic) illustrates what a modern TeX engine entails. It and its runtime/package behavior are excluded.

Small .NET packages advertised as AsciiDoc or LaTeX parsers were also considered. Their presence does not remove the need for a complete OfficeIMO model, source fidelity, safe resource behavior, adapters, or explicit compatibility contracts. The dependency constraint makes the conclusion simpler: OfficeIMO either owns a credible implementation or does not advertise the feature.
