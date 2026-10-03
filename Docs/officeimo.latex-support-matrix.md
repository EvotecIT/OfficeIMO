# OfficeIMO.Latex support matrix

- Status: bounded-profile implementation
- Profiles: `LatexDocumentProfile.OfficeIMO`, `LatexDocumentProfile.PreserveOnly`
- Runtime dependencies: BCL and existing OfficeIMO project references only

`OfficeIMO.Latex` provides LaTeX document interoperability. It is not a TeX engine, package runtime, or promise that arbitrary `.tex` projects can be compiled.

## Status meanings

- **Semantic:** source-backed, typed, editable where the API exposes an edit, and written losslessly.
- **Source-preserved:** retained exactly in tokens/syntax without executing its TeX meaning.
- **Converted:** mapped to a typed Markdown or Reader representation.
- **Fallback:** emitted visibly with a diagnostic when the target lacks an equivalent.
- **Unsupported:** deliberately outside the bounded profile.

## Tokenizer, parser, and writer

| Capability | Status | Current contract |
|---|---|---|
| Decoded character source and locations | Semantic | Complete source string with offset/line/column spans. |
| TeX-aware tokens | Semantic | Control words, control symbols, braces, brackets, comments, whitespace/line endings, math shifts, alignment tabs, parameter markers, superscript/subscript, and other special tokens. |
| Lossless nested syntax | Semantic | Commands with arguments, groups, environments, math, comments, and trivia cover the source exactly. |
| Profile-aware argument binding | Semantic | Known commands bind brace groups or single character/control-sequence tokens, including starred headings and optional `tabular` placement. `LatexArgument.IsSingleToken` identifies unbraced arguments; edited replacements are braced. Missing arguments are diagnosed and remain visible during conversion. Environment names require braces. Literal square brackets and standalone brace groups are not claimed by unrelated zero-argument commands. Unknown commands retain a source-preserving fallback binding policy. |
| Preserve writer | Semantic | Unchanged input is returned character-for-character. Non-overlapping semantic edits replace only their source spans. Conflicting edits to identical or overlapping spans are rejected; identical replacements to the same span are deduplicated. |
| Canonical writer | Semantic | Normalizes line endings when requested; canonical source generation is owned by the Markdown adapter. |
| Recovery | Semantic | Unclosed groups, unexpected closing braces, unclosed math, unclosed environments, and mismatched `\end` names are diagnosed without discarding source. |
| Budgets | Semantic | Encoded input bytes, decoded characters, token count, nesting, macro input/output length, aggregate expansion tokens, and expansion depth are bounded. Character limits apply before line-map allocation; substitution checks output limits before appending. Options and opaque environment names are snapshotted. |
| Original encoding and BOM | Unsupported | The engine's fidelity contract is decoded character source, not byte-for-byte file encoding. |
| TeX execution | Unsupported | Parsing never invokes TeX, a shell, an external process, or package code. |
| `PreserveOnly` profile | Semantic structure only | Tokens, syntax, commands, environments, and math remain available, while OfficeIMO headings, paragraphs, lists, figures, tables, references, theorems, footnotes, and macro semantics are deliberately not bound. |
| Named profile factory | Semantic | `LatexParseOptions.CreateProfile` rejects unknown profiles and keeps macro expansion disabled until the caller explicitly selects `SafeSimpleDefinitions`. |

## OfficeIMO document profile

| LaTeX feature | Native status | Markdown/Reader outcome | Current boundary |
|---|---|---|---|
| `article`, `report`, and `book` document classes | Semantic | Metadata/profile recognition | Other classes remain source-preserved and are reported as outside the recognized profile. |
| Package declarations | Source-backed commands | Preserved/diagnosed | A declaration never loads or activates package code. |
| Title, author, and date | Semantic | Converted | Common front-matter commands are typed. `\maketitle` is structural and not emitted as visible paragraph text. |
| Parts, chapters, sections, and subsections | Semantic | Converted | Common starred/unstarred headings map to the closest target level. |
| Paragraphs | Semantic | Converted | Source-backed paragraph spans within the active document body and outside structural environments; inert source after `\end{document}` remains native source. |
| Common inline formatting | Semantic command binding | Converted | Common bold, emphasis, monospace, underline, URL/link, and line-break commands map through the inline adapter. Unknown commands remain visible or diagnosed. |
| `itemize`, `enumerate`, `description` | Semantic | Converted | Items and optional labels are typed and editable. Item bodies and description definitions retain ordered child blocks, multiple paragraphs, nested lists and quotations. Custom unordered or ordered labels remain visible in item text and report the target marker simplification; description labels become definition terms. |
| `footnote` | Semantic | Converted | `LatexDocument.Footnotes` exposes editable bodies and existing optional marks. References and structured bodies map to typed Markdown footnotes, Reader source locations and PDF text. Explicit marks report unevaluated counter state. Nested notes retain a source fallback; separate `footnotemark`/`footnotetext`, insertion placement and TeX counter execution remain outside this profile. |
| Figures and `includegraphics` | Semantic | Converted | Native target, option text, caption, and label are retained. Conversion decodes literal paths and reports unevaluated graphics options. No resource is loaded by the native engine. |
| `table` and `tabular` | Semantic | Converted | Alignment preamble, rows, empty cell positions, captions, and labels are carried where representable. A headerless table receives a visible blank Markdown header and a simplification diagnostic. Reverse conversion generates logical column counts and valid common `multicolumn`/`multirow` nesting. Unrepresented container source remains visible and diagnosed. |
| Labels and references | Semantic | Converted | `label`, `ref`, `pageref`, `eqref`, and common hyperlink forms are typed where recognized. Conversion reports counter, page-number, prefix, and equation-format simplifications; no counter engine computes page numbers. |
| Citations | Semantic | Converted with diagnostics | Common `cite`, `citep`, `citet`, and `nocite` keys are typed. Bibliography formatting is not executed. |
| Theorem-like environments | Semantic | Converted | Theorem, lemma, proposition, corollary, definition, remark, and proof environments map to named callouts. Reverse conversion emits the required `amsthm` package and deterministic `newtheorem` declarations for generated non-proof environments. |
| Verbatim and quotation environments | Source-preserved/semantic conversion | Converted | Verbatim is carried as code; quotes retain ordered paragraph and nested block children. Configured opaque variants become code. Line comments and comment environments are suppressed in projected text, including quotations. Percent comments consume their line ending while a following blank line still ends the paragraph. |
| Unknown commands and environments | Source-preserved | Visible fallback with diagnostics | They remain in the lossless tree and never disappear merely because the profile does not understand them. |

## Mathematics

| Feature | Status | Current contract |
|---|---|---|
| `$...$` and `\(...\)` | Semantic | Typed inline math retaining original delimiter and content. |
| `$$...$$`, `\[...\]`, and common math environments | Semantic | Typed display math retaining source and environment/delimiter metadata. |
| Labels around equations | Semantic where recognized | Reference metadata is retained; no TeX counter/equation-number calculation. |
| Markdown transport | Converted with diagnostics | Math uses Markdown's semantic fenced/code carrier so source remains visible and transportable. |
| Word/PDF transport | Simplified through Markdown | OfficeIMO semantic output; no claim of TeX math layout or full LaTeX-to-OMML conversion. |
| TeX math typesetting | Unsupported | The engine does not select fonts, shape glyphs, break formulas, or reproduce TeX layout. |

## Safe simple macros

Macro expansion is opt-in through `LatexMacroExpansion.SafeSimpleDefinitions`. The default is preservation without expansion.

| Feature | Status | Current contract |
|---|---|---|
| `newcommand`, `renewcommand`, `providecommand` metadata | Semantic | Simple command name, parameter count, and replacement body are retained. Commands and environments inside definition bodies do not contribute active document semantics. |
| Argument-only expansion | Semantic, opt-in | A deliberately small subset substitutes `#1`-`#9` with depth, cycle, and output limits. Current edits and optional defaults are rebound before document expansion; braces shield brackets inside optional arguments. |
| Replacement commands | Allow-listed | A replacement may contain parameters, plain source, another transitively safe document-local simple macro, or an explicitly allowed formatting/reference command. Every other control word—including definition, package, file, graphics, dynamic-control-sequence, and category-code primitives—is rejected for expansion. |
| `def`, `edef`, package macros, conditionals, counters | Source-preserved | No general TeX expansion or execution. |
| Document-controlled code/process execution | Unsupported | There is no shell escape, process invocation, package download, or assembly loading. |

The expander only performs bounded string substitution. Its allow-list does not sanitize invocation arguments and does not make the returned TeX safe to compile with an external TeX engine. Treat expanded output as untrusted source unless the caller applies its own compilation policy.

## Integration surfaces

| Surface | Status | Contract |
|---|---|---|
| `OfficeIMO.Latex` | Implemented | Native BCL-only tokenizer, parser, semantic profile, safe macro subset, and writer. |
| `OfficeIMO.Latex.Markdown` | Implemented both directions | Forward conversion reports simplification/fallbacks and retains residual container source. Reverse conversion emits escaped arguments, collision-free deterministic TeX-safe labels and bracket-safe optional arguments, declarations required by generated theorem environments, canonical bounded-profile LaTeX, and reparses it losslessly. |
| `OfficeIMO.Reader.Latex` | Implemented | Modular `.tex` path/stream handler with heading hierarchy, ordered whole-document projection, source spans, typed unordered/ordered/description lists, figures with captions, tables, math diagnostics, limits, and a diagnosed visible fallback for unrecognized plain TeX. Reader uses the canonical Markdown block projection for both text and Markdown, honors native byte limits and opaque environment names, and passes cancellation through loading and parsing. |
| Word, HTML, and PDF | Available through Markdown | Semantic OfficeIMO output, not TeX or package-renderer parity. |

Conversion uses the current edited document source. Native source spans retain the original parse offsets; conversion and Reader spans refer to the rebound source after edits. The `PreserveOnly` profile produces a diagnosed source fallback during conversion.

Reverse conversion preserves rich title content and labels, visible fragment-link text through `\hyperref[label]{text}`, and ordered loose-list child content. Custom numbering, task markers and ordinary callout layout report simplifications. Unrepresented image metadata reports omissions. Code closing-delimiter variants are escaped with a diagnostic. PDF destinations cover heading, theorem, table, figure and paragraph labels; unresolved links retain visible text and produce a warning.

## Qualification boundary

The native and conversion suites include unchanged `small2e.tex` and `sample2e.tex`
inputs from the official LaTeX2e source distribution. Their pinned revision, hashes
and redistribution terms are recorded in the [fixture provenance](../OfficeIMO.Latex.Tests/Fixtures/latex2e/README.md).
These independent inputs exercise source preservation, lists, quotations and footnotes.

Generated optional arguments, literals, labels, links, structured footnotes and
common table alignments are also qualified with an isolated external TeX compiler.
Independent PDF extraction and rendering check text order, single emission of note
bodies, font emphasis and real link destinations. Compilation remains a qualification
step outside shipped packages and normal installation. It does not establish arbitrary
package semantics, TeX page-layout equivalence or exhaustive independent-producer coverage.

Native .NET 8 and .NET 10 builds and focused LaTeX, conversion and Reader suites are
qualified on Windows, Linux/WSL x64 and macOS Arm64; Windows also runs .NET Framework
4.7.2. Public Reader file, stream and byte-array routes are compared on the pinned
official inputs. Linux uses an NTFS-mounted WSL checkout, so native ext4 checkout
behavior and other Reader formats are outside this qualification. The [model memory
evidence](benchmarks/officeimo.latex-model-memory-2026-10-03.md) separates parsing from
complete public-token inspection and records timing and process-memory limits.

## Deliberate limits

OfficeIMO does not compile `.tex`, resolve `input`/`include`, load CTAN packages, execute arbitrary macros, run BibTeX/Biber, calculate TeX counters/page references, build indexes or glossaries, execute code-generation environments, or reproduce TeX line breaking and page composition. Complex package-defined syntax remains lossless source and receives an explicit conversion fallback when it cannot be represented semantically.
