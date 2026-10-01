# HTML shared-owner convergence and comparison evidence

This run qualifies source `40bd3f7bbbf0aff338336511ed899d27f79cd3fe` after integrating master `344dd83dc12e48d98fba933194f7773f84be212d`. The [compact record](qualification.json) binds test results, package inspection, reference versions, archive identity, output hashes and operation costs. The references are package/source pins, not a claim that every runtime dependency has been upgraded to its latest release. Open delivery work belongs in [the product roadmap](../../../../../Docs/ROADMAP.md#independent-html-engine).

## Integration result

The engine uses the current owned VP8 decoder and MIT Core package metadata. The obsolete prediction scratch implementation and CodeGlyphX license copy are removed; libvpx license and patent notices remain. Drawing uses the shared inspected caller-codec boundary once. PDF logical-cluster isolation preserves authored font tracking and its following text origin. Both system-font loaders prefer the exact family while retaining localized full/PostScript aliases. The runtime worker uses the current responsive-image selection contract, and its depth-limit test follows the shared error-as-failure contract.

An independent read-only review found two font-selection defects. Both reproduced in bounded fixtures, were fixed, and received one targeted confirmation. Full HTML tests pass 3,935/3,935, PDF 8,600/8,600 and Drawing 2,828/2,828 on .NET 8 and .NET 10. Runtime passes 646/646 on .NET 10. The HTML-to-PDF netstandard2.0 graph builds without warnings. Existing PDF test-source warnings remain outside this change.

H4 advanced held-out acceptance passes 8/8. Its unchanged macOS static budgets pass. All ten planetary PDF/editable operations pass their unchanged macOS ceilings; every operation still reports loss. This does not qualify Windows/Linux H10 costs or editable appearance.

## Frozen-page result

The planetary archive retains SHA-256 `305c6522e6c84a70ec1b9399725a1e7ed320762866abbcde99ff434f692e5ae2`. All eleven OfficeIMO PDF variants complete and preserve their previous page counts, extracted text-token multisets and image counts. Poppler extraction counts 3,292 word tokens per variant; that tokenizer is separate from earlier OfficeIMO extraction counts. Zero-margin print remains 16 pages with 18 color-image occurrences and six soft masks, or 24 image objects including masks.

PDF bytes change across integration. At 72 dpi, eight zero-margin pages are pixel-identical to the preceding checkpoint; eight have text-spacing/wrapping differences. Pages 9, 12 and 16 were inspected alongside the previous output. The table remains on OfficeIMO page 10 versus frozen Chromium page 9, and heading/footer differences remain open. The changed pixels are recorded, not accepted as browser equivalence.

PeachPDF **0.9.20**, isolated in the opt-in comparison project, now converts this archive successfully to 14 pages. The earlier 0.9.19 bookmark-geometry failure is historical and must not describe the current comparator. Fourteen pages do not establish better fidelity: resource, content and geometry criteria still need independent acceptance. The same input has distinct print, screen-media and snapshot contracts; PeachPDF's print lane does not measure OfficeIMO's other intents or editable exports.

## Capability assessment at the published reference boundary

### Static gap checkpoint

The [static-gap record](static-gap-checkpoint.json) binds the eighteen frozen inputs
to source `22157e58d`, published PeachPDF 0.9.20 and Chromium 151. All cases
finish within the predeclared 30-second process cap. Their PDF text and page
rasters use independent Poppler tools, and pypdf inspects image/link/structure
objects. These are observations, not complete visual or conformance acceptance.

Core's separate-alpha WebP decoder is merged through PR #2657. The HTML branch
also preserves its inspected caller-codec boundary through raster and SVG paths;
placeholder pixels remain a reported failure representation. Sixteen raw/compressed
alpha image occurrences extracted from the PDFs retain exact alpha and RGB within
one channel value of the independent reference. Drawing passes 2,850 tests on
each runtime, HTML 3,935 on each, PDF 8,600 on .NET 10, and Core netstandard2.0
builds without warnings. H4 passes 8/8 and both unchanged macOS budget gates pass.
All eleven planetary PDF variants are byte-identical to source `40bd3f7bb`.

The frozen baseline exposes missing AVIF pixels with PDF omission warnings; page floats
remain inline; column cases lose content and encounter ancillary Drawing bounds
failures; SVG filter graphs and a symbol without a viewBox lose paint. The selected
pattern output visually matches the browser, while OfficeIMO's own PDF reader hits
its clipping-work ceiling on that output. Formula/associated-file semantics and
requested PDF/A/PDF/UA outputs still need their separate policy lanes. PeachPDF
also omits these AVIF fixtures and the rotated pattern, and the browser does not
implement the declared paged floats/notes. None of those reference limits closes
an adopted OfficeIMO gap. Open outcomes remain in the single product roadmap.

The comparison uses the published [v0.9.20 HTML/CSS](https://github.com/jhaygood86/PeachPDF/blob/v0.9.20/docs/html-css-support.md), [SVG](https://github.com/jhaygood86/PeachPDF/blob/v0.9.20/docs/supported-svg-features.md) and [MathML](https://github.com/jhaygood86/PeachPDF/blob/v0.9.20/docs/supported-mathml-features.md) contracts. Their hashes are retained in the record. A documented reference capability is not a paired-fixture pass.

| Area | OfficeIMO evidence | Comparison conclusion |
| --- | --- | --- |
| Unfamiliar-page flow, responsive sizing and captured resources | Named WAI, MDN, NASA, NIST, NOAA, Yellowstone, EPA and FWS limits remain in the roadmap; H4 qualifies a selected corpus | Highest-priority outcome gap. Qualify appearance, content, links and losses across frozen pages rather than treating equal page counts as acceptance. |
| Default raster decode | Owned opaque VP8, raw/lossless-compressed ALPH, and VP8L. Animated WebP uses a bounded caller codec. The managed decode graph has no AVIF path | Sixteen independent alpha-WebP PDF occurrences pass decoded pixel checks at the recorded checkpoint. AVIF pixels remain absent with omission warnings. Keep codec policy in Core. |
| Page and column floats | The component checkpoints qualify ordinary left/right column floats and bounded horizontal page-edge boxes. Column top/bottom placement and column-scoped notes remain open | Whole-box page-edge placement passes content, edge and no-overlap checks. The qualified subset and fallback limits belong in the generated support matrix; reference failures do not close remaining gaps. |
| Automatic language hyphenation | Manual soft hyphens, limits and host-supplied callback/dictionary work. No shipped language-pattern catalog | Default-provider gap against shipped pattern-based languages. Qualify language fallback, dictionary provenance and mixed-language input through the shared text owner. |
| SVG paint and resource coverage | Native geometry, text, selected gradients/clips/masks/effects already exist; the support matrix identifies bounded fallbacks | Partial coverage, not an absent SVG engine. Establish paired cases for patterns, filter graphs, references and resource-dependent foreignObject before choosing additions. |
| Accessibility and archival output | HTML interactive AcroForms, semantic tags and bounded vector MathML already exist; PdfOptions exposes the PDF owner's policy | Qualification and adapter gaps: validate requested PDF/A/PDF/UA artifacts independently; compare Formula semantics and original MathML associated-file mapping. Do not label the PDF core or forms unsupported. |
| Editable targets and scripted pages | Six editable adapters and optional managed/browser runtime paths have separate contracts | Additional OfficeIMO scope. These are not PeachPDF static-print parity criteria and retain their own loss and execution gates. |

The architecture remains coherent: shared document/layout/drawing/PDF owners and retained parser/script providers serve the conversion product. This assessment does not justify renderer replacement, Skia adoption or dependency retirement. The next scope is bounded static HTML-to-PDF capability and unfamiliar-page qualification, with browser print as an independent oracle.

## Reproduction and retained artifacts

Rebuild the comparison, editable, static-budget and H10-budget tools in Release from the recorded clean head. Run `html-corpus-evidence --corpus advanced-held-out --verify-acceptance --require-clean-source`, `html-mhtml-evidence --mhtml <frozen-archive> --require-clean-source`, the static budget with three iterations, and H10 case `nasa-solar-system-terrestrial-planets` with the recorded archive hash and unchanged case ceiling. The tool READMEs define complete commands and output contracts.

Current ignored artifacts remain under `Ignore/HtmlUnknownPageQualification/`: `h10-nasa-master-clean-40bd3f7bb`, `h10-h4-master-clean-40bd3f7bb`, `h10-static-budget-master-clean-40bd3f7bb`, and `h10-budget-planets-master-clean-40bd3f7bb`. Source archives and independent browser references remain separate from superseded generated outputs.

## SVG symbol component qualification

At clean source `b040f7d85`, local symbols without a `viewBox` keep user coordinates and viewport clipping; both symbol forms honor nonzero containing viewport origins. The [component record](svg-symbol-checkpoint.json) retains the independently sampled browser pixels, PDF geometry, review remediation, current tests and unchanged budget results. The frozen symbol paints 2,000 green pixels at `[44,92,94,132]` in all three PDF lanes, retains the end marker once and preserves the link. The cyclic-gradient diagnostic remains; this does not qualify every SVG resource behavior.

Drawing passes 2,859 tests on .NET 8 and .NET 10, HTML passes 3,935 on .NET 10, Core netstandard2.0 builds without warnings, H4 passes 8/8, and all eleven NASA PDF variants remain byte-identical to `22157e58d`. Both unchanged macOS budget gates pass. Windows/Linux operation-budget qualification remains open. Caller-raster safety conservatively requires explicit no-viewBox symbol dimensions when the containing nested viewport is unresolved; native import resolves those defaults. The single roadmap owns that remaining safety-context extension.

The [viewport browser fixture](svg-symbol-browser/browser.html) and [origin fixture](svg-symbol-browser/origin-browser.html) reproduce the independent pixel oracle when served on loopback. Current complete outputs are under ignored `Ignore/HtmlUnknownPageQualification/static-gap-symbol-*`. Superseded passing test logs were compacted into their test summary; before-fix failures and final gate logs remain.

Superseded `22157e58d` OfficeIMO NASA PDFs were removed after verifying all eleven are byte-identical to the retained `b040f7d85` PDFs; the baseline archive, reports and compact hashes remain. This removes about 81 MiB of duplicate PDF output in addition to about 22 MiB of superseded passing test logs.

## Float flow inside columns

At clean source `0093a146c`, ordinary left/right floats share their exclusions with following paragraphs inside a multi-column container. Float paint stays whole at a column boundary, and only legal content breaks can split adjacent text. The [component record](column-float-checkpoint.json) binds the frozen input, independent PDF checks, review fixes and unchanged ceilings.

The `column-float-left` case prints in two columns at normal size. All 172 extracted word tokens match Chromium's multiset; all 45 unique markers appear once inside the page, and the link remains bounded. The 120×96-pixel float rectangle has the same position in all three PDF rasters. OfficeIMO matches Chromium's column assignment and wrapping around the float. PeachPDF retains the text but overlaps the end paragraph with a preceding paragraph. OfficeIMO's font-ligature warning remains, so these results do not claim full pixel parity or qualify page-edge/top/bottom floats or column-scoped notes.

Full HTML tests pass 3,942/3,942 on .NET 8 and .NET 10; the HTML-to-PDF netstandard2.0 graph builds without warnings. The independent review's arbitrary-text-cut and quadratic-filter findings were reproduced or source-confirmed, fixed and targeted-confirmed. H4 passes 8/8. All eleven NASA PDF variants remain byte-identical to the accepted `b040f7d85` output, and both unchanged macOS budget gates pass. Every H10 operation still reports loss; Windows/Linux budgets remain open.

Current complete outputs are under ignored `Ignore/HtmlUnknownPageQualification/static-gap-column-*`. Eleven superseded `b040f7d85` NASA PDFs were removed after verifying equality with these retained PDFs; their archive, reports and compact hashes remain. Preliminary column output and obsolete test runs were also compacted, removing about 95 MiB in total. Before-fix failures and final gate evidence remain.

## Page-edge float component qualification

At clean source `6d2fd539e`, `float-reference:page` places whole horizontal-writing boxes at the top or bottom content edge and reserves body space without advancing the source anchor. The [component record](page-float-checkpoint.json) binds the three frozen inputs, independent PDF extraction and geometry checks, review remediation and unchanged ceilings. Bare `float:snap` is a nearest-edge compatibility value. Nested, full-page, oversized and nonhorizontal page-edge floats keep their diagnosed normal-flow fallback.

All three cases preserve the 203-word token multiset against both references, print every unique marker once inside the page, place one float box at the declared edge, and keep body text clear. All six OfficeIMO page rasters were inspected. Chromium uses normal flow for these values; PeachPDF's bottom case overlaps two body paragraphs. Existing font-ligature warnings remain, so this qualification does not claim full pixel parity.

Full HTML tests pass 3,961/3,961 on .NET 8 and .NET 10, with 149 focused clipping/stacking/footnote/page-float tests and three catalog checks on each runtime. The netstandard2.0 HTML/PDF graph and four evidence tools build without warnings. One independent full review and one targeted confirmation exposed terminal-anchor, smaller-page, footnote-coordination and retained-path-clip defects. Each is fixed with regression proof; the final narrow path-clip correction received direct before/after validation rather than another review pass.

H4 passes 8/8. All eleven NASA PDF variants are byte-identical to the accepted `0093a146c` output; zero-margin output remains 16 pages with 24 images. Both unchanged macOS budget gates pass, including all ten H10 operations. Every H10 operation still reports loss, and Windows/Linux operation budgets remain open. Column-edge floats and column-scoped notes remain in the product roadmap.

Current complete outputs are under ignored `Ignore/HtmlUnknownPageQualification/static-gap-page-float-{clean,h4-clean,nasa-clean,static-budget-clean,h10-budget-clean}-6d2fd539e`. The three-case record contains reference output hashes, page bounds, marker counts and overlap checks; the source archive remains separately retained for replay.

Superseded authored output, passing/probe test files, eleven duplicate `0093a146c` NASA PDFs and ten old budget operation folders were removed after containment, provenance and activity checks. Their compact results, prior reports and source archive remain; current raw proof is retained at the clean-source paths above. This cleanup removes about 168.7 MiB.

## AVIF container and sequence-header stage

At source `eea5c578a`, Core has an internal bounded reader for whole in-file AV1 still items and their optional auxiliary alpha. It preserves CICP metadata and reads Main 8-bit reduced-still sequence headers without allocating image planes. The [header-stage record](avif-header-checkpoint.json) binds the independent frozen assets, source hashes, input-limit checks and review evidence to this checkpoint. External/stitched items, transformations and unsupported codec paths are rejected.

The focused suite passes 17/17 on .NET 8 and .NET 10; Core builds for netstandard2.0 with zero warnings. An independent review found an identity-matrix conformance error. Both color and monochrome regressions fail before the correction and pass after it; one targeted confirmation closes that finding.

This stage does not decode pixels. Both frozen AVIF files still fail the public default raster decode, and the support catalog remains unchanged. Frame/tile decoding, actual frame dimensions, pixel comparison, alpha composition, caller-codec integration and HTML/PDF qualification remain open in the product roadmap. Prior H4, planetary and performance results qualify their recorded sources; they are not decoder evidence for this stage. Superseded probes and intermediate passing logs were removed, retaining the final two framework results and before-fix failure proof under ignored `Ignore/HtmlUnknownPageQualification/static-gap-avif-container-work-c3b02831b`.


## AV1 still-frame and tile-boundary stage

At source `2f890df39`, the internal Main-8 reduced-still parser reaches bounded entropy payloads. It checks actual image dimensions against the AVIF item, retains quantizer/segmentation/filter parameters, supports uniform and nonuniform tile layouts, and rejects malformed lengths, nonzero alignment and explicit tile ranges inside a combined frame OBU. A shared syntax cursor also retains the previous sequence-header tests.

The [frame-stage record](avif-frame-checkpoint.json) binds the source and final 29/29 tests on .NET 8 and .NET 10 to this checkpoint. Core builds for netstandard2.0 with zero warnings. Independent FFmpeg header traces agree for the frozen color and alpha items and a four-tile libavif image. Handwritten syntax fixtures additionally check nonuniform tile geometry, segment-controlled lossless syntax and super-resolution/restoration dimensions; they do not supply independent producer or pixel proof for those features. One independent read-only parser review found no actionable defects.

Entropy validity, pixel reconstruction, applied grain and separate frame/tile OBUs remain unqualified. The public default decode path and support catalog are unchanged; this stage does not close the AVIF preservation gap. Keep the original frozen HTML/PDF acceptance and completed-decoder H4/NASA/performance gates open. Final results, native traces, tile payloads and the specification text needed for reconstruction remain in ignored `Ignore/HtmlUnknownPageQualification/static-gap-avif-frame-work-6e23fbd4c`; superseded logs, the duplicate tracked image and PDF copy were removed.

## AV1 arithmetic symbol stage

At source `2cdf55141`, Core's internal arithmetic reader decodes rising Q15 CDFs, adapts probabilities, reads pseudo-raw bits and literals, and checks tile termination. Encoded slices cannot borrow neighboring bytes; the caller supplies a finite symbol budget, and malformed termination, excessive synthetic lookahead and cancellation stop decoding. The [component record](avif-entropy-checkpoint.json) binds this implementation and its independent reference to the checkpoint.

The focused AVIF suite passes 47/47 on .NET 8 and .NET 10, and Core builds for netstandard2.0 with zero warnings. An isolated AOM v3.13.1 oracle generates and independently decodes twelve symbol/literal streams and four short boolean streams. OfficeIMO matches every symbol and final CDF, including disabled adaptation and duplicate probabilities. Three first-symbol prefixes from the frozen color, alpha and multi-tile AVIF slices also match native observations. A fresh hash-pinned download and build reproduces the checked-in fixture bytes. One independent read-only review found no actionable defects.

The opt-in reference generator is `OfficeIMO.Drawing.Tests/TestAssets/Avif/GenerateEntropyFixtures.py`; run it with `--work-dir` set to task-owned scratch. It requires Python's standard library and a C11 compiler, retains AOM's license/patent notices with the downloaded sources, and never runs during normal builds. No reference implementation or external dependency enters OfficeIMO's runtime.

Complete tile syntax, context selection and pixel reconstruction remain open, as do public raster integration and the original HTML/PDF acceptance gates. The prefix checks do not prove complete real-tile termination or decoded pixels. Windows/net472 execution and Windows/Linux operation budgets remain unqualified. Final framework results, the receipt and the small native oracle remain under ignored `Ignore/HtmlUnknownPageQualification/static-gap-av1-entropy-work-53fc23d2e` (about 252 KiB); superseded logs and duplicate reference-build output were removed.

## AV1 partition context stage

At source `34f6194e2`, the internal tile context selects adaptive partition CDFs from completed neighboring block dimensions. It handles temporary bottom/right edge probabilities, implicit edge splits and the decode order of all ten partition layouts. Each tile owns its probability arrays and bounded neighbor storage. The [component record](avif-partition-checkpoint.json) binds source hashes, reference inputs, final tests and review limits.

The focused AVIF suite passes 53/53 on .NET 8 and .NET 10; Core builds for netstandard2.0 with zero warnings. Forty component streams cover all five sizes and four contexts, with adaptation enabled and disabled, interleaving interior and clipped-edge reads. AOM supplies the pinned default tables and independent arithmetic encoder/decoder; an original harness implements the specification's partition-selection rules. The tests also cover rectangular neighbors, independent tile adaptation, implicit partitions, child order, geometry rejection and cancellation. These synthetic streams preseed neighbors and do not execute a complete recursive walk with leaf grammar.

Frozen native prefix observations and OfficeIMO agree that the color and alpha tiles split from 64×64 to a first 32×32 leaf, and the multi-tile image splits through 32×32 to a first 16×16 leaf. These observations stop before leaf modes and coefficients. One independent review found no actionable defects; an initially stale replay receipt was refreshed after the generator's notice-preservation edit. The current generator reproduces the fixture exactly.

Run `OfficeIMO.Drawing.Tests/TestAssets/Avif/GeneratePartitionFixtures.py` with `--work-dir` set to task-owned scratch to rebuild the opt-in oracle. Reference code and its notices remain outside the runtime graph. Final logs, receipts and the small native reference remain under ignored `Ignore/HtmlUnknownPageQualification/static-gap-av1-partition-work-9c5d5ec67` (about 359 KiB); obsolete logs, intermediate JSON and the duplicate replay build were removed.

Public AVIF decode and the support catalog are unchanged. Complete tile traversal, block modes, transform/coefficient syntax, reconstruction, filters, color/alpha composition and rendered HTML/PDF acceptance remain open. H4, NASA and budget results still qualify their prior recorded runtime paths, not a finished AVIF decoder. Windows/net472 execution and Windows/Linux operation budgets remain unqualified.

## AV1 leaf prelude stage

At source `d9e344379`, Core reads the leaf state preceding intra prediction: segment identity, skip, CDEF selection, quantizer delta and loop-filter deltas. Each tile owns its adaptive probabilities and neighbor storage. Superblock entry resets CDEF regions and the first-leaf delta flag while preserving running quantizer/filter state. Neighbor values are committed only after the remaining leaf syntax succeeds. The [component record](avif-prelude-checkpoint.json) binds the source, reference replay, tests and review limits.

Twenty-four native component streams cover 440 leaves with adaptation enabled and disabled, shifted independent tiles, rectangular leaves and full 64/128-pixel superblocks. The isolated AOM entropy implementation and pinned tables produce the reference bytes; an original harness implements the prelude syntax. A fresh rebuild reproduces those bytes. The frozen color, alpha and multi-tile first-leaf prefixes also match, including the color quantizer change from 32 to 18. The focused AVIF suites pass 58/58 on .NET 8 and .NET 10, Core builds for netstandard2.0 without warnings, and one independent read-only review found no actionable defects.

These streams do not qualify a complete recursive tile decoder. Positive smaller segment ranges, mixed recursive/A/B order and clipped frame-edge preludes need further replay evidence. Prediction modes, transforms/coefficients, reconstruction, filters, color/alpha composition and original HTML/PDF acceptance remain open. Public AVIF decode and the support catalog are unchanged. H4, NASA and performance evidence still belongs to the previously qualified runtime paths; Windows/net472 execution and Windows/Linux operation budgets remain unqualified.

Run `OfficeIMO.Drawing.Tests/TestAssets/Avif/GeneratePreludeFixtures.py` with a task-owned `--work-dir` to rebuild the opt-in oracle. Retained final tests, receipts and the small native reference are under ignored `Ignore/HtmlUnknownPageQualification/static-gap-av1-prelude-work-13ecbc4cc`; duplicate replay output and intermediate captures are removed.

## AV1 intra prediction mode stage

At source `bef583edd`, Core reads tile-local luma/chroma modes, directional angle deltas and chroma-from-luma signs and magnitudes. Small-block chroma availability, lossless eligibility and shared probability adaptation follow the Main-8 4:2:0 rules; monochrome leaves omit chroma. The optional intra-block-copy flag stops before motion syntax. Completed leaves publish neighbor modes through the same bounded tile geometry used by partition and prelude readers. The [component record](avif-mode-checkpoint.json) binds source hashes, native replay and final tests.

Twenty-four component streams cover 2,308 leaves and all 25 luma neighbor contexts. The pinned AOM entropy APIs and default probabilities supply an isolated reference; an original harness implements mode grammar. A fresh build reproduces the fixture exactly, and three frozen first-leaf prefixes match. Earlier entropy, partition and prelude fixtures remain byte-identical. The focused suites pass 62/62 on .NET 8 and .NET 10; Core builds for netstandard2.0 without warnings. One independent read-only review found no actionable defects.

Palette/filter-intra, motion, transforms/coefficients, complete mixed-size traversal and decoded pixels remain open. Public AVIF decode, support claims and rendered HTML/PDF qualification are unchanged. Prior H4/NASA and performance evidence qualifies its recorded runtime paths. Run `OfficeIMO.Drawing.Tests/TestAssets/Avif/GenerateModeFixtures.py` with a task-owned `--work-dir` for the opt-in reference. Final tests, receipts and the small native reference remain under ignored `Ignore/HtmlUnknownPageQualification/static-gap-av1-mode-work-d377d00b1`; duplicate replay builds and intermediate output are removed.

## AV1 palette and filter-intra stage

At source `f4bc2211c`, Core reads palette colors, filter-intra selection and padded palette index maps through the transform-size boundary. Color caches are tile-local; above color reuse resets every 64 pixels while above/left palette-presence probabilities remain available. Cached/new Y/U colors, raw and wrapped-delta V colors, small-block chroma rules and diagonal index order follow the Main-8 syntax. Completed leaves publish their palette neighbors only after the remaining leaf succeeds. The [component record](avif-palette-checkpoint.json) binds source, native replay, tests and review limits.

Forty-eight streams cover 2,484 leaves, every palette size from 2–8, all five filter modes and every reachable color-map context. AOM's pinned entropy APIs, default probabilities and native color-index context function supply the reference; the remaining grammar uses an original harness. A fresh rebuild reproduces the fixture exactly. The frozen color AVIF's first leaf matches four luma colors, four U/V colors and complete 32×32/16×16 index maps; the alpha and multi-tile prefixes have no palettes. Earlier component fixtures remain byte-identical. The focused suites pass 66/66 on .NET 8 and .NET 10; Core builds for netstandard2.0 without warnings. One independent read-only review found no actionable defects.

Intra-block-copy motion, transform/residual syntax, mixed-size completed traversal and decoded pixels remain unfinished. Public AVIF decode, support claims and original HTML/PDF acceptance are unchanged; completed-decoder H4/NASA/performance and missing platform proof remain open. Run `OfficeIMO.Drawing.Tests/TestAssets/Avif/GeneratePaletteFixtures.py` with a task-owned `--work-dir` for the opt-in reference. Final tests, receipts and the small native reference remain under ignored `Ignore/HtmlUnknownPageQualification/static-gap-av1-palette-work-daaeb66ad` (about 580 KiB); duplicate builds and intermediate captures are removed.

## AV1 transform-size and residual geometry stage

At source `9b1baa2cc`, Core selects luma transform sizes, reads bounded variable transform trees for typed intra-block-copy input, and orders Main-8 residual geometry across luma/chroma and 64-pixel chunks. The returned leaf grid is immutable. Tile border state retains completed transform sizes, preceding block dimensions and inter/skip flags; partial or failed leaves cannot publish it. The [component record](avif-transform-checkpoint.json) binds source hashes, native replay, framework tests and review limits.

The 150 component streams cover 7,526 leaves, all 22 block shapes, all 19 transform sizes, every depth/split probability context, recursive mixed-size neighbors, clipped frame edges and adaptation enabled/disabled. Pinned AOM entropy/defaults supply the reference bytes; an original normative grammar uses full-frame grids rather than the production border caches. A fresh rebuild reproduces every fixture, including earlier stages. Three frozen first-leaf prefixes reach the coefficient boundary: color/alpha select 32×32 luma transforms and the multi-tile item selects 16×16, with matching chroma geometry and residual order. The focused suites pass 70/70 on .NET 8 and .NET 10; Core builds for netstandard2.0 without warnings. One independent read-only review found no actionable defects.

This stage does not decode coefficients, transform types or pixels. Real intra-block-copy motion, completed real leaf traversal, tile termination, prediction/reconstruction, filters and composition remain unfinished. Public AVIF decoding/support and original HTML/PDF acceptance are unchanged. Completed-decoder H4/NASA/performance checks and Windows/net472 execution or Windows/Linux budgets remain unqualified. Run `OfficeIMO.Drawing.Tests/TestAssets/Avif/GenerateTransformFixtures.py` with a task-owned `--work-dir` for the opt-in reference. Final test results, receipts and the small native reference remain under ignored `Ignore/HtmlUnknownPageQualification/static-gap-av1-transform-work-c731f26db`; duplicate rebuilds, intermediate fixtures and first-run output are removed.

## AV1 coefficient component qualification

At source `2c7b1da30`, Core reads transform types, end-of-block positions, signed quantized levels and bounded Golomb extensions. Tile-owned probability tables use the frame's base quantizer; transform-type eligibility uses the segment-adjusted quantizer independently. Neighbor changes stay private until the leaf completes, while later transforms in that leaf see the tentative state. The [component record](avif-coefficient-checkpoint.json) binds source hashes, fixtures, native replay, tests, package inspection and review.

The focused AV1/AVIF suite passes 75/75 on .NET 10. Full Drawing tests pass 2,934/2,934 on .NET 8 and .NET 10; Core builds for netstandard2.0 without warnings. The local three-target Core package has no external runtime dependencies. One independent read-only review found no actionable defects.

The opt-in native oracle covers 660 streams, 16,900 leaves and 139,022 residual transforms, all nineteen sizes and sixteen types, adaptive updates on/off, segment quantizers, shifted tiles, clipped frame edges, skips and sign contexts. Forty-two normalized AOM scans independently check the owned scan generator. Frozen color, alpha and multi-tile inputs match their first native coefficient block; maximum and malformed Golomb probes protect termination and 20-bit masking. A fresh replay reproduces the compressed fixture, all twenty-one generated CDF files and all six earlier component fixtures exactly.

Run `OfficeIMO.Drawing.Tests/TestAssets/Avif/GenerateCoefficientFixtures.py --work-dir <task-scratch>` to rebuild the reference. `--core-tables-dir` explicitly regenerates probability facts for inspection. The normal build only consumes the checked-in compressed fixture; it never downloads or links native reference code. Pinned AOM source and license/patent notices stay with the opt-in oracle.

This is syntax proof. The original native grammar harness is not an independent full image decoder. Real tile traversal, motion, prediction, dequantization/inverse transforms, reconstruction/filters, color/alpha composition and HTML/PDF pixel acceptance remain open. Public AVIF decoding and the support catalog are unchanged. Prior H4/NASA/performance evidence retains its earlier source boundary; Windows/net472 execution and Windows/Linux H10 operation budgets remain unqualified.

## AV1 intra-block-copy motion qualification

At source `083248dab`, Core derives tile-local copy predictors from completed leaf sizes and displacements, reads integer motion differences, and validates the source footprint against tile boundaries and both decoded-region delay rules. Ordinary intra leaves consume no motion symbols. Incomplete, failed or canceled leaves cannot publish neighbors. The [component record](avif-copy-motion-checkpoint.json) binds source, reference fixtures, tests, package and review.

The opt-in native reference covers 258 streams, 135,288 leaves and 46,158 copy leaves, all twenty-two block shapes, 64/128-pixel superblocks, recursive partition order, mixed neighbor sizes, shifted/clipped tiles, color/monochrome and adaptation on/off. Eighty-eight rejection streams exercise all eleven classes, signs and integer-bit endpoints. Eighteen source/control probes distinguish linear and wavefront delays, tile edges, shared chroma footprints and exact accepted class 0/9/10 displacements. A fresh replay reproduces the compressed fixture exactly.

Full Drawing tests pass 2,939/2,939 on .NET 8 and .NET 10. Core builds for netstandard2.0 without warnings, and the inspected three-target Core package has no external runtime dependencies. One independent read-only review found no actionable defects. The final reference adds exact-class controls after review; production files remain identical to the reviewed candidate and broad validation uses the final fixture.

Run `OfficeIMO.Drawing.Tests/TestAssets/Avif/GenerateCopyFixtures.py --work-dir <task-scratch>` to rebuild the pinned AOM entropy/defaults reference. Its grammar is original harness code, not an independent full AV1 decoder. Complete real tile traversal, reconstructed samples, filtering/composition and direct HTML/PDF pixel acceptance remain unfinished. Public AVIF support is unchanged; earlier H4/NASA/performance evidence keeps its earlier source boundary. Windows/net472 execution and Windows/Linux operation budgets remain unqualified.

## AV1 coordinated tile and restoration qualification

At source `1b02c56ef`, Core traverses complete reduced-still Main-8 tiles in superblock/partition order, streams each leaf and its signed residuals, and publishes tile completion only after entropy termination. Each tile owns mutable contexts; failure, cancellation or consumer exceptions permanently retire it. The [checkpoint record](avif-tile-checkpoint.json) binds source, reference replay, tests, package and review.

Pinned AOM v3.13.1 supplies an independently decoded reference through an opt-in trace patch. Four frozen color/alpha/multi-tile items contain seven tiles, 367 leaves and 1,530 residual records; every managed syntax value and quantized coefficient matches the native trace. Separately, 192 native restoration streams cover 1,938 units, all sixteen SGR sets, Wiener taps, adaptive updates, monochrome/chroma, shifted edges and superresolution geometry. A fresh source checkout/build reproduces both compressed fixtures byte for byte. Normal builds never download or link this reference.

One independent read-only review found an encoded-input memory-accounting defect. A focused regression fails before the fix; whole-array accounting now rejects it before callbacks. One targeted confirmation found no new defects. Final focused tests pass 13/13 on .NET 10 and full Drawing suites pass 2,952/2,952 on .NET 8 and .NET 10. Core packs all three targets with no external runtime dependencies or native reference assets; netstandard2.0 builds without warnings.

The seven complete tiles contain no skipped/copy leaves or active restoration; those positive paths still have component evidence. Prediction, dequantization/inverse transforms, reconstructed pixels, filters/composition and direct HTML/PDF acceptance remain unfinished. Public AVIF support and the generated support catalog are unchanged. Prior H4/NASA/performance proof keeps its earlier source boundary; Windows/net472 execution and Windows/Linux operation budgets remain unqualified.

Run `OfficeIMO.Drawing.Tests/TestAssets/Avif/GenerateTileFixtures.py --work-dir <task-scratch>` to rebuild both references from the pinned official source. Final tests, compact replay/package facts, license/patent notices and tight native YUV/OBU samples remain under ignored `Ignore/HtmlUnknownPageQualification/static-gap-av1-tile-work-1f9528ac9` (about 8.7 MiB). Duplicate source/build copies, raw traces, intermediate test results and inspected packages are removed.

## AV1 residual reconstruction qualification

At source `f2a03e747`, Core's internal tile-owned residual context snapshots frame and segment quantization, dequantizes bounded signed coefficients, and applies integer DCT, ADST, identity or lossless Walsh-Hadamard transforms. Returned residuals have full transform dimensions and final flip orientation; prediction and sample clipping remain the reconstruction consumer's responsibility. The [checkpoint record](avif-residual-checkpoint.json) binds source hashes, numeric facts, native samples, tests, package inspection and review.

The pinned AOM v3.13.1 reference executes its actual C inverse kernels and records signed pre-clip residuals plus native clipped samples. All 749 cases and 156,816 samples match, covering every supported size and type, nine lossless cases, selected coefficient patterns, plane deltas, segment quantizers and matrix levels. All 100,320 normative matrix bytes and the 8-bit DC/AC lookup tables match independently normalized native facts. The dequantization harness implements the coefficient storage grammar over native factors; it is not a complete native image decode. A fresh source checkout/build and final dataset replay reproduce the checked-in fixture exactly.

Focused residual/tile/restoration tests pass 17/17 on .NET 10; full Drawing suites pass 2,956/2,956 on .NET 8 and .NET 10. Core packs netstandard2.0, .NET 8 and .NET 10 without warnings or external runtime dependencies. Inspection verifies the exact embedded matrix bytes in every packed assembly and excludes native reference assets. One independent read-only review found no actionable defects.

Run `OfficeIMO.Drawing.Tests/TestAssets/Avif/GenerateResidualFixtures.py --work-dir <task-scratch> --spec-text <pinned-av1-spec-text>` to rebuild the opt-in reference. The generator checks the archived specification hash; `--core-tables-dir` explicitly regenerates numeric facts. Normal builds consume the embedded facts and checked-in fixture without downloading or linking the oracle. Final tests, replay/package receipts, notices and one native residual reference executable remain in ignored `Ignore/HtmlUnknownPageQualification/static-gap-av1-residual-work-5d6ff1c49` (about 9.9 MiB). Superseded builds, raw outputs and inspected packages are removed.

Prediction, reconstructed full-image pixels, in-loop filters, color/alpha composition and the original HTML/PDF acceptance remain open. The reconstruction consumer must account for retained outputs. Public AVIF decoding and the support catalog are unchanged; prior H4/NASA/performance proof retains its earlier boundary. Windows/net472 execution and Windows/Linux operation budgets remain unqualified.
