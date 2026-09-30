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

The suite exposes missing AVIF pixels with PDF omission warnings; page floats
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
| Page and column floats | Float normalization accepts left/right, logical sides and footnote; top/bottom/snap/inside/outside are diagnosed fallbacks. No column-scoped float-reference layout path | Source-confirmed scope gap against the reference's page/column float contract. Existing page footnotes must not be counted as absent. |
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
