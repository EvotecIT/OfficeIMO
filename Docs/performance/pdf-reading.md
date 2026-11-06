# PDF browser-print reading evidence

## Workload and method

The 6 November 2026 run uses a 115-page Chromium/Skia PDF with 150 sections,
four paragraphs and an eight-row bordered table per section. The input has
2,655,493 bytes and SHA-256
`d9bc26d9d04f22032d5a07813358b4344e96518eecd12c98fdc00fa9c273d51c`.
The large PDF is an external reproduction input; the compact browser-column
fixtures live in `OfficeIMO.Pdf.Tests/Pdf/Fixtures/Interoperability/BrowserColumns`.

Measurements use Apple M4, macOS 27.0.1 ARM64, SDK 10.0.403 and Release builds.
The baseline is `186819ce958a928a0d3b13cd41e9d8f808d1481b`.
The measured candidate is `b2193d970c8a782bc655f6f36115913cc6b1e8be`.
For each profile, warm one fresh `PdfDocument.Load(path).Read(options)` and
Markdown extraction, then take three fresh loads/reads. Time `Load`,
`Read().Text` and `ExtractMarkdown()` separately; use the read median.
Structured uses default image/widget inclusion; Fast disables both.
The two builds run serially at default process priority and affinity. The shared
heavy-validation lock serializes cooperating work, but does not establish a quiet
host. Timing is opt-in evidence, outside correctness gates.

## Measured reads

Samples and medians below exclude Load and separate Markdown extraction.

| Profile | Baseline samples, ms | Candidate samples, ms | Median, before / after | Reduction |
| --- | --- | --- | --- | --- |
| Structured | 6490, 6251, 6858 | 5044, 4878, 4653 | 6490 / 4878 ms | 24.8% |
| Fast, images/widgets off | 4349, 3885, 5184 | 3610, 2879, 3985 | 4349 / 3610 ms | 17.0% |

The median per page falls from 56.4 to 42.4 ms for Structured and from
37.8 to 31.4 ms for Fast. These three-sample medians describe this run;
the Fast samples show substantial host variability.

Complete Text and Markdown exports are byte-identical to baseline for both
profiles. Structured retains 429,033 characters and 168 table pieces; Fast
retains 429,035 characters and 156 table pieces. Profile differences predate
this change. Markdown extraction takes 16–25 ms in the candidate samples.

## Investigation decisions

| Attempt | Evidence and disposition |
| --- | --- |
| Share one interpretation globally | Root text, resource, effect and visual passes repeat interpretation, but carry different transforms, callbacks, clipping and work budgets. A shared tokenization remains unqualified and is not adopted. |
| Reuse root resource evidence and skip empty invocation traversal | A performance-only control preserves all four complete large-file exports. Existing primitive/effect parsing, Type 3 paint, pattern/shading observers, cancellation and limits remain responsible for their contracts. |
| Reject aligned wrapped prose as tables | This explains six of eleven supplied failures; five fail without any table candidate. The new wrapped-prose branch requires column-major paint order; the established narrative rule remains intact. |
| Relax component line counts or suppress the old narrative rule | Existing article, rotated and native fixtures regress. These attempts are rejected. Genuinely empty gutters take priority over the previous few-crossings fallback. |
| Count positioned artifacts as words | Chromium can supply individual glyphs, including digits. A fragmented numeric-row negative control reproduces incorrect splitting. Alphabetic lexical words use the existing line-spacing rules, with separate ordinary and prose-only gutter evidence. |

Narrow anchors also require local clearance when applied: a gutter in a distant
section must not split justified text where a neighboring line crosses that
position. The component negative control reproduces the split before that check.

The supplied matrix passes all 24 two/three-column PDFs with both default 24 pt
and explicit 8 pt gutters: complete paragraph Text/Markdown in column order,
without false tables. All 90 font/size character streams remain correct;
five existing CJK spacing variants remain unchanged. Real table controls and
existing native/rotated layout tests protect table preservation.

Stop broader interpretation-sharing work at this measured candidate: both
profiles improve with exact complete exports. Preserve focused correctness,
cancellation and resource-limit checks when changing these paths again.
