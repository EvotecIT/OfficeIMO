# H10 frozen-page expansion, macOS arm64

The [selection manifest](../../../../../OfficeIMO.Pdf.Benchmarks.Comparisons/Corpus/html-h10-page-selection.json) named these sources and acceptance markers before capture or OfficeIMO rendering. Capture used source commit `5ac589cd6`, an 816 x 900 CSS-pixel viewport, and live DOMContentLoaded plus 1000 ms. The archives are retained under `Ignore/HtmlUnknownPageQualification/` for exact offline replay; third-party page bytes are not committed. The manifest records each source and its rights link.

The final replay used clean source commit `4d5009b3939f5e091f4af979f839d7aac3d1baa6` and matching owner/runner assemblies, Chromium `151.0.7922.34`, and PeachPDF `0.9.19`. Chromium replay had network access disabled. All eight PDF operations completed on each archive, with no runner failures. The H4 advanced held-out gate passed 8/8 after the PDF fix.

| Frozen source | MHTML SHA-256 | Bytes / resources | Chromium screen / print pages | OfficeIMO screen-media / snapshot / print pages | PeachPDF print pages |
| --- | --- | ---: | ---: | ---: | ---: |
| [W3C HTML guide](https://www.w3.org/MarkUp/Guide/) | `ffd7f719936ff44208cdf1e150dbdacf3e8a6d1a35b638b22aafb7c623d2d3fb` | 75,026 / 4 | 5 / 5 | 5 / 6 / 5 | 5 |
| [NASA SVS Voyager gallery](https://svs.gsfc.nasa.gov/gallery/voyager/) | `0eef6666f91254dcfb57471cbd31b0f81c3cc2e79fb6d574c60e1fdfd4b6f666` | 4,929,842 / 23 | 4 / 4 | 17 / 20 / 3 | 3 |
| [Playwright TodoMVC](https://demo.playwright.dev/todomvc/) | `f9a8008271c52496e8264aee8001c0b46977e39a6910946ca766fc738585f340` | 10,901 / 2 | 1 / 1 | 1 / 1 / 1 | 1 |

First-page PDF raster inspection and extracted text show different outcomes by intent:

- W3C print keeps both required markers and the document text. At zero margins its page count matches both references, but its top navigation, heading, image alignment and paragraph density differ from Chromium. Equal pages are not visual equivalence.
- NASA is the highest-impact layout gap. Chromium print keeps the heading, overview, sidebar and two-column gallery together on page one. OfficeIMO print places a large dark region over nearly the entire first page and displaces or clips content; PeachPDF print also misplaces much of this page. OfficeIMO screen-media pagination retains overview text but loses the reference styling and gallery presentation, spreading content over 17 pages; snapshot pagination takes 20. Neither reference's poor print output excuses OfficeIMO's defect. Page counts across intents are not directly comparable.
- TodoMVC retains both predeclared markers in OfficeIMO and Chromium PDF text. OfficeIMO's heading is too dark and large and collides visually with the input area. PeachPDF's first page has a closer pale heading but omits the input prompt from rendered and extracted output. This frozen DOM proves only the post-load static surface, not interactive execution.

The NASA archive exposed two PDF conversion failures before `4d5009b39`: a missing icon-font private-use glyph aborted print conversion, and multiline image alternative text aborted screen conversion. The fix omits an unpaintable private-use icon with a reported loss and normalizes ASCII whitespace in fallback image text. Focused regressions, a full .NET 10 HTML suite pass (3,244/3,244 on rerun), and the H4 gate cover the change. One run of the full suite had an intermittent failure in the untouched concurrent detached-projection test; that test passed alone and in the full rerun, so the intermittence remains unclassified.

The NASA page was predeclared held-out, but its first capture exposed the failures and informed the fix. It is now a development case; a new independently chosen source must replace its held-out role before a held-out acceptance claim. Resource/loss reports for each replay lane, editable Word/Excel/PowerPoint/OneNote/RTF/Markdown artifacts, visual comparisons beyond the inspected first pages, and supported-platform time/allocation/peak-memory budgets remain open in the [roadmap](../../../../../Docs/ROADMAP.md#unfamiliar-page-conversion-qualification). This evidence neither establishes general browser equivalence nor a blanket PeachPDF parity claim.

Reproduce a page from its retained `source.mhtml` with `html-mhtml-evidence --mhtml <archive> --output <new-directory> --replay-browser --require-clean-source`, as documented in the [comparison runner](../../../../../OfficeIMO.Pdf.Benchmarks.Comparisons/README.md). The clean replay directories are `h10-w3c-guide-clean-4d5009b39`, `h10-nasa-voyager-clean-4d5009b39`, and `h10-playwright-todomvc-clean-4d5009b39` under `Ignore/HtmlUnknownPageQualification/`.

## Explicit screen-resource and report follow-up

Source `666c0b77ab28e607cae10ebc9d80488c2f47e51a` adds an MHTML-aware explicit PDF request path. The previous comparison runner sent screen intents through the generic HTML adapter, which blocked archived stylesheets and images as remote resources. The runner now uses the MHTML bridge for both screen intents and records MIME diagnostics plus per-operation PDF warnings and loss status in schema-2 JSON. A heading shifted above the first snapshot page also no longer aborts PDF bookmark creation. The independent read-only review found a missing exact-head check for the MHTML/PDF bridge; the final gate includes that bridge and the relevant Email, HTML-core, AngleSharp-adapter and Core assemblies.

All three unchanged archives replayed with `--require-clean-source`, matching owner versions and zero operation failures. The full .NET 10 HTML suite passed 3,247/3,247; the MHTML/PDF bridge built for netstandard2.0 and net8.0 with no warnings; the H4 advanced held-out visual gate passed 8/8. The replay directories use the suffix `clean-666c0b77a` under `Ignore/HtmlUnknownPageQualification/`.

| NASA intent | Chromium | OfficeIMO before MHTML screen routing | OfficeIMO at `666c0b77a` | PeachPDF |
| --- | ---: | ---: | ---: | ---: |
| Screen-media paged | 4 pages | 17 pages | 3 pages | Not measured |
| Screen snapshot paged | Frozen screen reference | 20 pages | 2 pages | Not measured |
| Print | 4 pages | 3 pages | 3 pages | 3 pages |

The screen counts changed because the archived CSS and images now reach the managed renderer. First-page raster inspection confirms that the screen snapshot has the sidebar and overview layout absent from the previous unstyled output. It still clips the title and mispaints or omits gallery cards. Print still puts a dark field over the first page and displaces content. These are layout and paint gaps, not proof of a remaining screen-resource routing failure. The NASA archive has only 23 saved resources; the OfficeIMO report records 50 unavailable resource references per print or screen-media operation, including assets outside the saved archive. Some gallery assets are saved yet still do not appear correctly, so capture completeness and paint/layout require separate investigation. W3C and TodoMVC page counts did not change. W3C still reports one missing CID stylesheet resource; TodoMVC reports a form-field typography approximation.

The schema-2 report measures operation-level loss, but raw warning counts include favicons and offscreen references and should not be read as visible-content loss counts. Editable-format artifacts, full-page visual scoring, and supported-platform allocation and peak-memory budgets remain open.
