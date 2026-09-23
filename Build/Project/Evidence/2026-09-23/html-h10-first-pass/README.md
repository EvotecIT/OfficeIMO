# H10 unfamiliar-page first pass, macOS arm64

This run uses OfficeIMO source `37045f949747ec060424e347d0a4941f4848368e`,
Chromium `151.0.7922.34`, PeachPDF `0.9.19`, .NET 10, and an 816 x 900 CSS-pixel
viewport. The runner's clean-source gate confirmed that the HTML, HTML PDF, MHTML,
PDF and runner assemblies came from that commit. Both archives were captured from
public pages on 2026-09-23 and replayed with Chromium's network access disabled.
They remain in the named ignored task directory, not in this repository's source
or default test corpus. The sources are [W3C WAI Tables tutorial](https://www.w3.org/WAI/tutorials/tables/)
([WAI reuse terms](https://www.w3.org/WAI/about/using-wai-material/)) and the
[MDN Grid layout guide](https://developer.mozilla.org/en-US/docs/Web/CSS/Guides/Grid_layout/Basic_concepts)
([MDN content license](https://developer.mozilla.org/en-US/docs/MDN/Writing_guidelines/Attrib_copyright_license)).

| Frozen source | SHA-256 of MHTML | Bytes / resources | Chromium screen / print pages | OfficeIMO screen-media / screen-snapshot / print pages | PeachPDF print pages |
| --- | --- | ---: | ---: | ---: | ---: |
| W3C WAI Tables | `c33e96c21941db20dc750f7dd01b1a9190b34b81460c7e419836dc8a3065abf9` | 181,175 / 10 | 3 / 3 | 6 / 7 / 4 | 3 |
| MDN Grid layout | `b7e01055e9c6ea917f9e095eba221616cf3abfd5c313485dc4104ff692f1d55a` | 882,840 / 79 | 21 / 16 | 12 / 13 / 9 | 25 |

These are distinct rendering intents. Page-count differences across screen-media,
snapshot and print output do not establish a defect by themselves. Both exact-archive
replays completed all six PDF operations without runner failures. The H4
advanced-held-out acceptance gate also passed all eight cases at this commit.

Visual inspection of the first printed pages and extracted text identified these
remaining product gaps:

- W3C: OfficeIMO's navigation links wrap into narrow columns and its tutorial
  illustrations occupy roughly half the reference width, disturbing adjacent text.
  Chromium and PeachPDF both keep 120px illustrations beside the copy. This is a
  CSS/layout fidelity gap; the HTML and ordinary tutorial text remain present.
- MDN: OfficeIMO's managed print omits the code examples serialized inside
  declarative shadow-root templates. Chromium and PeachPDF print their code text.
  In the default `pdftotext` extraction, `wrapper` occurs 29 times in each
  reference and only twice in OfficeIMO, where those two are surrounding prose.
  The browser-backed route remains necessary for this page
  until the managed static contract has qualified shadow-root projection.

The first pass also exposed three PDF/layout failures that this commit fixes:
an empty image alternative text value could abort PDF output, a positive-axis path
clip could exceed the PDF canvas, and a definite grid container's `auto` track
could expand to an item's max-content width. Focused regressions cover each;
the grid change also preserves inline-grid shrink wrapping. The complete HTML
suite passed 3,220/3,220 tests on .NET 10.

Reproduce from the saved archives with `html-mhtml-evidence --mhtml <archive>
--output <new-directory> --replay-browser --require-clean-source`, as documented
in the [comparison runner](../../../../../OfficeIMO.Pdf.Benchmarks.Comparisons/README.md).
This is diagnostic evidence from two documentation pages, not an H10 corpus
acceptance claim or a general PeachPDF parity claim. The open corpus, editable
format, resource-report and cross-platform gates remain in the
[roadmap](../../../../../Docs/ROADMAP.md#unfamiliar-page-conversion-qualification).
