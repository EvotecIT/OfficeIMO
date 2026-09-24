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

## W3C follow-up on the same frozen bytes

Source `cb2f25f0fddd851b82478ee5a446d80abe3cd3ca` fixes two additional
engine defects exposed by the W3C print comparison. The media-query length
parser previously treated the `e` in `35em` as an incomplete scientific
exponent, so the page's wide-layout breakpoint never applied. The flex auto
basis also measured only DOM text, excluding generated print URLs from nested
`::after` content; breadcrumb items therefore received overly narrow widths.
Focused regressions cover both, the full .NET 10 HTML suite passed 3,225/3,225,
and exact-commit offline replay again completed all six operations without
runner failures. The H4 advanced-held-out gate passed 8/8 at this commit.

First-page visual inspection now shows 120px tutorial illustrations beside
their copy and substantially less breadcrumb wrapping. OfficeIMO print still
uses four pages versus three for Chromium and PeachPDF; its breadcrumb occupies
more lines and the header differs. MDN's omitted code examples remain. These
residual differences need further layout and shadow-root qualification before
any print-fidelity claim.

## MDN saved-component follow-up on the same frozen bytes

Source `823ef0b292a4517ce15debc75ffe843adc84aa51` projects Chromium MHTML
`template shadowmode` snapshots into the managed static render tree. It retains
named/default slot content, including nested slot chains, without changing
ordinary HTML templates or the archive source. The PDF operation reports the
shadow-root approximation and omits shadow-scoped stylesheets with a separate
warning so their CSS cannot leak into unrelated content. Callers can disable
projection explicitly; exact scoped styling and live behavior remain browser
workflows. A read-only review found two nested-slot and scaling defects in the
first candidate; both were fixed and confirmed before this source commit.

The exact-commit offline replay used the unchanged MDN archive SHA
`b7e01055e9c6ea917f9e095eba221616cf3abfd5c313485dc4104ff692f1d55a`.
The runner recorded a clean source, zero failures across six operations, and
these print results:

| Engine | Pages | `wrapper` occurrences in Poppler text | Extracted words |
| --- | ---: | ---: | ---: |
| Chromium | 16 | 29 | 2,837 |
| OfficeIMO | 21 | 29 | 2,722 |
| PeachPDF 0.9.19 | 25 | 29 | 2,783 |

Before projection, OfficeIMO printed nine pages and `wrapper` appeared only
twice in surrounding prose. Visual inspection of the replayed first and second
pages confirms that code blocks now appear, while OfficeIMO's breadcrumb, Copy
control, preview spacing and pagination still differ from Chromium. The H4
advanced-held-out gate passed 8/8 on the initial projection commit, the final
MHTML tests passed 32/32 on both .NET 8 and 10, and the full .NET 10 HTML suite
passed 3,230/3,230 after the nested-slot fixes. This closes the observed static
content omission, not the remaining component-style or general unfamiliar-page
qualification gaps.

## W3C print-intent qualification on the frozen archive

Source `007a58e4125fe77ea48dc1f1a44b0d5d37ee5d12` adds explicit
zero-margin and opt-in local-font print diagnostics without changing OfficeIMO's
default PDF settings. It also reads packed SVG arc flags in the WAI wordmark,
keeps replaced and definite-width descendants at their automatic flex minimum,
and embeds a document-selected installed font when the caller explicitly allows
it, even for ASCII-only text. The source and runner received a read-only review;
its two flex findings were fixed with row and column regressions. The full
.NET 10 HTML suite passed 3,235/3,235, the SVG parser regression passed, and
the H4 advanced-held-out acceptance gate passed 8/8 at the clean source commit.

Both offline replays used unchanged archive hashes, a clean worktree, and exact
owner/runner assembly versions. They reported no operation failures:

| Frozen page and print intent | Chromium | OfficeIMO default | OfficeIMO zero margin | OfficeIMO zero margin with local fonts | PeachPDF 0.9.19 |
| --- | ---: | ---: | ---: | ---: | ---: |
| W3C WAI Tables, pages | 3 | 4 | 3 | 3 | 3 |
| MDN Grid layout, pages | 16 | 21 | 20 | 20 | 25 |

The W3C default-page difference is largely the margin setting: OfficeIMO
defaults to 48 CSS-pixel outer margins, while this Chromium print reference
uses zero. Matching that setting yields three pages. The local-font lane embeds
Trebuchet MS on this macOS host; font availability is host-dependent and does
not alter the default policy. At that source commit, visual comparison of the
first W3C print page showed a substantive layout gap: breadcrumb URLs wrap,
heading treatment and text density differ, and fewer tutorial entries fit on
page one than in Chromium. Equal page count is therefore not visual parity.
The restored WAI logo and 120px tutorial illustrations render in the expected
positions. MDN's
saved code remains present (`wrapper` appears 29 times in extracted OfficeIMO
print text), but its component styling and pagination still differ. These
residuals, broader independently chosen page classes, editable target routes,
and cross-platform qualification remain open in the roadmap.

## W3C installed-font measurement follow-up

The opt-in local-font PDF lane now measures text with the same installed face
that it embeds. On the unchanged W3C archive, this keeps the printed breadcrumb
URLs for Home, Design & Develop, Tutorials and Tables on their respective lines,
as in the frozen Chromium print. Both outputs remain three pages with zero
margins. The default font policy is unchanged and may still wrap the URLs.
First-page inspection still shows misplaced breadcrumb separators and different
tutorial-card text flow and vertical spacing; matching page counts and link
lines do not establish full visual parity. The other frozen pages, editable
projections and supported-platform budgets remain open in the roadmap.
