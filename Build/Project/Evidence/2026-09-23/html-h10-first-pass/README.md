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

At clean source `17fc174b4`, the same MDN archive replayed offline with no operation failures. Chromium print remained 16 pages; OfficeIMO's zero-margin local-font print produced 19. All print pages were rasterized and inspected at contact-sheet scale, with the first two page pairs also inspected at higher resolution. The article text and saved code blocks are present throughout, but OfficeIMO's breadcrumb items crowd together and its code/preview components repeatedly retain large bordered boxes where Chromium print uses compact controls and code. These repeated boxes contribute to the extra pages. The managed report explicitly marks the serialized shadow-root projection and unavailable shadow-scoped styling; this is a current static-rendering limit, not evidence that the missing styling can safely be applied as ordinary document CSS. The report also contains 36 unavailable-resource observations and many repeated OpenType feature warnings, whose counts do not measure visible loss. The exact report, PDFs, page rasters and contact sheets are retained under `Ignore/HtmlUnknownPageQualification/h10-mdn-baseline-17fc174b4/`. Fine-detail review beyond the first two page pairs and the separate screen PDF intents remain open; exact component appearance remains a browser-backed route.

A September 25 browser probe clarified the MDN style boundary. At the 816 × 900
capture viewport, the document light tree exposed 75 open shadow-root hosts;
60 held adopted stylesheets with 1,787 CSS rules. This count does not traverse
nested shadow roots. The frozen MHTML contains 179 serialized
open-shadow templates but no `<style>` elements. Chromium's
`Page.captureSnapshot` also omitted a `<style>` deliberately inserted into a
minimal open shadow root immediately before capture. Copying those 60 adopted
stylesheets into their live shadow roots likewise produced no serialized
shadow styles.
The frozen input therefore lacks those component rules; applying them as
ordinary page CSS would lose their scope. This establishes a capture-format
limit for the current static MDN replay, not a managed-renderer fix or a new
browser-fidelity qualification. Exact component appearance still requires the
explicit browser-backed route. The minimized and live snapshots, probe source,
and measured counts are retained under
`Ignore/HtmlUnknownPageQualification/h10-mdn-shadow-style-probe/`.

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

## W3C paragraph-backed list follow-up

Source `da49000169cb048c76afba32f40cd64fad72f3b9` places outside list
markers beside the first text line when an item contains block children, and
reserves the normal 2.5em list inset unless authored CSS overrides it. The
unchanged W3C archive replayed with a clean worktree and no operation failures;
the zero-margin local-font print and Chromium print still have three pages each.
All three OfficeIMO pages were rasterized and inspected. The two bullets in
"Why is this important?" now appear in the list gutter on page two instead of
being absent or clipped. The page still differs from Chromium in text weight,
line flow and spacing, while the breadcrumb separators and footer layout remain
visibly different. This is a bounded list-layout correction, not page-level
browser equivalence.

The full .NET 10 HTML suite passed 3,318/3,318, the H4 advanced-held-out gate
passed 8/8, and the clean-source macOS H4 budget passed with a 6,097 ms cold
process and a 2,647 ms maximum warm iteration. The replay, inspected rasters
and budget report are retained under `Ignore/HtmlUnknownPageQualification/`
in `h10-wai-list-inset-clean-da4900016` and `h10-h4-budget-da4900016`.

## W3C positioned breadcrumb separators

Source `cfffdbd1bc7e9cd3d0ea81032e68c9282c4e50ca` positions explicitly sized
absolute block pseudo-elements within a relative or sticky host without adding
their height to normal flow. It paints them in the host's positioned stacking
bands, including the order relative to absolute and in-flow positioned children.
Other pseudo-element positioning forms remain outside this qualified path.

The exact frozen W3C archive above replayed offline from a clean worktree with
no operation failures across ten output and reference operations. Chromium and
OfficeIMO's zero-margin local-font print each produced three pages; OfficeIMO's
default print still produced four. All three print pages from both engines were
rasterized and inspected. The breadcrumb separators now sit beside their links,
and the local-font OfficeIMO report has no `li::after` positioning warnings.
Text weight, tutorial-card density and footer layout still differ visibly from
Chromium, so the W3C print-fidelity item remains open.

The full .NET 10 HTML suite passed 3,320/3,320, focused .NET 8 positioning
tests passed 3/3, the H4 advanced-held-out acceptance gate passed 8/8, and the
clean-source macOS static budget passed with a 5,374 ms cold process and a
2,008 ms maximum warm iteration. The replay, inspected rasters, acceptance
report and budget report are retained under `Ignore/HtmlUnknownPageQualification/`
in `h10-wai-positioned-pseudo-clean-cfffdbd1b`,
`h10-h4-acceptance-cfffdbd1b` and `h10-h4-budget-cfffdbd1b`.

## W3C tutorial-card text beside floated images

Source `64d1f971ce29335c44790b57ff8154031936525e` lets a float that follows
inline text share the current line when its measured box fits. This matters for
the W3C tutorial cards: each link's illustration floats into the left gutter,
between the bold title and its generated print URL. The renderer previously
ended the title line at that image even though the URL had room beside it. The
shared float path now also collapses whitespace across a same-line float and
still moves the float below a line that lacks room.

The unchanged frozen W3C archive replayed offline from this clean source with
no failures across ten operations. Chromium and OfficeIMO's zero-margin
local-font print still produce three pages; OfficeIMO's default print still
produces four. All three local-font print pages were rasterized and compared
with the frozen Chromium pages. Four tutorial cards now fit on OfficeIMO's
first page, and their title and print URL stay on the same line where space
permits. Chromium still starts the fifth card on page one, while OfficeIMO
starts it on page two. Text weight and footer layout also differ, so the W3C
print-fidelity item remains open. The OfficeIMO screen-media projection moves
from four pages to three; the screen-snapshot projection remains four.

The full .NET 10 HTML suite passed 3,325/3,325, the focused .NET 8 cases passed
5/5, and the H4 advanced-held-out acceptance gate passed 8/8. The clean-source
macOS static budget passed with a 5,364 ms cold process and a 1,920 ms maximum
warm iteration. Reports and inspected rasters are retained under
`Ignore/HtmlUnknownPageQualification/` in
`h10-wai-float-after-text-clean-64d1f971c`,
`h10-h4-float-after-text-64d1f971c`, and
`h10-h4-budget-float-after-text-64d1f971c`.

## W3C tutorial-card float containment

Source `272f1fc2f82d9c7c595c0dce92e75b782e2f1ad2` corrects a further
16-pixel gap after each tutorial card. The card list uses `overflow:auto` to
contain images floated inside paragraphs. When the float extends below the
paragraph's text, its overhang now consumes the paragraph's final bottom
margin inside the containing list item. Explicitly sized paragraphs and
paragraphs that contain their own floats retain that margin.

The unchanged W3C archive was replayed offline from clean source. In the
zero-margin local-font print comparison, OfficeIMO and Chromium both produce
three pages, and the fifth card begins on page one in both. The first four
card images align closely at normal reading size. OfficeIMO's default-margin
print still has four pages; text weight, footer layout and some density differ,
so the W3C print-fidelity item remains open. The .NET 10 HTML suite passed
3,332/3,332, H4 advanced-held-out acceptance passed 8/8, and the macOS static
budget passed with a 5,350 ms cold process, 1,974 ms maximum warm iteration,
and matching cold/warm fingerprints. Reports and the inspected page raster are
retained under `Ignore/HtmlUnknownPageQualification/` in
`h10-wai-float-bfc-clean-272f1fc2f`, `h10-h4-float-bfc-272f1fc2f`, and
`h10-h4-budget-float-bfc-272f1fc2f`.

## W3C installed font-face selection

At clean source `a7762d428`, the macOS Trebuchet MS family now keeps its
installed Bold, Italic and Bold Italic programs when the name-table style is
localized. The system-font loader uses the OS/2 weight and style flags for face
classification, prefers weights nearest 400 and 700 within the regular and
bold slots, and retains the independent oblique flag. This fixes the W3C
print lane's misleading `TrebuchetMS-Bold` PDF resource, which previously
embedded weight-400 data. All three pages now embed weight-700 Bold data;
the tutorial-card links and emphasized list text have the expected visible
weight. This remains an opt-in local-font result on a host with those fonts.

The unchanged W3C archive replayed from the clean code head with no runner
failures. Chromium and OfficeIMO zero-margin local-font print both have three
pages. The fifth card's floated image still begins in a sliver at the bottom
of OfficeIMO page one and continues at the top of page two, whereas Chromium
places the complete image on page two. Footer spacing and some text flow also
remain different. Those are open W3C print-fidelity gaps, not resolved by
the font correction. The full HTML suite passed 3,390/3,390 on .NET 10 and
.NET 8, the PDF font-family group passed 128/128 on .NET 10, H4
advanced-held-out acceptance passed 8/8, and the clean-source macOS H4 static
budget passed with a 5,483 ms cold process, 2,060 ms maximum warm iteration,
497 MB sampled peak and matching cold/warm fingerprints. Reports and page
rasters are retained under `Ignore/HtmlUnknownPageQualification/` in
`h10-wai-font-face-clean-a7762d428`, `h4-font-face-clean-a7762d428`, and
`h4-font-face-budget-a7762d428`.

## W3C tutorial image at a page break

At clean source `0d9f85fdf`, line ends inside a floated box are no longer offered as page breaks. The layout keeps those lines for widow/orphan counting; an initial version that removed them from both lists forced a break through text, so a focused regression now covers that failure too. The full HTML suite passed 3,398/3,398 on .NET 10 and on a .NET 8 rerun. An unrelated concurrent detached-projection test failed once in the first .NET 8 run, passed in isolation, and passed in that full rerun. Independent read-only review identified the widow/orphan regression before the final commit and confirmed the repair.

The unchanged W3C archive replayed offline with no runner failures. Chromium and OfficeIMO zero-margin local-font print each have three pages; default-margin OfficeIMO print has four. All three final managed and Chromium print pages were rasterized at 96 dpi and inspected. The fifth tutorial illustration no longer starts as a sliver at the bottom of managed page one: it begins intact on page two. Chromium leaves that card's title on page one, while OfficeIMO moves the title with the illustration. Text weight, footer placement, and some line flow also differ, so the W3C print-fidelity item remains open. The local-font operation reports 64 warnings and declared loss, with no forced-fragment diagnostic. The report, PDFs, and compared page rasters are retained under `Ignore/HtmlUnknownPageQualification/h10-wai-float-fragment-clean-0d9f85fdf/`.

The exact-head H4 advanced-held-out acceptance gate passed 8/8. The macOS static budget passed with a 6,083.8 ms cold process, 2,071.5 ms slowest warm iteration, 501,465,088-byte sampled peak, and matching cold/warm fingerprints. Those reports are under `h10-h4-float-fragment-clean-0d9f85fdf/` and `h10-h4-float-fragment-budget-clean-0d9f85fdf/` beneath the same ignored parent. Windows and Linux budgets remain open.

The remaining title-placement mismatch was reduced to a 100-by-60-pixel
print fixture. A 40-pixel prelude leaves space for the first lines of a
paragraph but not its 30-pixel floated image. Chromium keeps the title and
two text lines on page one, defers the complete image to page two, and flows
the remaining text around that image. OfficeIMO moves the paragraph and image
to page two. Both outputs have three pages; page count therefore conceals
this difference. Moving only the image's painted visual would leave its
reserved height and subsequent text flow wrong. The remaining renderer work
is page-aware float deferral with flow continuation and an intact replaced
box, followed by a replay of all frozen W3C print pages and H4. The fixture,
both PDFs, inspected first-two-page rasters and worker reports are retained
under `Ignore/HtmlUnknownPageQualification/h10-wai-float-break-probe/`.

## W3C tutorial float continuation at the page boundary

At clean source `2eda4f38e`, the layout now defers a floated image that fits a
fresh page while allowing preceding paragraph lines to use the current page.
The float starts at the legal break chosen by widow/orphan pagination, and
subsequent lines wrap around it. If there is no legal break, the original
paragraph moves intact; sibling and parent margin collapse retain their
positions. Focused regressions cover these cases. The full HTML suite passed
3,412/3,412 on both .NET 10 and .NET 8. Independent read-only review found
the no-break and collapsed-margin cases and confirmed their fixes.

The unchanged W3C archive (`c33e96c21941db20dc750f7dd01b1a9190b34b81460c7e419836dc8a3065abf9`)
replayed offline at that clean commit with no runner failures. Chromium and
OfficeIMO zero-margin local-font print each produced three pages; OfficeIMO's
default-margin print produced four. All three comparable print pages were
rasterized at 96 dpi and inspected. The fifth card's title and first text
lines now remain on page one, and its complete illustration begins at the top
of page two, as in Chromium. The third page is unchanged. Breadcrumb spacing,
some text wrapping and the footer's line placement still differ. The managed
report declares degraded fidelity with 74 warnings, including 26 unavailable
resources, 11 unavailable font faces, 17 unsupported SVG instances and 14
unsupported OpenType features; these counts are diagnostic observations, not
distinct visible omissions. The H4 advanced-held-out visual acceptance gate
passed 8/8. The clean report, PDFs, compared page rasters and H4 acceptance
report are retained under `Ignore/HtmlUnknownPageQualification/` in
`h10-wai-page-deferral-clean-2eda4f38e/` and
`h10-h4-page-deferral-clean-2eda4f38e/`.

The first macOS H4 static time-budget run at that commit failed: its
slowest warm iteration was 3,892 ms against a 3,000 ms ceiling. A later run
under heavy concurrent Xcode/Swift load also exceeded the time ceiling.
Cold/warm fingerprints and allocation remained stable. On an idle host at
clean documentation head `d1e6a5866`, the unchanged source passed the H4
static budget: 6,359.8 ms cold process, 2,835.4 ms slowest warm iteration,
471,252,992-byte sampled peak, matching cold/warm fingerprints, and passing
cancellation. The passing report is retained in
`h10-h4-wai-float-idle-budget-d1e6a5866/`; the first contended run is in
`h10-h4-page-deferral-budget-clean-2eda4f38e/`. Windows and Linux budgets
remain open.

## Matched-print all-page recheck

At clean renderer head `bb3526a06`, the same frozen WAI archive
(`c33e96c21941db20dc750f7dd01b1a9190b34b81460c7e419836dc8a3065abf9`)
replayed with no runner failures. Chromium and OfficeIMO's zero-margin,
opt-in installed-Trebuchet print lanes each have three pages. At 96 dpi, all
three page pairs were inspected at normal reading size. The breadcrumb uses
two rows in both outputs, the fifth card continues across the same page
boundary, and the footer begins on page three in both. Normalized ordered
Poppler words match exactly on pages one and two (334/334 and 326/326);
page three has a 0.990 sequence ratio (254 browser and 253 managed words),
with the difference around the “Next: One Header” navigation label. This is
selected-content and page-flow evidence under a matched macOS font policy,
not a claim for OfficeIMO's default-margin or portable-font print output.

The print images still differ in some link wrapping, text weight and footer
placement. The managed operation remains `Degraded` with 74 warnings: 26
unavailable resources, 11 unavailable `@font-face` sources, 17 unsupported
SVG-content instances, 14 unsupported OpenType features, five overflow
snapshots and one stylesheet-URL note. The frozen MHTML has no
`/WAI/assets/images/icons.svg` part although its HTML references that
external icon sprite. These diagnostics require source/visibility
classification before a resource-complete claim. The exact-head replay is
under `Ignore/HtmlUnknownPageQualification/h10-wai-root-rem-clean-bb3526a06/`;
the inspected page rasters are under
`Ignore/HtmlUnknownPageQualification/h10-visual-inspection-bb3526a06/`.

At clean documentation head `72eaa7e02`, the macOS H4 static budget kept
identical fingerprints for all 76 cold/warm outputs and passed allocation,
peak memory, output-size and cancellation limits. Warm iterations took
2,520, 2,621 and 3,638 ms, so the last exceeded the 3,000-ms ceiling.
Nine concurrent `swift-frontend` processes each used roughly 85–91% CPU on
this ten-core host immediately after the run; the timing failure cannot yet
be attributed to the renderer alone. The failed report is
retained under `Ignore/HtmlUnknownPageQualification/h10-h4-budget-root-rem-clean-72eaa7e02/`.
An idle-host repeat and Windows/Linux budgets remain open.

At clean source `e983c43c1`, an offline replay of the same WAI archive and browser reference reconfirmed three pages for Chromium print and OfficeIMO zero-margin print with installed Trebuchet. Every page pair was inspected at reading size. The article and cards remain close, while the navigation box's top border starts in the last two pixels of managed page two and its remaining sides continue on page three without the browser's bottom edge. Its URL also wraps differently. This is a paged layout/paint gap, not a changed source page. `pdfinfo -url` found the same 31 distinct link destinations in both print PDFs, with 63 Chromium and 89 managed annotation rectangles; matching destination sets do not establish matching link geometry.

All 26 `HtmlRenderResourceUnavailable` URLs in that managed print report are absent from the archive's MIME parts: two favicons, the external icon sprite and checkbox SVG, and 22 font-file candidates from 11 `@font-face` rules. The five captured tutorial PNGs, four social/email SVGs, stylesheet and HTML are present. The missing icon and font bytes therefore cannot be recovered by an offline resource-resolver fix against this frozen input; a separate provenance-tracked capture or portable-font policy would be needed to qualify their appearance. The 17 SVG-content warnings include repeated observations of uncaptured `<use>` references and do not mean 17 distinct missing image files. The exact-head PDFs, all six print rasters and report are under `Ignore/HtmlUnknownPageQualification/h10-wai-current-baseline-e983c43c1/`; H4 advanced held-out acceptance passed 8/8 with no failures under `Ignore/HtmlUnknownPageQualification/h10-h4-wai-current-e983c43c1/`. WAI print remains unqualified because of visible layout, link-geometry and portable-font gaps.

At clean renderer head `5996171f9`, the bordered navigation box no longer leaves its top edge on page two: its full border starts on page three, as in the offline Chromium print. The three page counts and 31 distinct PDF link destinations remain unchanged. All three page pairs were inspected at reading size; the first managed raster is pixel-identical to the prior head, and page three improves from 15.512 to 14.253 RGB mean absolute error at 96 dpi over the overlapping image area. The managed URL still wraps differently, and matching destination sets do not prove matching annotation rectangles. Focused regressions cover the page entry, collapsed sibling margins, ordinary and `display:contents` link targets, and negative top margins. The full HTML suite passed 3,434/3,434 on each of .NET 10 and .NET 8; the netstandard2.0 HTML build passed; clean-source H4 advanced held-out acceptance passed 8/8 with no failures. The exact-head WAI PDF/report and six print rasters are under `Ignore/HtmlUnknownPageQualification/h10-wai-page-entry-clean-5996171f9/`; H4 evidence is under `Ignore/HtmlUnknownPageQualification/h10-h4-page-entry-clean-5996171f9/`.

The clean-source macOS H4 static budget at this head kept matching cold/warm output fingerprints and stayed below the peak-memory ceiling, but failed time limits: 12,890.7 ms cold versus 8,000 ms allowed and 3,057.8–4,143.3 ms warm versus 3,000 ms allowed. Concurrent CasaRay simulator and Xcode helper processes were consuming several cores when the failure was inspected, so this run does not isolate a renderer performance regression. Its report is under `Ignore/HtmlUnknownPageQualification/h10-h4-budget-page-entry-clean-5996171f9/`. An idle-host macOS repeat plus Windows and Linux budgets remain open.

At clean renderer head `b2beae184`, the frozen WAI archive replayed offline with no operation failures. Chromium and OfficeIMO's zero-margin, installed-Trebuchet print lanes still have three pages and the same 31 distinct link destinations. The page-three "Next" URL now breaks after `one-` in both PDFs. All three page pairs were rasterized at 96 dpi and inspected: managed pages one and two are pixel-identical to the previous OfficeIMO head; page three's RGB mean absolute error against Chromium improves from 14.253 to 14.077. Chromium also measured `alpha/beta` and `abc/def` at equal min-content and max-content widths, supporting the intrinsic-sizing rule used for this URL. The full HTML suite passed 3,435/3,435 on .NET 10 and .NET 8, the netstandard2.0 HTML build passed, and clean-source H4 advanced-held-out acceptance passed 8/8. The exact-head report, PDFs and six print rasters are under `Ignore/HtmlUnknownPageQualification/h10-wai-solidus-clean-b2beae184/`; H4 evidence is under `h10-h4-solidus-clean-b2beae184/` in the same ignored parent. This selected URL-wrap fix does not qualify link annotation geometry, the uncaptured resources or portable font appearance. The macOS idle-host repeat and Windows/Linux H4 budgets remain open.
