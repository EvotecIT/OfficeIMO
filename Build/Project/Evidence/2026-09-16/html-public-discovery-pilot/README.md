# Linux public-page discovery pilot

The [internal pilot](../../../HtmlPublicPilot/README.md) fetched three public
pages on 2026-09-16 and ran parsing, JavaScript, capture, screen rendering and
both PDF intents in a rootless, network-disabled Podman image. The host fetched
only the document and resources requested by the isolated worker, with public
address checks and byte budgets. Every successful `outcome.json` records the
same full image ID,
`sha256:a564776b03261e4629bdc8099337cda60421acd5f32af944b1e23f09a59357be`,
verified renderer and script-worker file-set hashes, output hashes, isolation
policy and confirmed container removal. `acquisition.json` retains the exact
URL, fetch time, redirect and connected-address provenance without retaining
third-party page bytes.

| Page | Fetch time UTC | Discovered assets | Result |
| --- | --- | ---: | --- |
| [OfficeIMO homepage](https://officeimo.com/) | 14:22:14 | 21 | Screen PNG, six-page print PDF and six-page screen-to-page PDF. One stylesheet inserted by JavaScript was fetched on a bounded retry. |
| [W3C CSS example](https://www.w3.org/Style/Examples/011/mypage.html) | 14:23:24 | 1 | Screen PNG and both PDF intents; the linked stylesheet was discovered automatically. |
| [WPT first-letter reference](https://wpt.live/css/css-pseudo/first-letter-001-ref.html) | 14:24:03 | 0 | Screen PNG and both PDF intents without external assets. The retained outputs are covered by the [WPT BSD 3-Clause license](../html-public-oci-pilot/WPT-LICENSE.md). |

The OfficeIMO input was 69,306 bytes (SHA-256
`b51373b1c45304c38b53b995ae52c4be746adf5636e2b5cb475c795a7b9271cf`).
Its [OfficeIMO screen](officeimo-homepage/screen.png) is 816 × 5721; the
[Chromium reference](officeimo-homepage/chromium-screen.png), captured at the
same 816 × 720 viewport, is 816 × 5503. The two images were visually inspected.
The navigation, hero text, cards and collapsed FAQ are recognizable, but the
Studio hero image is absent from the OfficeIMO result, dark strips appear behind
some section headings, and typography and spacing differ. The Chromium full-page capture did not scroll each lazy-loaded
example into view, so its example placeholders are not a controlled image
comparison. The OfficeIMO PDFs were reopened with `pdfinfo` and each has six A4
pages; the first print page was visually inspected. These results establish
usable output for these named pages, not browser pixel parity.

The first OfficeIMO run with corrected closed-dialog and closed-details
semantics exhausted the original 256 MiB container memory budget while
embedding a translucent gradient in the PDF; `oom-256m.json` retains that
failure. The completed run used a verified 512 MiB container limit, one CPU,
32 PIDs and the same no-network and read-only controls. The site has no
skipped resources in this run. Browser-reference differences, richer hostile
inputs, broader responsive-source and module-graph cases, cross-host redirects,
and Windows and macOS isolation remain open in [the product roadmap](../../../../../Docs/ROADMAP.md).

The [controlled OCI fixture run](hostile-fixtures/summary.json) used the same
whole-pipeline image definition with ID
`sha256:3aa6fe3bb7c25e35f211922019f6434bf6ca4459fce7caf84625709fed25412b`.
Malformed table markup with an inline script produced a [screen image](hostile-fixtures/malformed-markup/screen.png)
and both PDFs; the image was visually inspected and both PDFs reopened as
single-page A4 files. The responsive-picture case requested only its active
single-candidate wide SVG by canonical absolute URL, omitted the narrow and
fallback resources, and visibly rendered the selected blue image. The
module-graph case requested its canonical root URL in round one and its relative
dependency URL in round two, then rendered the dependency's exported text. The
static-frame case requested its canonical iframe document URL in round one and
resolved the frame's relative stylesheet and external script from the
document's final URL in round two. Root JavaScript observed the loaded nested
document while its inline, external and event-attribute child scripts remained
inert; the retained
[frame screen](hostile-fixtures/frame-document/screen.png) and both PDFs contain
only the outer-document text, matching the current boundary that does not create
child-frame execution realms or project frame bodies into captures and rendered output.
All four success-case screen images were visually inspected; all eight PDFs
reopened as single-page A4 files and contained their expected outer-document
text. The remaining fixtures rejected 129 distinct script URLs, stopped an
oversized captured document, and interrupted a nonterminating script. All seven
runs reported the expected result and confirmed exact container removal. These
generated fixtures validate isolated static resource replay and failure
recovery, not child-frame execution, frame-body rendering, public-host
redirects, DNS changes, or arbitrary-site compatibility.

The [controlled acquisition tests](../../../../../OfficeIMO.Html.Runtime.Tests/RuntimePublicResourceBrokerTests.cs)
direct synthetic public DNS answers through a loopback test transport. They
prove same-host and explicitly approved cross-host redirect provenance,
per-hop DNS revalidation, rejection before an unapproved second connection,
and declared-response byte limits without depending on mutable public DNS.
These checks cover host-side acquisition; the redirect and DNS-change cases
still need to pass through the complete OCI capture and rendering pipeline.
