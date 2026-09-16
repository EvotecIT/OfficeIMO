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
inputs, broader module and frame execution cases, XMLHttpRequest with request
headers or beyond bodyless GET, non-GET dynamic requests, and Windows and macOS
isolation remain open in [the product roadmap](../../../../../Docs/ROADMAP.md).

The [controlled OCI fixture run](hostile-fixtures/summary.json) used the same
whole-pipeline image definition with ID
`sha256:d62bf153be9b4cce4d91669703debac00072a2115f81a0726bdf01da3b97a35f`.
Malformed table markup with an inline script produced a [screen image](hostile-fixtures/malformed-markup/screen.png)
and both PDFs; the image was visually inspected and both PDFs reopened as
single-page A4 files. The responsive-picture case requested only its active
`800w` SVG from a three-candidate set by canonical absolute URL at device pixel
ratio 2 and a 400 CSS-pixel source size. It omitted the `400w`, `1200w`,
inactive-source and fallback resources, exposed the same device density to the
application, and visibly rendered the selected blue image. The
module-graph case requested its canonical root URL in round one and its relative
dependency URL in round two, then rendered the dependency's exported text. The
scoped import-map case requested its root module, the `pkg/` prefix target, the
scope-specific `theme` target and a dynamically imported relative module in
four exact rounds. The deliberately absent global `theme` target was not
requested. Top-level await completed before the retained
[import-map screen](hostile-fixtures/import-map-graph/screen.png) and both PDF
intents rendered `Scoped import map ready 42`. The script-driven fetch case
requested its relative JSON endpoint in one runtime-discovery round, preserved
the `view=summary` query, removed the client-only fragment, decoded the response
with `Response.json()` and waited for the promise chain before capture. The
retained [dynamic-fetch screen](hostile-fixtures/dynamic-fetch-get/screen.png)
and both PDF intents rendered `Dynamic fetch ready 42`. The asynchronous XHR
case settled its handled offline error, exposed the missing bodyless headerless
GET through the bounded runtime trace, requested the canonical JSON URL in one discovery
round, preserved the query, removed the client-only fragment and replayed the
same application with supplied bytes. Its retained
[XHR screen](hostile-fixtures/dynamic-xhr-get/screen.png) and both PDF intents
rendered `XHR ready 42`. A paired XHR with the custom `X-Variant` request header
produced zero discovery rounds and retained its offline error path; its
[screen](hostile-fixtures/dynamic-xhr-headered-get-blocked/screen.png) and both
PDF intents rendered `Headered XHR remained offline`. The static-frame case requested
its canonical iframe document URL in round one and
resolved the frame's relative stylesheet and external script from the
document's final URL in round two. Root JavaScript observed the loaded nested
document while its inline, external and event-attribute child scripts remained
inert; the retained
[frame screen](hostile-fixtures/frame-document/screen.png) and both PDFs contain
only the outer-document text, matching the current boundary that does not create
child-frame execution realms or project frame bodies into captures and rendered output.
The controlled acquisition lane followed one same-host redirect and one
explicitly approved cross-host redirect, retaining the requested and final URLs,
each resolution, connected address, HTTP hop and server request. The acquired
documents then produced the retained [same-host](hostile-fixtures/acquisition-same-host-redirect/screen.png)
and [cross-host](hostile-fixtures/acquisition-cross-host-redirect/screen.png)
screen images and both PDF intents in separate network-disabled containers. A
per-hop DNS change from a public address to loopback was rejected before the
second connection, and a response declaring 4 MiB plus one byte was rejected
before rendering. All ten success-case screen images were visually inspected;
all twenty PDFs reopened as single-page A4 files and contained their expected
outer-document text. The remaining fixtures rejected 129 distinct script URLs,
stopped an oversized captured document, and interrupted a nonterminating
script. All thirteen render runs and four acquisition cases reported the expected
result, and every started container was removed. These generated fixtures
validate isolated static resource replay, controlled redirect acquisition and
failure recovery. They do not qualify child-frame execution, frame-body
rendering, mutable public DNS, live public-host redirects or arbitrary-site
compatibility.

The [controlled acquisition tests](../../../../../OfficeIMO.Html.Runtime.Tests/RuntimePublicResourceBrokerTests.cs)
direct synthetic public DNS answers through a loopback test transport. They
prove same-host and explicitly approved cross-host redirect provenance,
per-hop DNS revalidation, rejection before an unapproved second connection,
and declared-response byte limits without depending on mutable public DNS.
The retained controlled corpus carries the accepted redirect documents through
the complete OCI capture and rendering pipeline and records the rebinding and
oversized-response rejections before container startup.
