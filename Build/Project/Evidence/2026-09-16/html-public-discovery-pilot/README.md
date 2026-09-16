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
inputs, responsive sources, module graphs, cross-host redirects and Windows and
macOS isolation remain open in [the product roadmap](../../../../../Docs/ROADMAP.md).

The [controlled OCI fixture run](hostile-fixtures/summary.json) used the same
whole-pipeline image definition with ID
`sha256:c508fe79426d7354b945f5bab8a56330e2cf86d89faec8c1718fcab7623e65fe`.
Malformed table markup with an inline script produced a [screen image](hostile-fixtures/malformed-markup/screen.png)
and both PDFs; the image was visually inspected and both PDFs reopened as
single-page A4 files. The other fixtures rejected 129 distinct script URLs,
stopped an oversized captured document, and interrupted a nonterminating
script. All four runs reported the expected result and confirmed exact
container removal. These generated fixtures validate isolated failure recovery,
not public-host redirects, DNS changes, or arbitrary-site compatibility.
