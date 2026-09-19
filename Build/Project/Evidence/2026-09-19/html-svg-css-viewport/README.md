# CSS-sized inline SVG viewport evidence

The [line reference](https://wpt.live/css/compositing/line-with-svg-background-ref.html)
contains ten inline SVGs with no intrinsic dimensions. Its stylesheet gives each
SVG a 500 × 60 CSS-pixel viewport and gray background; each SVG draws a blue
20-pixel stroke. The earlier isolated run retained
[here](../html-dynamic-public-profile/README.md) outlined the boxes but omitted
both paints. The shared HTML renderer now passes the resolved painted object size
to the bounded SVG drawing reader. It retains the CSS background and the SVG
stroke in the same scene used by all three output intents.

The live pages were acquired from Web Platform Tests on 2026-09-19 under its
[BSD 3-Clause license](../../2026-09-16/html-public-oci-pilot/WPT-LICENSE.md).
This folder retains acquisition metadata and rendered output, not the upstream
source bytes. The Windows host acquired the pages; a rootless, networkless
Ubuntu/Podman container parsed, scripted, captured, and rendered them. Each
`outcome.json` records the verified isolation policy, published binary hashes,
image ID `sha256:bb0cc1431348aad8545cc4396b5626754a9be3a4d0c29d43ff84ab4152a721d6`,
output hashes, and container removal.
The [whole-workflow summary](controlled-summary.json) records 22 passing render
cases and six passing acquisition cases under that same image and policy.

| Reference | OfficeIMO screen | Chromium reference | Observed comparison |
| --- | --- | --- | --- |
| [Blue line on gray SVG](https://wpt.live/css/compositing/line-with-svg-background-ref.html) | [PNG](named-line/screen.png) | [PNG](chromium-line.png) | Both 816 × 720; zero differing pixels. All ten backgrounds and strokes appear. |
| [Text on blue SVG](https://wpt.live/css/compositing/Text_with_SVG_background-ref.html) | [PNG](named-text-blue/screen.png) | [PNG](chromium-text-blue.png) | Both 816 × 720; 8,193 differing pixels. The CSS box starts at the page origin in OfficeIMO and at the browser's 8-pixel body margin in Chromium; glyph metrics also differ. |
| [Text on gray SVG rows](https://wpt.live/css/compositing/text-with-svg-background-ref.html) | [PNG](named-text-gray/screen.png) | [full-page PNG](chromium-text-gray-full.png) | Both 816 × 780; 42,203 differing pixels, primarily in default SVG font metrics. The full-page comparison matches `screen-full-page-v1` rather than cropping at the 720-pixel viewport. |

Each `named-*` folder also contains the [print](named-line/print.pdf) and
[screen-to-page](named-line/screen-to-page.pdf) PDFs for that case. Independent
Poppler inspection found one A4 page in each of the six PDFs. The two text
pages retain searchable text in both PDF modes. Rasterizing the line PDFs
confirms that both modes paint all ten gray backgrounds and blue strokes:
[print raster](named-line/print-raster.png) and
[screen-to-page raster](named-line/screen-to-page-raster.png). Their placement
differs because print and screen-to-page are distinct layout intents.

These observations qualify this selected SVG line case, not general SVG or
browser equivalence. The default body margin and SVG text-font differences
remain open in [the roadmap](../../../../../Docs/ROADMAP.md).

The unchanged eight-case H4 selection passed all screen, print, and
screen-to-page criteria. The [machine-readable report](h4-evidence.json)
records the measurements and working-tree source state; its large per-case
image output was not retained here. The final Windows static budget passed with
a 6.9-second cold process, 2.3-second slowest warm iteration, and matching
cold/warm output fingerprints ([report](budget-windows-passed.json)). The
configured ceiling was left unchanged.
