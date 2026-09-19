# Isolated dynamic request and named-page evidence

This run exercised the `NetworklessRootlessOciV1` public-page profile on a
Windows host with Ubuntu WSL2, rootless Podman 4.9.3 and .NET 10.0.112. The
renderer image was
`sha256:c7ecddfc10d7bf12d762fdfd2f4f9652146ee49da65733d474e7fa9143b144e6`.
The image, published renderer and script-worker hashes, and every container
removal are recorded in [controlled-summary.json](controlled-summary.json) and
the named-page outcome files. The container had no network route or host
mounts; only the host broker acquired public bytes.

## Controlled end-to-end results

The [controlled summary](controlled-summary.json) records 22 passing render
cases and six passing acquisition cases. The new cases cover an acquired POST
redirected to GET, a cross-origin POST preceded by OPTIONS, their exact offline
replay through parsing, scripting, capture and all three render intents, and a
strict encoded-output limit. The acquisition cases used synthetic public DNS
answers and a loopback fixture transport; each direct request was revalidated
by the broker. Their evidence is controlled, not proof that arbitrary public
sites or origin-changing redirects work.

The [redirect](acquisition-dynamic-redirect/screen.png) and
[CORS](acquisition-dynamic-cors/screen.png) screens visibly contain the
script-produced blue result text. Their print and screen-to-page PDFs retain
that text. The synthetic [replay redirect](dynamic-post-redirect/screen.png) and
[replay preflight](dynamic-cross-origin-preflight/screen.png) cases exercise the
same worker with explicitly supplied transcripts. The encoded-output case
rejected a valid page after rendering because its one-byte per-artifact limit
was exceeded; its container was removed.

## Named public pages

All three pages below were acquired from Web Platform Tests on 2026-09-19.
WPT sources are covered by its retained [BSD 3-Clause license](../../2026-09-16/html-public-oci-pilot/WPT-LICENSE.md).
The evidence retains acquisition metadata and output bytes, not source-page
bytes. Each output includes an 816 × 720 PNG and tagged, single-page A4 PDFs
for print and screen-to-page.

| Page | Observed screen result | Evidence |
| --- | --- | --- |
| [First-letter reference](https://wpt.live/css/css-pseudo/first-letter-001-ref.html) | Readable instruction, filled green rectangle, no visible red | [PNG](named-wpt-first-letter/screen.png), [print PDF](named-wpt-first-letter/print.pdf), [screen-to-page PDF](named-wpt-first-letter/screen-to-page.pdf), [acquisition](named-wpt-first-letter/acquisition.json) |
| [Cascade reference](https://wpt.live/css/css-cascade/initial-color-background-001-ref.html) | Readable instruction and large black W, no visible red | [PNG](named-wpt-cascade/screen.png), [print PDF](named-wpt-cascade/print.pdf), [screen-to-page PDF](named-wpt-cascade/screen-to-page.pdf), [acquisition](named-wpt-cascade/acquisition.json) |
| [Inline SVG reference](https://wpt.live/css/compositing/line-with-svg-background-ref.html) | Outlined SVG boxes appear, but the authored gray backgrounds and blue lines are missing | [PNG](named-wpt-svg-background/screen.png), [print PDF](named-wpt-svg-background/print.pdf), [screen-to-page PDF](named-wpt-svg-background/screen-to-page.pdf), [acquisition](named-wpt-svg-background/acquisition.json) |

The first two screens match the visible pass markers authored by their WPT
references; both PDF intents reopen and retain their text. The SVG page is a
specific rendering gap, not a pass. Its source defines gray SVG backgrounds and
blue strokes, but those paints are absent from the retained PNG. The work to
qualify those paints remains in the [roadmap](../../../../../Docs/ROADMAP.md).

The [one-byte public output-limit failure](named-wpt-output-limit-final/failure.json)
records phase `Output`, the effective limit, the worker's byte-budget error and
confirmed container removal. It used the first-letter page and the same
immutable image. No output was returned as a successful artifact.
