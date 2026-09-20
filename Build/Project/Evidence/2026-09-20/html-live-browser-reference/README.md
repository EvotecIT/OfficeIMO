# Offline browser-reference qualification

This bundle qualifies four Web Platform Tests pages with immutable,
same-byte browser comparisons. The sources are covered by the included
[BSD 3-Clause license](WPT-LICENSE.md).

`HtmlPublicPilot` retained the exact acquired HTML and rendered it inside the
qualified rootless, network-disabled Podman profile. Each `outcome.json` records
image ID
`sha256:bc8238082eade5ffe42ac0cecb15c16603144dab10b470c12041a87e54043341`,
renderer and script-worker hashes, isolation controls, trace, output hashes, and
confirmed container removal. The comparison-only runner then intercepted every
Chromium request and fulfilled it from that retained acquisition. An unrecorded
request fails the run; all four evidence files record an empty `blockedUrls`
collection.

The Chromium reference used Windows x64 Chromium 151.0.7922.34 through
Playwright 1.62.0.0 and HtmlTinkerX 3.0.1.0. OfficeIMO used the image's declared
generic-serif fallback, DejaVu Serif. Each `browser-evidence.json` records host,
framework, viewport, acquisition, source, browser, and image hashes together
with decoded-pixel metrics.

| Reference | Retained source | Observed result |
| --- | --- | --- |
| [Blue line on gray SVG](https://wpt.live/css/compositing/line-with-svg-background-ref.html) | 2,017 bytes; `38603d5eef8f492535cb706b065474dcc21cf6866b70520576ab91de733120fe` | Both screens are 816 x 720 with zero differing pixels. CSS sizing, centering, gray backgrounds, and blue strokes match. [Evidence](named-wpt-svg-background/browser-reference/browser-evidence.json) |
| [Text on blue SVG](https://wpt.live/css/compositing/Text_with_SVG_background-ref.html) | 451 bytes; `926205a6959a09c681c03132ac8fb5f14083729ed6489f4472278a868c07c9f5` | Both screens are 816 x 720. The browser body inset and SVG box geometry match. Remaining font rasterization gives MAE 0.0889 and RMSE 4.0826. [Evidence](named-wpt-svg-text-blue/browser-reference/browser-evidence.json) |
| [Text on gray SVG rows](https://wpt.live/css/compositing/text-with-svg-background-ref.html) | 1,899 bytes; `b109c5db621bf08782d9ad41902078f308376fb9ee59f1ff5d1217f89edfb166` | Both full-page screens are 816 x 780 and every SVG box matches. DejaVu Serif versus Times New Roman glyph geometry gives MAE 3.5430 and RMSE 19.8239. [Evidence](named-wpt-svg-text-gray/browser-reference/browser-evidence.json) |
| [Scripted details mutation](https://wpt.live/html/semantics/interactive-elements/the-details-element/details-add-summary.html) | 819 bytes; `e070e160be18d81e58d743840bedd8a92e148248c2b1d1aa702be9b5148634a3` | Both 816 x 720 screens contain `new summary`, which exists only after the inline `load` handler creates a `summary`, assigns `textContent`, finds both nodes, and calls `insertBefore`. OfficeIMO captures and renders that live DOM revision. Its body inset and text position match; Chromium additionally paints the native closed-details disclosure marker, which is outside the current static renderer subset. Font and marker rasterization give MAE 0.1234 and RMSE 4.7739. [Evidence](named-wpt-script-dom/browser-reference/browser-evidence.json) |

OfficeIMO now selects browser-like user-agent defaults explicitly for this
application profile. Authored CSS can still reset the body margin, and the SVG
drawing reader receives the same default font family as HTML text. The font
difference is retained as a platform-font provenance classification; supplying
the same licensed font bytes to both renderers is the route to deterministic
glyph identity.

These results prove the named cases. They do not establish general SVG, CSS,
JavaScript, DOM, form-control, user-agent-widget, or browser compatibility.
