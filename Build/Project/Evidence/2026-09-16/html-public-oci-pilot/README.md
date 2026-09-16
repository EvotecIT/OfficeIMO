# Linux public-page pilot: two capability results

The retained [Web Platform Tests reference page](https://wpt.live/css/css-pseudo/first-letter-001-ref.html)
was fetched on 2026-09-16 at 12:43:24 UTC. Its exact HTML and acquisition
metadata are under `wpt-first-letter/`. WPT source is covered by the included
[BSD 3-Clause license](WPT-LICENSE.md). This page needed no external assets.
The [pilot runner](../../../HtmlPublicPilot/README.md) sent its bytes to a
rootless, network-disabled Podman container, where the entire parse, script,
capture and render workflow ran. The image ID was
`sha256:082e2355ecdbf0cb06f6ccc1a316eb676a3031bc964a4cbd82ca4bde0f54cbc1`.
`outcome.json` records verified isolation controls, matching renderer/worker
DLL and complete published-file-set hashes, trace entries, output digests and
confirmed container removal. `worker-mismatch/failure.json` records a negative
run where an extra file in the locally expected worker payload made the
file-set digest differ; the result was rejected and the container was removed.

The three retained outputs are an 816 × 110 screen PNG, one-page print PDF and
one-page screen-to-page PDF. The screen PNG was visually inspected: text is
readable and the green rectangle is present with no red. This is a successful
end-to-end result for one small static page. It is not a pixel-match claim
against Chromium or general website compatibility.

The live [OfficeIMO homepage](https://officeimo.com/) was fetched on
2026-09-16 at 12:45:35 UTC: 69,306 HTML bytes with SHA-256
`b51373b1c45304c38b53b995ae52c4be746adf5636e2b5cb475c795a7b9271cf`
and 19 explicitly supplied direct assets. `officeimo-homepage/acquisition.json`
and `failure.json` retain its provenance and failure classification; the site
HTML and assets are not redistributed here. Inside the same isolated image,
script execution stopped before capture with “The string did not match the
expected pattern.” The container was removed. This is a failed site case, not
a PDF or appearance result.

The adjacent `html-oci-linux-isolation-probe` covers normal worker capture,
cancellation, failed inspection and removal retry. Broader hostile-input
recovery, automatic resource discovery, browser reference comparisons and
Windows/macOS isolation remain open in `Docs/ROADMAP.md`. The internal pilot
does not establish the public untrusted-content profile.
