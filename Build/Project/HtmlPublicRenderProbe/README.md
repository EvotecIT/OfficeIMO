# Isolated renderer fixture probe

This opt-in Linux probe sends generated HTML fixtures through the same rootless,
network-disabled Podman image as the [public-page pilot](../HtmlPublicPilot/README.md).
It exercises the whole parser, JavaScript, capture and renderer pipeline without
fetching public pages or creating another public runtime API.

Publish the script worker and renderer into separate directories, then build the
image as described in the public-page pilot. Run the probe with that image's
full SHA-256 ID and the exact published DLL paths:

```bash
dotnet run --project Build/Project/HtmlPublicRenderProbe/OfficeIMO.Html.PublicRenderProbe.csproj -c Release -- \
  "sha256:${image_id}" \
  /tmp/officeimo-public-pilot/renderer/OfficeIMO.Html.PublicRenderWorker.dll \
  /tmp/officeimo-public-pilot/worker/OfficeIMO.Html.Runtime.Worker.dll \
  /tmp/officeimo-public-pilot/fixture-output
```

The output directory must be new. Each case starts a separate verified
container and checks removal after completion. The generated cases cover
malformed markup that still renders all three outputs, 129 distinct external
script URLs exceeding static discovery, a captured-document size limit, and a
nonterminating script interrupted by its command deadline. `summary.json`
records the image and published-file digests, each outcome, elapsed time, and
container removal. The successful case also writes a PNG and both PDFs for
visual and independent PDF inspection. This is controlled fixture evidence;
public acquisition, redirects, DNS changes, and cross-platform containment
need separate tests.
