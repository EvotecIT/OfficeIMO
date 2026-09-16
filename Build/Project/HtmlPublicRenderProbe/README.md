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

The output directory must be new. Each render case starts a separate verified
container and checks removal after completion; acquisition rejection cases stop
before container startup. The generated cases cover
malformed markup that still renders all three outputs, active multi-candidate
responsive-picture selection by viewport, `sizes`, and device density, a two-level JavaScript module graph,
a scoped import-map graph with prefix mapping, top-level await and dynamic import,
a static frame document plus its relative stylesheet and script, 129 distinct
external script URLs exceeding static discovery, a captured-document size limit, and a
nonterminating script interrupted by its command deadline. Frame readiness proves
nested document loading and proves inline, external and event-attribute scripts
remain inert. Child-frame execution realms and frame-body capture/rendering are
outside this fixture contract. `summary.json`
records the image and published-file digests, ordered discovery rounds with
canonical absolute resource URLs, each outcome, elapsed time, and container
removal. The module fixture requires the root module in the first round and its
relative dependency in the second. Successful cases also write
a PNG and both PDFs for visual and independent PDF inspection.

The same run uses the host broker with synthetic public DNS answers and a
loopback transport to acquire same-host and explicitly approved cross-host
redirect documents before rendering them in separate network-disabled
containers. It also proves that a per-hop public-to-loopback DNS change and an
oversized declared response are rejected before container startup. The summary
records requested and final URLs, resolutions, validated connection endpoints,
HTTP redirect hops and fixture requests. This is controlled fixture evidence;
mutable public DNS, live public-host redirects and cross-platform containment
need separate qualification.
