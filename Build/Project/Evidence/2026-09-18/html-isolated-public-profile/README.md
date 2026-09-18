# Isolated public-page profile evidence

This evidence qualifies the public `NetworklessRootlessOciV1` workflow on a
Windows host using rootless Podman in WSL2. The Windows process acquired the
page and retained its provenance; the complete HTML, JavaScript, capture and
rendering pipeline ran in a verified networkless OCI container.

The host was Windows `10.0.26200.0` with WSL `2.7.11.0`. The isolation backend
was Ubuntu 24.04.3 LTS, kernel `6.18.33.2-microsoft-standard-WSL2`, rootless
Podman 4.9.3 and .NET SDK/runtime 10.0.112. Podman reported seccomp, cgroup v2,
and CPU, memory and PID controllers. The public renderer image was
`sha256:7e5d0cc4b80270aaa35db44b390f3bcea0c9bb4dc1a20292493b5366366ffdfb`,
built from the pinned .NET runtime base
`sha256:8a153b5889d796b6450295b383596b13308c24c230515f8a7770ce1b94e0c460`.

## Named public page

The retained [Web Platform Tests reference page](https://wpt.live/css/css-pseudo/first-letter-001-ref.html)
was fetched on 2026-09-18 at 09:17:11 UTC. WPT source is covered by the
[BSD 3-Clause license](../../2026-09-16/html-public-oci-pilot/WPT-LICENSE.md).
Input bytes were not retained. The `windows-wsl/wpt/` directory contains the
acquisition record, isolated outcome and all three outputs. The 816 x 720 PNG
was visually inspected: the text is readable, the required green rectangle is
present and no red is visible. Both PDFs reopen as one tagged A4 page.

The outcome records the immutable image, verified policy, renderer and worker
assembly hashes, complete published-directory hashes, output hashes and
confirmed container removal. This is a result for the named page and profile;
it does not claim arbitrary-site or Chromium compatibility.

## Boundary and failure probes

`controlled-summary.json` records 14 passing render cases and four passing
acquisition cases. They cover redirects, per-hop DNS rebinding rejection,
oversized-response rejection, malformed markup, responsive resources, module
graphs and import maps, dynamic GET replay, child frames, resource fanout,
capture limits and nonterminating root and child scripts.

`oci-probe.txt` records normal worker protocol, cancellation, rejected
inspection cleanup, rejected inspection with a failed first removal, and a
live-lease failed first removal followed by a successful retry. The probe used
the separate script-worker image
`sha256:b86fd0de6b006da9242a370091fa47b91a4cc5f6f81d0074803641f940f95d21`.

`windows-wsl/runtime-deadline/` records a 20-second operation deadline reached
during isolated resource discovery. The failure retained the acquired-page
provenance, exact container name, expected payload hashes and successful
cleanup. The journal evidence independently records container creation and
removal, and a subsequent existence check returned absent.

`windows-wsl/payload-mismatch/` records a deliberate one-byte change to the
locally expected renderer DLL. The workflow rejected the response, preserved
distinct expected and reported assembly and directory hashes, and confirmed
container removal.

The final managed validation passed 456 runtime/rendering tests on both .NET 8
and .NET 10, plus 3,164 broader HTML tests on each framework.
