# Optional Linux OCI worker probe

This opt-in runner exercises the internal Podman worker lease, not a public-page
API. It requires rootless Podman with working CPU, memory and PID cgroups. The
Linux x64 container file pins the runtime base used in the recorded probe.

From the repository root on Linux, publish the worker into a task-owned build
context, build the image, and run the probe with its immutable image ID:

```bash
mkdir -p /tmp/officeimo-html-oci-probe/worker
dotnet publish OfficeIMO.Html.Runtime.Worker/OfficeIMO.Html.Runtime.Worker.csproj -c Release -f net10.0 -o /tmp/officeimo-html-oci-probe/worker
podman build -f Build/Project/HtmlOciProbe/Containerfile.linux-x64 -t localhost/officeimo-html-oci-probe:local /tmp/officeimo-html-oci-probe
image_id="$(podman image inspect --format '{{.Id}}' localhost/officeimo-html-oci-probe:local)"
dotnet run --project Build/Project/HtmlOciProbe/OfficeIMO.Html.OciProbe.csproj -c Release -- "sha256:${image_id}" /tmp/officeimo-html-oci-probe/worker/OfficeIMO.Html.Runtime.Worker.dll
```

The runner verifies a normal worker capture, a cancelled runaway script, a
post-create inspection rejection, that rejection combined with a failed first
removal, and a live-lease first-removal failure followed by a successful retry.
Each case checks that no task-named container remains after the lease closes.
The failure cases use the runner itself as a Podman CLI proxy to inject a
deliberately rejected inspection or one failed removal.
The worker image, build context and base-image cache are not removed
automatically. Remove only the exact task-owned build context and derived
image when finished; keep shared caches intact.

The internal lease checks that Podman reports rootless operation and that
container inspection retains the declared read-only filesystem, no network,
non-root user, no host mounts, private PID namespace, no-new-privileges flag,
dropped default capabilities, and CPU/memory/PID limits. It requires Podman's
seccomp support and all three cgroup controllers. This does not yet prove that
public content is safe:
the worker image identity, child-process restrictions, host resource broker,
cross-platform OCI backends and hostile-input corpus remain open work in
`Docs/ROADMAP.md`.
