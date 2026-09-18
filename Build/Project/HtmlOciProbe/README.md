# Optional OCI worker probe

This opt-in runner exercises the internal Podman worker lease, not a public-page
API. It requires rootless Podman with working CPU, memory and PID cgroups. The
multi-architecture Linux container file pins the runtime base used in recorded
Linux and macOS/AppleHV probes.

From the repository root on Linux, publish the worker into a task-owned build
context, build the image, and run the probe with its immutable image ID:

```bash
mkdir -p /tmp/officeimo-html-oci-probe/worker
dotnet publish OfficeIMO.Html.Runtime.Worker/OfficeIMO.Html.Runtime.Worker.csproj -c Release -f net10.0 -r linux-x64 --self-contained false -o /tmp/officeimo-html-oci-probe/worker
podman build --platform linux/amd64 -f Build/Project/HtmlOciProbe/Containerfile.linux -t localhost/officeimo-html-oci-probe:local /tmp/officeimo-html-oci-probe
image_id="$(podman image inspect --format '{{.Id}}' localhost/officeimo-html-oci-probe:local)"
dotnet run --project Build/Project/HtmlOciProbe/OfficeIMO.Html.OciProbe.csproj -c Release -- "sha256:${image_id}" /tmp/officeimo-html-oci-probe/worker/OfficeIMO.Html.Runtime.Worker.dll
```

On Apple Silicon macOS, first start a rootless Podman machine, publish the
worker with `-r linux-arm64 --self-contained false`, and build with
`--platform linux/arm64`. Run the probe itself as a macOS host process so its
Podman subprocess uses the active machine connection.

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
seccomp support and all three cgroup controllers. The public-page workflow adds
immutable image and payload identities, the host resource broker,
full-pipeline hostile-input probes and named-page evidence; this worker probe
alone is not that claim.
