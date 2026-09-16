# Linux OCI worker isolation probe

This probe tested whether the existing OfficeIMO HTML runtime worker can speak
its normal command protocol inside a rootless OCI container. It is a feasibility
result for the separate untrusted-content profile, not an implementation or
security qualification of that profile.

The source was `feature/html-engine` commit `7e21ff8795`. On Ubuntu 24.04.3
under WSL2, rootless Podman 4.9.3 used `crun` and a local Linux x64 image built
from `mcr.microsoft.com/dotnet/runtime@sha256:8a153b5889d796b6450295b383596b13308c24c230515f8a7770ce1b94e0c460`.
The worker was published from the same commit with .NET 10. The derived image
ran as UID/GID 65532, contained only the published worker and its runtime base,
and exposed the worker protocol on standard input and output.

The probe launched it with these OCI controls:

```text
--network none --read-only --pids-limit 32 --memory 256m --cpus 1
--cap-drop all --security-opt no-new-privileges --user 65532:65532
```

Inside the container, cgroup readback showed a 268435456-byte memory limit,
`pids.max` of 32 and `cpu.max` of `100000 100000`. A write to the read-only
filesystem failed, the host checkout path was absent, and an Alpine control
image under the same network policy could not make an outbound request.
A `podman create -i` / `podman start --attach --interactive` round trip also
returned an input line over standard output. With the OfficeIMO worker image,
a trusted sample document completed through `HtmlProcessRuntimeProvider` via a
temporary Podman wrapper and returned `Isolated worker`.

The temporary wrapper is **not** a product launcher. It did not own an exact
container ID across cancellation, prove container deletion after a killed
attach process, prevent child process execution inside the container, or
produce an isolation report. No public URL or hostile script was admitted.
The image is Linux x64 specific; Windows and macOS OCI backends were not
available on the validation hosts. Network acquisition, DNS/IP/redirect
policy, resource brokering, credentials and retained page provenance remain
unimplemented for the public-site pilot. The existing trusted process
provider and `WebApplicationV1` must continue to be used only with explicitly
trusted content.
