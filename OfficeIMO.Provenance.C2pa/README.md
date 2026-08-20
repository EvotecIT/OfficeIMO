# OfficeIMO.Provenance.C2pa

`OfficeIMO.Provenance.C2pa` is the optional process adapter between OfficeIMO's provider-neutral provenance contracts and the official `c2patool` command-line application. Install it only in applications that need cryptographic C2PA Content Credential verification or signing.

```shell
dotnet add package OfficeIMO.Provenance.C2pa
```

Structural inspection, lifecycle policy, claims, Unicode evidence, and carrier removal remain in `OfficeIMO.Core` and the format packages. This package does not depend on `OfficeIMO.Security`, and ordinary OfficeIMO format packages do not depend on this package.

## Supply c2patool explicitly

OfficeIMO does not download, bundle, or auto-discover `c2patool`. Install a supported build from the [official c2pa-rs releases](https://github.com/contentauth/c2pa-rs/releases), then supply either an absolute executable path or a command name already available through your application's controlled `PATH`:

On Unix hosts, the adapter also requires `setsid` so the external tool and any descendants remain in a process group that OfficeIMO can terminate on timeout or disposal. Linux distributions normally provide it through `util-linux`. On macOS, install the Homebrew `util-linux` formula (`brew install util-linux`); the adapter recognizes its standard Apple Silicon and Intel keg paths even when the formula is not linked into `PATH`. If `setsid` is unavailable, verification and signing fail closed without launching `c2patool`.

```csharp
using OfficeIMO.Provenance;
using OfficeIMO.Provenance.C2pa;

IOfficeProvenanceVerifier verifier = new C2paToolProvenanceVerifier("c2patool");
OfficeProvenanceVerificationResult result = verifier.Verify("image.jpg");

Console.WriteLine(result.Status);
foreach (string finding in result.Findings) {
    Console.WriteLine(finding);
}
```

Use an absolute path in services, build agents, and other controlled deployments where executable selection must not depend on the ambient environment.

Remote manifest and OCSP fetching are disabled by default. Local trust material can be supplied without enabling network access:

```csharp
var options = new OfficeProvenanceVerificationOptions {
    TrustAnchorsPath = "/etc/my-app/c2pa-trust-anchors.pem",
    AllowedListPath = "/etc/my-app/c2pa-allowed-list.pem",
    IncludeRawReport = false
};

OfficeProvenanceVerificationResult result = verifier.Verify("image.jpg", options);
```

`Valid` means the configured provider found a manifest, verified it, and produced no validation findings. `Untrusted` distinguishes trust-list failures from content or signature failures reported as `Invalid`. `NotPresent`, `ProviderUnavailable`, `Indeterminate`, and `Error` remain separate outcomes. Provider output is bounded and omitted unless `IncludeRawReport` is enabled.

## Production signing

The signer requires `c2patool` 0.27.0 or newer. Keep private keys outside OfficeIMO and `c2patool`; supply a signer subprocess backed by your HSM, KMS, key vault, or signing service:

```csharp
IOfficeProvenanceSigner signer = new C2paToolProvenanceSigner(
    executablePath: "/opt/c2pa/c2patool",
    signerPath: "/opt/my-app/c2pa-kms-signer --profile production");

var claim = new OfficeProvenanceClaim(
    "OfficeIMO/3.2.4",
    new[] {
        new OfficeProvenanceAction(OfficeProvenanceActionKind.Opened),
        new OfficeProvenanceAction(
            OfficeProvenanceActionKind.Edited,
            OfficeProvenanceDigitalSourceKind.CompositeWithTrainedAlgorithmicMedia)
    },
    title: "Edited image");

OfficeProvenanceSigningResult signed = signer.Sign(
    new OfficeProvenanceSigningRequest(
        inputPath: "edited.png",
        outputPath: "signed.png",
        claim: claim,
        parentPath: "original.png"));
```

The adapter writes manifests and provider output to temporary staging paths, requires an embedded C2PA manifest before commit, and atomically installs the finished asset. A provider error or partial output cannot replace the source or an existing destination. `CreateWithBuiltInTestCredentials(...)` is an explicit development-only path; its successful results carry a warning that they are not production credentials.

New claims begin with a concrete `Created` intent. Parent-derived claims begin with `Opened`, and c2patool creates the matching `parentOf` ingredient reference. Later actions use `c2pa.actions.v2`. The adapter does not emit a watermark action because those actions require matching soft-binding assertions that OfficeIMO does not currently create.

## Dependency footprint

- **NuGet:** `OfficeIMO.Core` and `System.Text.Json`.
- **Runtime tool:** a host-supplied official `c2patool` executable; never downloaded or bundled by this package.
- **Not required by:** `OfficeIMO.Security`, `OfficeIMO.Word`, `OfficeIMO.Pdf`, `OfficeIMO.Email`, or other format packages.
