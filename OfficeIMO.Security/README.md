# OfficeIMO.Security

`OfficeIMO.Security` is the optional cryptographic provider for OfficeIMO. It owns bounded CMS/PKCS#7, S/MIME,
RFC 3161, X.509, XML Digital Signature, and enveloped-data operations. Word, PDF, Email, and the other format
packages do not depend on this package: applications that use cryptographic features install it explicitly and pass
its strongly typed provider to the format API.

```powershell
dotnet add package OfficeIMO.Security
```

The dependency-free `IOfficeSecurityProvider` contract and result models ship in the common `OfficeIMO.Core`
foundation under the `OfficeIMO.Security` namespace. The concrete `OfficeSecurityProvider` implementation lives in
this package. This keeps normal document creation, reading, conversion, signature inspection, and safe mutation
policies free of Bouncy Castle and XML DSig dependencies while avoiding reflection, dynamic plug-in discovery, or a
new package for every format.

```csharp
using OfficeIMO.Security;

IOfficeSecurityProvider security = OfficeSecurityProvider.Default;
```

Applications remain responsible for key custody, certificate and recipient selection, and trust policy. The provider
does not silently discover keys or enable network revocation.

## Use with format packages

Pass the same provider to the format API that owns the document structure:

```csharp
using OfficeIMO.Security;
using OfficeIMO.Word;

IOfficeSecurityProvider security = OfficeSecurityProvider.Default;
WordDocument.SignPackage("report.docx", security, signingCertificate);

using WordDocument signed = WordDocument.Load("report.docx");
WordSignatureValidationReport validation = signed.ValidateSignatures(security);
```

`OfficeIMO.Pdf` accepts it through `PdfCmsExternalSigner` and
`PdfCmsSignatureCryptographyProvider`. `OfficeIMO.Email` accepts it through `EmailSmime.Verify` and
`EmailSmime.Decrypt`. Structural signature inspection and fail-safe mutation policies do not require this package.

## CMS signing and verification

The provider is also usable directly when no format adapter is involved:

```csharp
byte[] signature = security.SignCmsDetached(content, signingCertificate);
CmsVerificationResult result = security.VerifyCmsDetached(signature, content);

foreach (CmsSignerVerificationResult signer in result.Signers) {
    Console.WriteLine($"{signer.Subject}: {signer.SignatureStatus}, {signer.CertificateValidation.ChainStatus}");
}
```

Signing uses the platform `RSA` handle and does not export the private key. Verification supports RSA and ECDSA
signers and keeps mathematical signature, message digest, certificate trust, revocation, and timestamp outcomes
separate. `CreateCmsVerificationSession(...)` shares operation-wide timestamp limits across related CMS containers.

## Certificate trust validation

Use the provider when an application needs the same certificate-chain, revocation, and usage-policy result without
first parsing CMS. Additional certificates are chain-building candidates, not trusted merely because a caller
supplied them.

```csharp
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

var options = new CertificateValidationOptions {
    RevocationMode = X509RevocationMode.NoCheck,
    DisableCertificateDownloads = true
};

CertificateTrustValidationResult trust = security.ValidateCertificate(
    signingCertificate,
    additionalCertificates: new[] { intermediateCertificate },
    options: options,
    purpose: CertificateValidationPurpose.DocumentSigning);

Console.WriteLine($"Chain: {trust.Validation.ChainStatus}");
Console.WriteLine($"Revocation: {trust.Validation.RevocationStatus}");
```

The secure default disables certificate downloads and uses `X509RevocationMode.NoCheck`. Set an explicit verification
time, revocation mode, download policy, or chain evaluator when the application owns a different trust policy. Usage
and enhanced-key-usage checks remain active even when platform chain building is disabled.

## EnvelopedData and timestamps

```csharp
byte[] envelope = security.EncryptCms(content, new[] { recipientCertificate });
CmsDecryptionResult decrypted = security.DecryptCms(envelope, recipientWithPrivateKey);
```

Recipient selection is exact and caller-owned. The current Bouncy Castle key-transport adapter requires an
exportable RSA private key for envelope decryption; a non-exportable key produces the stable
`EnvelopePrivateKeyNotExportable` finding.

`VerifyTimestamp(...)` validates RFC 3161 signatures, TSA certificate profiles, message imprints, caller trust policy,
and revocation as a separate operation. TSA chain validation defaults to the token generation time unless the caller
supplies another verification time.

## XML Digital Signatures

`CreateXmlSignature(...)`, `VerifyXmlSignature(...)`, and `CanonicalizeXml(...)` expose a closed, bounded XML DSig
algorithm set for format-owned signing workflows. RSA/SHA signature methods, SHA digests, canonicalization, and
enveloped-signature transforms are intersected with immutable provider support; caller allowlists can narrow that set
but cannot register or enable another implementation. External references and unsupported transforms are rejected.
Document packages remain responsible for package relationships, content types, signed-part selection, and mutation
safety.

## C2PA Content Credentials

The dependency-free `OfficeProvenanceInspector` in `OfficeIMO.Core` finds C2PA carriers and IPTC Digital Source Type declarations. Cryptographic verification and claim signing are optional and stay in this package through the host-supplied `c2patool` adapter.

The verifier invokes an official `c2patool` executable supplied by the application. OfficeIMO does not download, bundle, or discover the executable. Remote manifest and OCSP fetching are disabled by default; local trust material can be supplied without enabling network access.

```csharp
using OfficeIMO.Provenance;
using OfficeIMO.Security;

IOfficeProvenanceVerifier verifier = new C2paToolProvenanceVerifier("/opt/c2pa/c2patool");
var options = new OfficeProvenanceVerificationOptions {
    TrustAnchorsPath = "/etc/my-app/c2pa-trust-anchors.pem",
    AllowedListPath = "/etc/my-app/c2pa-allowed-list.pem",
    IncludeRawReport = false
};

OfficeProvenanceVerificationResult result = verifier.Verify("image.jpg", options);
Console.WriteLine(result.Status);
foreach (string finding in result.Findings) {
    Console.WriteLine(finding);
}
```

`Valid` means the configured provider found a manifest, verified it, and produced no validation findings. `Untrusted` distinguishes trust-list failures from content or signature failures reported as `Invalid`. `NotPresent`, `ProviderUnavailable`, `Indeterminate`, and `Error` remain separate outcomes so callers do not have to infer policy from exception text. Set `AllowNetworkAccess = true` only when remote manifests, remote trust material, or OCSP are part of the application’s policy. Provider output is bounded and omitted from the result unless `IncludeRawReport` is enabled.

For production signing, keep private keys outside OfficeIMO and `c2patool`. Supply a subprocess signer backed by your HSM, KMS, key vault, or signing service:

```csharp
using OfficeIMO.Provenance;
using OfficeIMO.Security;

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

The adapter requires `c2patool` 0.27.0 or newer. It writes the manifest definition and provider output to temporary/staging paths, requires a separately embedded C2PA manifest before commit, and atomically installs the finished asset. A provider error or partial output cannot replace the source or an existing destination. `CreateWithBuiltInTestCredentials(...)` is an explicit development-only path and every successful result carries a warning that it is not a production credential.

For a new asset, the first action must be `Created` with a concrete IPTC Digital Source Type. For a derived asset with `parentPath`, the first action must be `Opened`. The adapter passes those intents to current c2patool, which creates the required `c2pa.actions.v2` first action and, for derived assets, the matching `parentOf` ingredient reference. Later application actions are emitted in the same v2 assertion. It does not emit any watermark action: standards-defined watermark actions require real soft-binding assertions and therefore need a future watermark provider rather than a label alone. The adapter commits only output formats that both the installed tool can sign and OfficeIMO.Core can structurally confirm.

## NativeAOT and trimming

Ordinary OfficeIMO applications do not carry this package unless they opt in. The repository publishes a separate
`OfficeIMO.Security.AotSmoke` executable that signs and verifies both CMS and XML DSig from NativeAOT. The provider
roots only its accepted XML DSig algorithms to satisfy `SignedXml`'s name-based algorithm resolution; consumers do
not need linker descriptors or reflection-based registration.

## Dependency footprint

- **External:** `BouncyCastle.Cryptography` 2.x, `System.Security.Cryptography.Xml`, and `System.Text.Json`. The optional C2PA path also requires a host-supplied `c2patool` executable.
- **OfficeIMO:** the zero-dependency `OfficeIMO.Core` foundation for provider contracts and result models.
- **Not included transitively by:** `OfficeIMO.Word`, `OfficeIMO.Pdf`, `OfficeIMO.Email`, or other format packages.

`IOfficeSecurityProvider` and its result constructors are public in the dependency-free contract assembly, so an
application can supply a policy-specific provider or test double without referencing this concrete implementation.
