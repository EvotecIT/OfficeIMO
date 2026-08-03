# OfficeIMO.Security

`OfficeIMO.Security` is the optional cryptographic provider for OfficeIMO. It owns bounded CMS/PKCS#7, S/MIME,
RFC 3161, X.509, XML Digital Signature, and enveloped-data operations. Word, PDF, Email, and the other format
packages do not depend on this package: applications that use cryptographic features install it explicitly and pass
its strongly typed provider to the format API.

```powershell
dotnet add package OfficeIMO.Security
```

The dependency-free `IOfficeSecurityProvider` contract and result models ship in the common `OfficeIMO.Drawing`
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

## NativeAOT and trimming

Ordinary OfficeIMO applications do not carry this package unless they opt in. The repository publishes a separate
`OfficeIMO.Security.AotSmoke` executable that signs and verifies both CMS and XML DSig from NativeAOT. The provider
roots only its accepted XML DSig algorithms to satisfy `SignedXml`'s name-based algorithm resolution; consumers do
not need linker descriptors or reflection-based registration.

## Dependency footprint

- **External:** `BouncyCastle.Cryptography` 2.x and `System.Security.Cryptography.Xml`.
- **OfficeIMO:** the zero-dependency `OfficeIMO.Drawing` foundation for provider contracts and result models.
- **Not included transitively by:** `OfficeIMO.Word`, `OfficeIMO.Pdf`, `OfficeIMO.Email`, or other format packages.

`IOfficeSecurityProvider` and its result constructors are public in the dependency-free contract assembly, so an
application can supply a policy-specific provider or test double without referencing this concrete implementation.
