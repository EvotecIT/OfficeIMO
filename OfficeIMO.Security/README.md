# OfficeIMO.Security

`OfficeIMO.Security` is the shared cryptographic protocol owner for OfficeIMO. It provides neutral CMS/PKCS#7,
S/MIME, RFC 3161, certificate-chain, and enveloped-data operations backed by Bouncy Castle.

It does not reference PDF, Email, Drawing, or another OfficeIMO format package. Format packages translate their own
models to and from these neutral results.

```powershell
dotnet add package OfficeIMO.Security
```

The package deliberately owns protocol orchestration, validation policy, bounded parsing, and stable result models.
It does not invent cryptographic primitives, manage certificate stores, select recipients, or silently use network
revocation services. Applications remain responsible for key custody and trust policy.

## CMS signing and verification

```csharp
using OfficeIMO.Security;

byte[] signature = CmsSignedDataSigner.SignDetached(content, signingCertificate);
CmsVerificationResult result = CmsSignedDataVerifier.VerifyDetached(signature, content);

foreach (CmsSignerVerificationResult signer in result.Signers) {
    Console.WriteLine($"{signer.Subject}: {signer.SignatureStatus}, {signer.CertificateValidation.ChainStatus}");
}
```

Signing uses the platform `RSA` handle and does not export the private key. Verification supports RSA and ECDSA
signers, keeps mathematical signature, message digest, certificate trust, revocation, and timestamp outcomes
separate, and never enables network revocation implicitly.

## Certificate trust validation

Use `CertificateValidator.Validate` when an application needs the same certificate-chain, revocation, and usage-policy
result without first parsing CMS. Supply issuer certificates carried by the document or protocol through
`additionalCertificates`; they are added as chain-building candidates, not treated as trusted merely because they were
provided.

```csharp
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

var options = new CertificateValidationOptions {
    RevocationMode = X509RevocationMode.NoCheck,
    DisableCertificateDownloads = true
};

CertificateTrustValidationResult trust = CertificateValidator.Validate(
    signingCertificate,
    additionalCertificates: new[] { intermediateCertificate },
    options: options,
    purpose: CertificateValidationPurpose.DocumentSigning);

Console.WriteLine($"Chain: {trust.Validation.ChainStatus}");
Console.WriteLine($"Revocation: {trust.Validation.RevocationStatus}");
foreach (SecurityFinding finding in trust.Findings) {
    Console.WriteLine($"{finding.Code}: {finding.Message}");
}
```

The secure default disables certificate downloads and uses `X509RevocationMode.NoCheck`. Set an explicit
`VerificationTime`, revocation mode, download policy, or `ChainEvaluator` when the application owns a different trust
policy. Use `CertificateValidationPurpose.TimestampAuthority` for TSA certificates; usage and enhanced-key-usage checks
remain active even when platform chain building is disabled.

## EnvelopedData

```csharp
byte[] envelope = CmsEnvelopedDataService.Encrypt(content, new[] { recipientCertificate });
CmsDecryptionResult decrypted = CmsEnvelopedDataService.Decrypt(envelope, recipientWithPrivateKey);
```

Recipient selection is exact and caller-owned. The current Bouncy Castle key-transport adapter requires an exportable
RSA private key for envelope decryption; a non-exportable key produces the stable
`EnvelopePrivateKeyNotExportable` finding instead of silently falling back or exporting key material elsewhere.

`Rfc3161TimestampVerifier` validates timestamp signatures, TSA certificate profiles, message imprints, caller trust
policy, and revocation as a separate neutral operation. TSA chain validation defaults to the token generation time;
callers can override that instant through `CertificateValidationOptions.VerificationTime`.

## Dependency footprint

- **External:** `BouncyCastle.Cryptography` 2.x.
- **OfficeIMO:** None. PDF and Email depend on Security; Security never depends on a format package.
