---
title: "Signatures and Encryption"
description: "Optional CMS, S/MIME, RFC 3161 timestamp, X.509 trust and XML Digital Signature provider for signing, verifying and encrypting with OfficeIMO packages."
order: 59
meta.seo_title: "Sign, verify and encrypt documents in C# with OfficeIMO.Security"
---

`OfficeIMO.Security` is the optional cryptographic provider for OfficeIMO. It implements bounded CMS/PKCS#7 signing and verification, CMS EnvelopedData encryption, RFC 3161 timestamps, X.509 trust validation and XML Digital Signatures. Format packages such as Word, PDF and Email do not depend on it: install it when you need cryptography and pass the provider to the format API.

```shell
dotnet add package OfficeIMO.Security
```

Targets .NET Standard 2.0, .NET 8, .NET 10 and .NET Framework 4.7.2. It depends on `BouncyCastle.Cryptography` 2.x and `System.Security.Cryptography.Xml`. The `IOfficeSecurityProvider` contract and result models live in the dependency-free `OfficeIMO.Core`, so structural signature inspection works without this package.

## Sign and validate a Word package

```csharp
using OfficeIMO.Security;
using OfficeIMO.Word;

// signingCertificate is an X509Certificate2 with an accessible private key.
IOfficeSecurityProvider security = OfficeSecurityProvider.Default;
WordDocument.SignPackage("report.docx", security, signingCertificate);

using WordDocument signed = WordDocument.Load("report.docx");
WordSignatureValidationReport validation = signed.ValidateSignatures(security);
```

`OfficeIMO.Pdf` accepts the provider through `PdfCmsExternalSigner` and `PdfCmsSignatureCryptographyProvider`, and `OfficeIMO.Email` through `EmailSmime.Verify` and `EmailSmime.Decrypt`. The format package owns document structure, signed-part selection and mutation safety.

## CMS signing and verification

```csharp
byte[] content = File.ReadAllBytes("payload.bin");
byte[] signature = security.SignCmsDetached(content, signingCertificate);
CmsVerificationResult result = security.VerifyCmsDetached(signature, content);

foreach (CmsSignerVerificationResult signer in result.Signers)
    Console.WriteLine($"{signer.Subject}: {signer.SignatureStatus}, {signer.CertificateValidation.ChainStatus}");
```

Signing uses the certificate's platform `RSA` or `ECDsa` key without exporting it. `CmsSigningOptions.UsePadesBaseline` produces baseline CMS with a SHA-256 signing-certificate-v2 binding. Verification reports the mathematical signature, message digest, certificate trust, revocation and timestamp outcomes separately.

## Certificate trust

```csharp
using System.Security.Cryptography.X509Certificates;

var options = new CertificateValidationOptions {
    RevocationMode = X509RevocationMode.NoCheck,
    DisableCertificateDownloads = true
};
CertificateTrustValidationResult trust = security.ValidateCertificate(
    signingCertificate,
    additionalCertificates: new[] { intermediateCertificate },
    options: options,
    purpose: CertificateValidationPurpose.DocumentSigning);
Console.WriteLine($"Chain: {trust.Validation.ChainStatus}, revocation: {trust.Validation.RevocationStatus}");
```

Additional certificates are chain-building candidates, not trusted anchors. The defaults disable certificate downloads and revocation checks; set a revocation mode, verification time or chain evaluator when your application owns a different trust policy. Key-usage checks stay active even without platform chain building.

## Encryption and timestamps

```csharp
byte[] envelope = security.EncryptCms(content, new[] { recipientCertificate });
CmsDecryptionResult decrypted = security.DecryptCms(envelope, recipientWithPrivateKey);
```

On .NET 8 and later, RSA PKCS#1 and OAEP recipients with AES-CBC content decrypt through the platform key handle, so non-exportable keys work. Other profiles and older targets use Bouncy Castle and need an exportable RSA key.

`VerifyTimestamp` validates RFC 3161 tokens, message imprints and TSA certificates. `Rfc3161TimestampClient` requests timestamps through a transport callback you supply, so your HTTP client owns timeouts and response limits. `CreateXmlSignature`, `VerifyXmlSignature` and `CanonicalizeXml` cover a closed set of RSA/SHA signature, digest, canonicalization and enveloped-signature algorithms; external references and other transforms are rejected.

## Limits

- Key custody, certificate and recipient selection and trust policy belong to your application. The provider does not discover keys or search certificate stores for you.
- Network revocation and certificate downloads are off unless you enable them.
- XML DSig algorithms cannot be extended beyond the built-in set.
- A separate NativeAOT smoke executable signs and verifies CMS and XML DSig; consumers need no linker descriptors.

See the [package README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Security/README.md), [PDF security](/docs/pdf/security/) and [email safety and verification](/docs/email/safety-and-verification/).
