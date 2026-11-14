---
title: "Signatures and Encryption"
description: "Sign and verify documents and data, validate certificates and encrypt for recipients with OfficeIMO.Security, the optional CMS, S/MIME, RFC 3161, X.509 and XML signature provider."
order: 59
verified: 2026-11-14
meta.seo_title: "Sign, verify and encrypt documents in C#"
packages:
  - OfficeIMO.Security
related:
  - /docs/security/sign-documents/
  - /docs/security/cms/
  - /docs/pdf/encrypt-sign/
  - /docs/email/safety-and-verification/
limits: Key custody, certificate and recipient selection and trust policy belong to your application; the provider does not search certificate stores or enable network revocation on its own.
search_terms:
  - digital signature
  - cms
  - pkcs7
  - smime
  - x509
  - encrypt document
---

`OfficeIMO.Security` is the optional cryptographic provider for OfficeIMO. It implements bounded CMS/PKCS#7 signing and verification, CMS EnvelopedData encryption, RFC 3161 timestamps, X.509 trust validation and XML Digital Signatures. Format packages such as Word, PDF and Email do not depend on it: install it when you need cryptography and pass the provider to the format API.

```sh
dotnet add package OfficeIMO.Security
```

It targets .NET Standard 2.0, .NET 8, .NET 10 and .NET Framework 4.7.2, and depends on `BouncyCastle.Cryptography` 2.x and `System.Security.Cryptography.Xml`. The `IOfficeSecurityProvider` contract and its result models live in the dependency-free `OfficeIMO.Core`, so structural signature inspection and fail-safe mutation policies work without this package.

| Task | Where |
|---|---|
| Sign and verify DOCX, XLSX, PPTX, VSDX and ODT/ODS/ODP packages | [Sign and verify documents](/docs/security/sign-documents/) |
| Sign, verify and encrypt raw data with CMS; validate certificates; timestamps | [CMS signatures and encryption](/docs/security/cms/) |
| Sign and verify PDFs | [PDF encryption and signing](/docs/pdf/encrypt-sign/), through `PdfCmsExternalSigner` and `PdfCmsSignatureCryptographyProvider` |
| Verify and decrypt S/MIME email | [Email safety and verification](/docs/email/safety-and-verification/), through `EmailSmime.Verify` and `EmailSmime.Decrypt` |

Password encryption for Office, PDF and OpenDocument files is owned by each format package and does not need this provider.

## Pass the provider to a format package

Create the provider once and hand it to the API that owns the document. The format package selects the signed parts and keeps mutation safe; the provider does the cryptography.

### C#

```csharp
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;
using OfficeIMO.Word;

IOfficeSecurityProvider security = OfficeSecurityProvider.Default;

using (WordDocument document = WordDocument.Create("report.docx")) {
    document.AddParagraph("Quarterly report.");
    document.Save();
}

// A throwaway test certificate; load your real signing certificate the way your application manages keys.
using RSA rsa = RSA.Create(2048);
var request = new CertificateRequest("CN=Report Signing Test", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));
using X509Certificate2 certificate = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));

WordDocument.SignPackage("report.docx", security, certificate);

using WordDocument signed = WordDocument.Load("report.docx");
WordSignatureValidationReport validation = signed.ValidateSignatures(security);
Console.WriteLine($"Signature math: {validation.CryptographicStatus}, signer trust: {validation.CertificateChainStatus}");
```

With default options the test certificate's chain is reported as untrusted, because no root vouches for it. [Choose a trust policy](/docs/security/sign-documents/#choose-a-trust-policy) shows how to pin a certificate or a private root.

## Sign and verify bytes directly

### C#

```csharp
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

IOfficeSecurityProvider security = OfficeSecurityProvider.Default;

using ECDsa key = ECDsa.Create(ECCurve.NamedCurves.nistP256);
var request = new CertificateRequest("CN=Payload Signing Test", key, HashAlgorithmName.SHA256);
request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));
using X509Certificate2 certificate = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));

byte[] content = File.ReadAllBytes("payload.bin");
byte[] signature = security.SignCmsDetached(content, certificate);
CmsVerificationResult result = security.VerifyCmsDetached(signature, content);

Console.WriteLine(result.IsCryptographicallyValid ? "Content matches the signature" : "Content or signature changed");
```

RSA and ECDSA signers are supported. See [CMS signatures and encryption](/docs/security/cms/) for trust options, encapsulated signatures, recipient encryption and timestamps.

## Know what the provider covers

- **CMS/PKCS#7:** detached and encapsulated SignedData, including baseline CMS for PDF signatures; verification with separate signature, digest, trust, revocation and timestamp outcomes.
- **EnvelopedData:** AES-128/192/256-CBC content for RSA recipients. On .NET 8 and later, RSA PKCS#1 and OAEP recipients decrypt through the platform key handle, so non-exportable keys work.
- **RFC 3161:** timestamp requests through a transport callback you supply, and token verification.
- **XML DSig:** `CreateXmlSignature`, `VerifyXmlSignature` and `CanonicalizeXml` over a closed set of RSA/SHA signature, digest, canonicalization and enveloped-signature algorithms; external references and other transforms are rejected.
- **NativeAOT:** a separate smoke executable signs and verifies CMS and XML DSig under NativeAOT; consumers need no linker descriptors.
