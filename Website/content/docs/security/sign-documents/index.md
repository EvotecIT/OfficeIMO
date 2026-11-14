---
title: Sign and Verify Word, Excel, PowerPoint and OpenDocument Files
description: Add digital signatures to DOCX, XLSX, PPTX and ODT/ODS/ODP packages with OfficeIMO.Security, validate existing signatures, and choose a certificate trust policy.
order: 60
verified: 2026-11-14
packages:
  - OfficeIMO.Security
  - OfficeIMO.Word
  - OfficeIMO.OpenDocument
related:
  - /docs/security/
  - /docs/security/cms/
  - /docs/pdf/encrypt-sign/
  - /docs/open-document/create-read/
limits: Signing works on saved files; editing a signed package afterwards invalidates its signatures, and format packages refuse such saves unless you remove the signatures explicitly.
search_terms:
  - digital signature
  - sign docx
  - sign xlsx
  - sign odt
  - verify signature
  - x509 certificate
---

Each format package owns its signature structure and passes the cryptography to the provider you supply, usually `OfficeSecurityProvider.Default` from `OfficeIMO.Security`. Word, Excel, PowerPoint and Visio packages get OPC (Open Packaging Conventions) signatures; OpenDocument packages get an XML signature over the package manifest. Signing reads the signature back and verifies it before the call returns.

The examples create a throwaway self-signed certificate in memory. In production, load your signing certificate the way your application manages keys; the provider never searches certificate stores for you.

## Sign and verify a Word document

### C#

```csharp
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;
using OfficeIMO.Word;

using (WordDocument document = WordDocument.Create("contract.docx")) {
    document.AddParagraph("Signed contract text.");
    document.Save();
}

using X509Certificate2 certificate = CreateTestCertificate("CN=Contoso Signing Test");
IOfficeSecurityProvider security = OfficeSecurityProvider.Default;

WordPackageSigningResult signed = WordDocument.SignPackage("contract.docx", security, certificate);
Console.WriteLine($"Signed {signed.SignedPartCount} parts");

using WordDocument loaded = WordDocument.Load("contract.docx");
var options = new WordSignatureValidationOptions();
options.CertificateValidation.ChainEvaluator = (cert, chain) => cert.Thumbprint == certificate.Thumbprint;   // trust only this test certificate

WordSignatureValidationReport validation = loaded.ValidateSignatures(security, options);
Console.WriteLine($"Cryptography: {validation.CryptographicStatus}, chain: {validation.CertificateChainStatus}");
Console.WriteLine(validation.IsValidUnderPolicy ? "Valid" : string.Join(Environment.NewLine, validation.Findings));

static X509Certificate2 CreateTestCertificate(string subject) {
    using RSA rsa = RSA.Create(2048);
    var request = new CertificateRequest(subject, rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
    request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));
    return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
}
```

The report keeps package structure, XML signature, cryptographic, certificate chain, revocation, timestamp and signed-part coverage outcomes separate, so you can tell "the bytes changed" from "the signer is not trusted".

## Sign Excel and PowerPoint files

Excel, PowerPoint and Visio share one OPC signing engine. Validation returns an `OfficePackageSignatureValidationReport` with `IsCryptographicallyValid` and `IsValidUnderPolicy`.

### C#

```csharp
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Security;

using (ExcelDocument workbook = ExcelDocument.Create("budget.xlsx")) {
    workbook.AddSheet("Budget").Cell(1, 1).SetValue("Approved");
    workbook.Save();
}
using (PowerPointDocument deck = PowerPointDocument.Create("deck.pptx")) {
    deck.Save();
}

using X509Certificate2 certificate = CreateTestCertificate("CN=Contoso Signing Test");
IOfficeSecurityProvider security = OfficeSecurityProvider.Default;

ExcelDocument.SignPackageSignature("budget.xlsx", security, certificate);
PowerPointDocument.SignPackage("deck.pptx", security, certificate);

// Check that the content is intact without deciding whether the signer is trusted.
var options = new OfficePackageSignatureValidationOptions { ValidateCertificateTrust = false };
OfficePackageSignatureValidationReport excel = ExcelDocument.ValidatePackageSignatures("budget.xlsx", security, options);
OfficePackageSignatureValidationReport powerPoint = PowerPointDocument.ValidateSignatures("deck.pptx", security, options);
Console.WriteLine($"budget.xlsx: {excel.IsCryptographicallyValid}, deck.pptx: {powerPoint.IsCryptographicallyValid}");

static X509Certificate2 CreateTestCertificate(string subject) {
    using RSA rsa = RSA.Create(2048);
    var request = new CertificateRequest(subject, rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
    request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));
    return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
}
```

`TrySignPackage` and `TrySignPackageSignature` return a result with `Succeeded` and `Details` instead of throwing.

## Sign and verify an OpenDocument file

`OdfDocument.SignPackage` writes `META-INF/documentsignatures.xml` and commits the file atomically. Validation recomputes the digest of every signed package entry, so a change to `content.xml` after signing is reported per entry.

### C#

```csharp
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.OpenDocument;
using OfficeIMO.Security;

using (OdtDocument document = OdtDocument.Create()) {
    document.AddParagraph("Signed policy text.");
    document.Save("policy.odt");
}

using X509Certificate2 certificate = CreateTestCertificate("CN=Contoso Signing Test");
IOfficeSecurityProvider security = OfficeSecurityProvider.Default;

var options = new OfficeXmlPackageSignatureOptions();
options.CertificateValidation.ChainEvaluator = (cert, chain) => cert.Thumbprint == certificate.Thumbprint;

OfficeXmlPackageSigningResult signed = OdfDocument.SignPackage("policy.odt", security, certificate, options);
Console.WriteLine(signed.Succeeded ? "Signed" : string.Join(Environment.NewLine, signed.Findings));

OfficeXmlPackageSignatureValidationReport validation = OdfDocument.ValidateSignatures("policy.odt", security, options);
foreach (OfficeXmlPackageSignatureResult signature in validation.Signatures) {
    Console.WriteLine($"{signature.SignatureId}: {signature.CryptographicStatus}, chain {signature.CertificateChainStatus}");
    foreach (OfficeXmlPackageEntryDigestResult entry in signature.Entries) Console.WriteLine($"  {entry.Path}: {entry.Status}");
}

static X509Certificate2 CreateTestCertificate(string subject) {
    using RSA rsa = RSA.Create(2048);
    var request = new CertificateRequest(subject, rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
    request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));
    return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
}
```

This creates the OfficeIMO package-manifest signature profile. Signatures from other ODF producers are detected and preserved; changing a signed package fails by default rather than silently invalidating them.

## Choose a trust policy

Validation separates "the signature matches the content" from "the signer is trusted". Certificate trust follows `CertificateValidationOptions`:

| Setting | Default | Use it to |
|---|---|---|
| `ValidateChain` | `true` | Build the platform certificate chain. |
| `ChainEvaluator` | none | Decide trust yourself, for example by pinning a thumbprint or a private root. |
| `RevocationMode` | `NoCheck` | Turn on online or offline revocation checking. |
| `DisableCertificateDownloads` | `true` | Allow the platform to fetch missing issuer certificates. |
| `VerificationTime` | now | Validate the chain as of the signing time. |
| `ExtraCertificates` | empty | Supply intermediates. They are chain-building candidates, not trust anchors. |

Set `ValidateCertificateTrust = false` on the package validation options when you only need to know whether the content changed after signing.
