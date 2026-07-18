# OfficeIMO.Pdf.Cryptography.Pkcs

This optional package adds first-party managed CMS signing, RFC 3161 validation,
and `X509Chain` validation to the dependency-free signature parser in
`OfficeIMO.Pdf`. Install it only when an application needs certificate signing
or cryptographic signature validation.

The managed signer plugs into the existing external-signature workflow and
currently supports RSA certificates with SHA-256 signed attributes:

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Cryptography;

using var signer = new PdfPkcsExternalSigner(certificate);
PdfExternalSignatureCompletion completion = PdfDocument
    .Open(sourcePdf)
    .SignExternal(
        signer,
        new PdfExternalSignatureOptions {
            Profile = PdfSignatureProfile.Approval,
            FieldName = "Approval"
        });

completion.ToDocument().Save("signed.pdf");
```

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Cryptography;

var provider = new PdfPkcsSignatureCryptographyProvider();
var report = PdfDocument.Open("signed.pdf").ValidateSignatures(provider);

foreach (var signature in report.Signatures) {
    var crypto = signature.CryptographicResult;
    Console.WriteLine($"Math: {crypto?.MathematicalSignatureStatus}");
    Console.WriteLine($"Digest: {crypto?.MessageDigestStatus}");
    Console.WriteLine($"Chain: {crypto?.CertificateChainStatus}");
    Console.WriteLine($"Revocation: {crypto?.RevocationStatus}");
    Console.WriteLine($"Timestamp: {crypto?.TimestampStatus}");
}
```

Certificate revocation defaults to `NoCheck`; the provider does not make
network calls unless caller policy enables online revocation:

```csharp
using System.Security.Cryptography.X509Certificates;

var provider = new PdfPkcsSignatureCryptographyProvider(
    new PdfPkcsSignatureValidationOptions {
        RevocationMode = X509RevocationMode.Online,
        UrlRetrievalTimeout = TimeSpan.FromSeconds(10)
    });
```

The PDF core owns byte-range validation, exact signed-byte extraction,
signature metadata, document permissions, and revision analysis. This package
owns CMS signature math, signed attributes, certificate chains, timestamps,
and revocation policy. CMS and RFC 3161 containers are handled by bounded
first-party DER code; key operations and certificate chains use target-framework
cryptography APIs.

## Dependency footprint

- **External runtime packages:** none.
- **OfficeIMO:** `OfficeIMO.Pdf` owns signature discovery, signed-byte extraction, permissions, and revision analysis.

See the [complete OfficeIMO package map](../README.md) for related formats and conversion paths.
