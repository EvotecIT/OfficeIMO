# OfficeIMO.Pdf.Cryptography.Pkcs

This optional package adds `SignedCms`, RFC 3161, and `X509Chain` validation to
the dependency-free signature parser in `OfficeIMO.Pdf`. Install it only when
an application needs cryptographic signature validation.

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
and revocation policy. RFC 3161 token validation uses the typed BCL API on .NET
8 and later; older targets report that timestamp dimension as indeterminate.
