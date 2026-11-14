---
title: Sign, Verify and Encrypt Data with CMS
description: Create detached or encapsulated CMS/PKCS#7 signatures, verify them with a trust policy, validate certificates, and encrypt data for X.509 recipients with OfficeIMO.Security.
order: 61
verified: 2026-11-14
packages:
  - OfficeIMO.Security
related:
  - /docs/security/
  - /docs/security/sign-documents/
  - /docs/pdf/encrypt-sign/
  - /docs/email/safety-and-verification/
limits: XML DSig algorithms are a closed built-in set, and on targets before .NET 8, or for envelope profiles other than RSA with AES-CBC, decryption needs an exportable RSA key.
search_terms:
  - pkcs7
  - cms signature
  - detached signature
  - encrypt for certificate
  - envelopeddata
  - certificate validation
---

`IOfficeSecurityProvider` works on plain bytes when no document format is involved: signing release artifacts, verifying a `.p7s` file, or encrypting a payload for a known recipient. Signing uses the certificate's platform `RSA` or `ECDsa` key without exporting it. Verification reports the signature, message digest, certificate trust, revocation and timestamp outcomes separately.

The examples use throwaway self-signed certificates created in memory.

## Sign data and verify the signature

A detached signature is stored apart from the content, like a `.sig` file next to a download. An encapsulated signature carries the content inside it.

### C#

```csharp
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using OfficeIMO.Security;

IOfficeSecurityProvider security = OfficeSecurityProvider.Default;
using X509Certificate2 signer = CreateTestCertificate("CN=Release Signing Test");

byte[] content = Encoding.UTF8.GetBytes("release 1.2.0 checksum list");
byte[] signature = security.SignCmsDetached(content, signer, new CmsSigningOptions { UsePadesBaseline = true });
File.WriteAllBytes("release.p7s", signature);

var options = new CmsVerificationOptions();
options.CertificateValidation.ChainEvaluator = (cert, chain) => cert.Thumbprint == signer.Thumbprint;   // trust this test certificate

CmsVerificationResult result = security.VerifyCmsDetached(signature, content, options);
foreach (CmsSignerVerificationResult s in result.Signers) {
    Console.WriteLine($"{s.Subject}: signature {s.SignatureStatus}, digest {s.DigestStatus}, chain {s.CertificateValidation.ChainStatus}");
}
Console.WriteLine(result.IsCryptographicallyValid ? "Content matches the signature" : "Content or signature changed");

byte[] encapsulated = security.SignCmsEncapsulated(content, signer);
CmsVerificationResult inner = security.VerifyCms(encapsulated, options);
Console.WriteLine(Encoding.UTF8.GetString(inner.EncapsulatedContent!));

static X509Certificate2 CreateTestCertificate(string subject) {
    using RSA rsa = RSA.Create(2048);
    var request = new CertificateRequest(subject, rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
    request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));
    return request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));
}
```

`UsePadesBaseline` produces baseline CMS with a SHA-256 signing-certificate-v2 binding and no CMS signing time, as PDF signatures expect. `CmsSigningOptions.DigestAlgorithm` defaults to SHA-256.

## Validate a certificate under your trust policy

`ValidateCertificate` returns the same chain, revocation and usage result without parsing CMS first. Additional certificates are chain-building candidates, not trusted anchors.

### C#

```csharp
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

IOfficeSecurityProvider security = OfficeSecurityProvider.Default;
using RSA rsa = RSA.Create(2048);
var request = new CertificateRequest("CN=Document Signer Test", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));
using X509Certificate2 certificate = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));

var options = new CertificateValidationOptions {
    RevocationMode = X509RevocationMode.NoCheck,
    DisableCertificateDownloads = true
};

CertificateTrustValidationResult trust = security.ValidateCertificate(
    certificate,
    options: options,
    purpose: CertificateValidationPurpose.DocumentSigning);

Console.WriteLine($"Chain: {trust.Validation.ChainStatus}, revocation: {trust.Validation.RevocationStatus}");
foreach (SecurityFinding finding in trust.Findings) Console.WriteLine($"{finding.Severity} {finding.Code}: {finding.Message}");
```

A self-signed certificate that no root trusts reports an invalid chain here, which is the expected answer. The defaults disable certificate downloads and revocation checks so validation never starts network traffic on its own; set `RevocationMode`, `VerificationTime` or a `ChainEvaluator` when your application owns a different policy. Key-usage checks stay active even without chain building.

## Encrypt data for a recipient

`EncryptCms` writes CMS EnvelopedData that only the holders of the recipient certificates' private keys can open. Content is encrypted with AES-256-CBC by default.

### C#

```csharp
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using OfficeIMO.Security;

IOfficeSecurityProvider security = OfficeSecurityProvider.Default;

using RSA rsa = RSA.Create(2048);
var request = new CertificateRequest("CN=Recipient Test", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.KeyEncipherment, critical: true));
using X509Certificate2 recipient = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));

byte[] payload = Encoding.UTF8.GetBytes("quarterly payroll export");
byte[] envelope = security.EncryptCms(payload, new[] { recipient }, new CmsEnvelopeOptions {
    ContentEncryptionAlgorithm = CmsContentEncryptionAlgorithm.Aes256Cbc
});
File.WriteAllBytes("payroll.p7m", envelope);

CmsDecryptionResult decrypted = security.DecryptCms(envelope, recipient);
Console.WriteLine(decrypted.Decrypted
    ? Encoding.UTF8.GetString(decrypted.Content!)
    : string.Join(Environment.NewLine, decrypted.Findings.Select(f => f.Message)));
```

Recipient selection is exact and yours: pass the certificate that holds the private key. On .NET 8 and later, RSA recipients with AES-CBC content decrypt through the platform key handle, so keys that forbid export still work. `CmsEnvelopeOptions` also bounds content size, encoded size and recipient count.

## Add trusted timestamps

`Rfc3161TimestampClient` requests RFC 3161 timestamps through a transport callback you supply, so your HTTP client owns timeouts and response-size limits. Plug it into signing through `SignatureTimestampProvider`.

### C#

```csharp
using System.Net.Http.Headers;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

const int maxResponseBytes = 1024 * 1024;
using var http = new HttpClient { Timeout = TimeSpan.FromSeconds(15), MaxResponseContentBufferSize = maxResponseBytes };

async Task<byte[]> SendAsync(byte[] request, CancellationToken cancellationToken) {
    using var body = new ByteArrayContent(request);
    body.Headers.ContentType = new MediaTypeHeaderValue("application/timestamp-query");
    using HttpResponseMessage response = await http.PostAsync("https://timestamp.example.com/tsa", body, cancellationToken);
    response.EnsureSuccessStatusCode();
    return await response.Content.ReadAsByteArrayAsync(cancellationToken);
}

var tsaTrust = new CertificateValidationOptions { RevocationMode = X509RevocationMode.Online };
var tsa = new Rfc3161TimestampClient(SendAsync, tsaTrust, maxResponseBytes);

var signingOptions = new CmsSigningOptions {
    UsePadesBaseline = true,
    SignatureTimestampProvider = tsa.GetTimestamp
};

using RSA rsa = RSA.Create(2048);
var request = new CertificateRequest("CN=Release Signing Test", rsa, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);
request.CertificateExtensions.Add(new X509KeyUsageExtension(X509KeyUsageFlags.DigitalSignature, critical: true));
using X509Certificate2 signer = request.CreateSelfSigned(DateTimeOffset.UtcNow.AddMinutes(-5), DateTimeOffset.UtcNow.AddDays(1));

byte[] signature = OfficeSecurityProvider.Default.SignCmsDetached(File.ReadAllBytes("release.zip"), signer, signingOptions);
File.WriteAllBytes("release.zip.p7s", signature);
```

The transport receives the DER request bytes and a cancellation token and returns the DER response. Replace the URL with your timestamp authority, and keep the response-size limit on the HTTP client so a large reply is refused before it is buffered. The client checks the nonce, message imprint, TSA signature, certificate profile and trust policy. `VerifyTimestamp` validates an existing token against the timestamped bytes.
