---
title: "Encrypt, Decrypt, and Sign PDFs"
description: "Encrypt PDFs with open and owner passwords and permissions, open or decrypt protected files, sign with a certificate, and validate signatures with OfficeIMO.Pdf, PSWriteOffice, and the officeimo CLI."
layout: docs
packages:
  - OfficeIMO.Pdf
cmdlets:
  - Protect-OfficeDocument
  - Get-OfficePdfSignature
limits: "Mathematical verification proves the signed bytes are intact; trust in the signer's certificate chain is a separate policy decision."
search_terms:
  - sign pdf
  - encrypt pdf
  - pdf password
  - digital signature
commands:
  - officeimo pdf decrypt
verified: 2026-11-14
---

Encryption controls who can open a PDF and what the viewer allows (printing, copying, editing). A digital signature proves who approved the document and that it has not changed since. For the architecture, signature formats, and long-term validation, see [PDF security and digital signatures](/docs/pdf/security/).

## Encrypt with a password and permissions

### C#

```csharp
using OfficeIMO.Pdf;

PdfSecurityMutationResult result = PdfDocument.Load("report.pdf")
    .Security.Encrypt(new PdfStandardEncryptionOptions("open-password") {
        OwnerPassword = "owner-password",
        AllowedPermissions = PdfStandardPermissions.Print | PdfStandardPermissions.Accessibility
    });

File.WriteAllBytes("report-protected.pdf", result.Pdf);
```

Read passwords from a secret store or environment variable in real code rather than from source.

### PowerShell

```powershell
Protect-OfficeDocument -Path .\report.pdf -Encrypt -Password (Read-Host -AsSecureString) -OutputPath .\report-protected.pdf
```

## Open or decrypt a protected PDF

### C#

```csharp
using OfficeIMO.Pdf;

PdfDocument opened = PdfDocument.Load("report-protected.pdf", new PdfLoadOptions { Password = "owner-password" });
Console.WriteLine(opened.ExtractText());

PdfSecurityMutationResult clear = opened.Security.Decrypt("owner-password");
File.WriteAllBytes("report-clear.pdf", clear.Pdf);
```

### PowerShell

```powershell
Unprotect-OfficeDocument -Path .\report-protected.pdf -Password (Read-Host -AsSecureString) -OutputPath .\report-clear.pdf
```

### CLI

The CLI reads the owner password from an environment variable so it never appears in shell history:

```sh
officeimo pdf decrypt report-protected.pdf --password-env PDF_OWNER_PASSWORD --output report-clear.pdf
```

## Sign with a certificate

Signing uses the external-signer API. `PdfCmsExternalSigner` creates a CMS signature with a certificate that has a private key; install `OfficeIMO.Security` for the managed cryptography provider.

### C#

```csharp
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Pdf;
using OfficeIMO.Security;

using X509Certificate2 certificate = X509CertificateLoader.LoadPkcs12FromFile("signer.pfx", Environment.GetEnvironmentVariable("PFX_PASSWORD"));
using var signer = new PdfCmsExternalSigner(OfficeSecurityProvider.Default, certificate);

PdfExternalSignatureCompletion signed = PdfDocument.Load("contract.pdf")
    .Security.SignExternal(signer, new PdfExternalSignatureOptions { FieldName = "Approval" });
File.WriteAllBytes("contract-signed.pdf", signed.Pdf);
```

## Validate signatures

### C#

```csharp
using OfficeIMO.Pdf;
using OfficeIMO.Security;

PdfSignatureValidationReport report = PdfDocument.Load("contract-signed.pdf")
    .Security.ValidateSignatures(new PdfCmsSignatureCryptographyProvider(OfficeSecurityProvider.Default));

Console.WriteLine(report.HasSignatures && report.MathematicalSignaturesVerified
    ? "Signatures verified"
    : "No signatures, or a signature did not verify");
```

### PowerShell

```powershell
Get-OfficePdfSignature -Path .\contract-signed.pdf
```
