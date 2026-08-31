---
title: "Protect a PDF with AES-256 passwords"
description: "Encrypt a PDF with AES-256 Standard security, set document-open and owner passwords, and review preservation evidence for the protected copy."
meta.workflow_id: "protect"
meta.eyebrow: "Secure a PDF"
meta.source_format: "PDF and passwords"
meta.destination_format: "AES-256 protected PDF"
meta.package: "OfficeIMO.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Pdf"
meta.runtime: "Browser-local WebAssembly or .NET"
meta.primary_url: "/apps/officeimo-converter/?workspace=pdf&tool=protect"
meta.primary_label: "Protect a PDF in the browser"
meta.secondary_url: "/docs/pdf/security/"
meta.secondary_label: "Read the security guide"
meta.summary_title: "Protection summary"
meta.limit: "Password encryption controls opening and permissions; it is not a digital signature or trust assertion."
meta.related_url: "/pdf/"
meta.related_label: "Browse all PDF workflows"
meta.howto.name: "Create an AES-256 password-protected PDF"
meta.howto.description: "Set separate user and owner credentials, rewrite the document, and verify the protected output."
meta.howto.steps:
  - name: "Select"
    text: "Choose one PDF that can be rewritten under the current security policy."
  - name: "Protect"
    text: "Enter a document-open password and a separate owner password."
  - name: "Verify"
    text: "Create the AES-256 output and review encryption and preservation evidence."
---

Password protection uses the PDF Standard security handler. The user password opens the document; the owner password controls privileged changes and should be stored separately from the delivered file.

## Protect from .NET

```csharp
using OfficeIMO.Pdf;

PdfDocument source = PdfDocument.Load("statement.pdf");
var encryption = new PdfStandardEncryptionOptions("reader-password") {
    OwnerPassword = "owner-password",
    Algorithm = PdfStandardEncryptionAlgorithm.Aes256
};

PdfSecurityMutationResult result = source.Security.Encrypt(encryption);
File.WriteAllBytes("statement.protected.pdf", result.Pdf);
Console.WriteLine(result.PreservationReport.Summary);
```

The mutation result exposes whether the output is encrypted and whether the rewrite preserved the expected document structure. Applications should handle passwords through their normal secret-management policy, not source code or logs.

## Encryption is not signing

Password protection does not identify the publisher, validate document integrity through a certificate, or timestamp a revision. Use the cryptographic signing and validation APIs through `OfficeIMO.Pdf` and `OfficeIMO.Security` when the workflow requires those guarantees.
