---
title: "Unlock a password-protected PDF copy"
description: "Remove PDF Standard password security with a valid owner password, verify the result is unencrypted, and download a separate unlocked document."
meta.workflow_id: "unlock"
meta.eyebrow: "Secure a PDF"
meta.source_format: "Protected PDF and owner password"
meta.destination_format: "Unprotected PDF"
meta.package: "OfficeIMO.Pdf"
meta.package_url: "https://www.nuget.org/packages/OfficeIMO.Pdf"
meta.runtime: "Browser-local WebAssembly or .NET"
meta.primary_url: "/apps/officeimo-converter/?workspace=pdf&tool=unlock"
meta.primary_label: "Unlock a PDF in the browser"
meta.secondary_url: "/docs/pdf/security/"
meta.secondary_label: "Read the security guide"
meta.summary_title: "Unlock summary"
meta.limit: "A valid owner password is required; OfficeIMO does not bypass or recover unknown credentials."
meta.related_url: "/pdf/"
meta.related_label: "Browse all PDF workflows"
meta.howto.name: "Create an unprotected copy of an encrypted PDF"
meta.howto.description: "Authenticate with the owner password, remove Standard security, and verify the separate output."
meta.howto.steps:
  - name: "Select"
    text: "Choose one Standard-security PDF that you are authorized to unlock."
  - name: "Authenticate"
    text: "Enter the owner password required for privileged document changes."
  - name: "Verify"
    text: "Create a separate PDF and confirm that the output no longer reports encryption."
---

Unlocking decrypts an authorized PDF and writes a new document without Standard password security. The browser never overwrites the protected source.

## Unlock from .NET

```csharp
using OfficeIMO.Pdf;

const string ownerPassword = "owner-password";
var readOptions = new PdfLoadOptions { Password = ownerPassword };
PdfDocument source = PdfDocument.Load("statement.protected.pdf", readOptions);

PdfSecurityMutationResult result = source.Security.Decrypt(ownerPassword);
File.WriteAllBytes("statement.unlocked.pdf", result.Pdf);
Console.WriteLine($"Encrypted: {result.IsEncrypted}");
```

The password is required both to open the encrypted content and to authorize removal of security. Keep it out of source, command history, telemetry, and exception messages.

## Authorization boundary

OfficeIMO does not guess, recover, or bypass unknown passwords. If the caller lacks the owner credential or permission to remove protection, the correct result is a blocked workflow rather than a weakened file.
