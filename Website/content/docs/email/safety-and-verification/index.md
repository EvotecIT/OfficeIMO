---
title: "Read Untrusted Email Safely and Verify Conversions"
description: "Set input limits for untrusted email, read structured diagnostics, compare two messages semantically, and handle signed or encrypted messages."
meta.seo_title: "Safely parse and verify email migrations in .NET"
order: 40
packages:
  - OfficeIMO.Email
cmdlets:
  - Open-OfficeEmail
  - New-OfficeEmailReaderOptions
limits: "Signed or encrypted messages can be saved unchanged in their source format; editing them or converting them to another format is blocked by default."
search_terms:
  - email limits
  - untrusted email
  - compare emails
  - s/mime
  - email migration check
related:
  - /docs/email/mailbox-stores/
verified: 2026-11-14
---

Treat saved email as untrusted input. `EmailReaderOptions` sets limits before decoded content is kept, and every read returns structured diagnostics for malformed input, exceeded limits, and conversions that would lose data.

## Set input limits and read diagnostics

Reader options cover source and header sizes, MIME part count and depth, attachment sizes, embedded-message depth, compound-file entries, MAPI property counts, and TNEF attributes. Set `includeAttachmentContent: false` when only metadata is needed.

### C#

```csharp
using OfficeIMO.Email;

var options = new EmailReaderOptions(
    maxInputBytes: 64L * 1024 * 1024,
    maxAttachmentBytes: 32L * 1024 * 1024,
    includeAttachmentContent: false);

using EmailReadResult result = await new EmailDocumentReader(options).ReadDocumentAsync("message.msg");

foreach (EmailDiagnostic diagnostic in result.Diagnostics) {
    Console.WriteLine($"{diagnostic.Severity}: {diagnostic.Code}: {diagnostic.Message}");
}

if (!result.HasErrors) {
    Console.WriteLine(result.Document.Subject);
}
```

### PowerShell

```powershell
$options = New-OfficeEmailReaderOptions -NoAttachmentContent -MaxInputBytes 64MB
$result = Open-OfficeEmail -Path .\message.msg -Options $options -AsResult
$result.Diagnostics | Format-Table Severity, Code, Message
```

## Compare two messages

Semantic comparison works on a canonical projection of the message, not on serialized bytes, so an EML and an MSG of the same message can match. Difference reports contain paths and lengths, not message values.

### C#

```csharp
using OfficeIMO.Email;

EmailDocument source = EmailDocument.Load("original.msg");
EmailDocument destination = EmailDocument.Load("converted.eml");

EmailSemanticComparisonReport comparison = EmailSemanticComparer.Compare(source, destination);

if (!comparison.IsMatch) {
    foreach (EmailSemanticDifference difference in comparison.Differences) {
        Console.WriteLine($"{difference.Kind}: {difference.Path}");
    }
}
```

The default `Migration` profile ignores store identity and serialization details. `Strict` includes representation details, and `Deduplication` ignores synchronization and modification state. Pass a random digest key of at least 16 bytes in `EmailSemanticComparisonOptions` before saving fingerprints of private mail, so the HMAC-SHA-256 values cannot be correlated without that key.

## Signed and encrypted messages

The reader detects clear-signed and opaque S/MIME and signed or encrypted OpenPGP/MIME. An unchanged protected message can be saved in its source format from its original bytes. Editing it, or converting it to another format, is blocked by default because regenerating it would break the signature.

S/MIME signing, encryption, verification, and decryption need explicit certificate, trust, revocation, and recipient choices through a security provider you supply. Verification and decryption do not change the original document; processed content is returned separately.

## Verify a store migration

Mailbox-to-PST conversion compares source and destination items and can write a manifest without message values. The report counts attempted, matched, and failed items, so a migration job can stop or quarantine exceptions instead of trusting a completed write. [Read and Convert PST, OST, and Mailbox Stores](/docs/email/mailbox-stores/) shows the code.
