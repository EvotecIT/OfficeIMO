---
title: "Email safety, diagnostics, and verification"
description: "Use bounded parsing, streaming attachments, conversion diagnostics, semantic comparison, and privacy-preserving verification for email workflows."
meta.seo_title: "Safely parse and verify email migrations in .NET"
order: 40
---

Persisted email is untrusted input. OfficeIMO.Email applies limits before retaining decoded content and exposes structured diagnostics when a source is malformed, exceeds a limit, or cannot be converted without known loss.

## Set resource limits

```csharp
var options = new EmailReaderOptions(
    maxInputBytes: 64L * 1024L * 1024L,
    maxAttachmentBytes: 32L * 1024L * 1024L,
    includeAttachmentContent: false);

EmailReadResult result =
    await new EmailDocumentReader(options).ReadAsync("message.msg");

foreach (EmailDiagnostic diagnostic in result.Diagnostics) {
    Console.WriteLine(
        $"{diagnostic.Severity}: {diagnostic.Code}: {diagnostic.Message}");
}
```

Reader options cover source and header sizes, MIME part count and depth, attachment sizes, embedded-message depth, compound-file directory entries, MAPI property counts, decoded property bytes, and TNEF attributes. Set `includeAttachmentContent: false` when only metadata is required.

## Compare message semantics

Semantic comparison uses a canonical, versioned projection rather than serialized bytes:

```csharp
EmailSemanticComparisonReport comparison =
    EmailSemanticComparer.Compare(source, destination);

if (!comparison.IsMatch) {
    foreach (EmailSemanticDifference difference
        in comparison.Differences) {
        Console.WriteLine($"{difference.Kind}: {difference.Path}");
    }
}
```

The migration profile normalizes store identity and serialization details. The strict profile includes representation details, while the deduplication profile excludes synchronization and modification state.

Difference reports contain canonical paths and lengths, not message values. Supply a random caller-owned digest key before persisting fingerprints for private mail so the resulting HMAC-SHA-256 values cannot be correlated without that key.

## Protected messages

The package detects clear-signed and opaque S/MIME as well as signed or encrypted OpenPGP/MIME wrappers. An unchanged protected document can be emitted in its source format from retained bytes. Editing it or converting it to another format is blocked by default because regeneration would invalidate its cryptographic meaning.

S/MIME verification and decryption require explicit trust, revocation, and recipient-certificate choices. The original protected document is not mutated; verified or decrypted content is returned separately.

## Store conversion verification

Mailbox-to-PST conversion can compare source and destination item semantics and write an optional value-free manifest. The report identifies attempted, matched, and failed items so a migration job can fail or quarantine exceptions instead of treating a completed file write as proof of equivalence.

See [mailbox stores](/docs/email/mailbox-stores/) for a complete OST/PST conversion example and [benchmark evidence](/docs/capabilities/benchmarks/#performance-guardrails) for the committed Email workload envelopes.
