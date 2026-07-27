---
title: "Read PST and OST Files in .NET Without Outlook"
description: "Read PST, OST, OLM, mbox, EMLX, Apple Mail, and Maildir in .NET without Outlook, then extract messages or write and verify a new PST."
meta.eyebrow: "Outlook data and mailbox migration"
meta.outcome: "Inspect and migrate mailbox stores inside your application"
meta.primary_label: "Read the mailbox store guide"
meta.primary_url: "/docs/email/mailbox-stores/"
---

OfficeIMO.Email opens Outlook PST and OST files plus OLM, mbox, EMLX, Apple Mail, Maildir, and mailbox directories without Outlook, COM, or a hosted conversion service. The same store API returns ordinary `EmailDocument` objects while retaining folder paths, associated items, diagnostics, and bounded attachment access.

## Read a PST or OST

```csharp
using OfficeIMO.Email.Store;

using EmailStoreSession session = EmailStoreSession.Open("archive.ost");

foreach (EmailStoreItemReference reference in session.EnumerateItems(
    new EmailStoreEnumerationOptions(maxItems: 10_000))) {
    EmailStoreItem item = session.ReadItem(reference);
    Console.WriteLine($"{reference.FolderPath}: {item.Document.Subject}");
}
```

Set item, attachment, total attachment, and input limits for untrusted or unusually large stores. Subjects and other message values are data, not safe output paths.

## Convert OST, OLM, mbox, or Maildir to PST

Conversion writes a new PST. It does not repair or rewrite the source store in place.

```csharp
using EmailStoreSession source = EmailStoreSession.Open("mailbox.ost");

EmailStorePstConversionReport report = source.ExportToPst(
    "mailbox-export.pst",
    new EmailStorePstConversionOptions(
        continueOnItemError: true,
        includeAssociatedItems: true,
        includeOrphanedItems: true));

Console.WriteLine($"Converted: {report.ConvertedItems}");
Console.WriteLine($"Skipped: {report.SkippedItems}");
Console.WriteLine($"Verified: {report.Verification?.IsSuccessful}");
```

The report separates converted and skipped items. Optional keyed semantic fingerprints can compare the source and destination without placing subjects, addresses, body text, or attachment names in the verification manifest.

## Choose the right output

| Goal | Route |
| --- | --- |
| Search, archive, or inspect messages | Enumerate the store and keep `EmailDocument` plus source metadata |
| Export portable individual messages | Save each document as EML with deterministic safe filenames |
| Consolidate a supported store | Export to a new PST and retain the conversion report |
| Normalize mailbox content for search or RAG | Use `OfficeIMO.Reader.Email` over the same owning email engine |
| Modify a PST | Use verified mutation APIs and reopen the result before acceptance |

Mailbox semantics vary across stores. Associated items, recurrence, RTF bodies, TNEF, embedded messages, address books, and damaged records need representative fixtures and explicit acceptance rules.

Continue with [mailbox stores](/docs/email/mailbox-stores/), [email safety and verification](/docs/email/safety-and-verification/), or the [Email API reference](/api/email/).
