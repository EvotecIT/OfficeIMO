---
title: "Read and convert mailbox stores"
description: "Read PST, OST, OLM, mbox, EMLX, Apple Mail, Maildir, and mailbox directories in .NET, export messages, and convert supported stores to PST."
meta.seo_title: "Read PST, OST, OLM, and export stores to PST"
order: 38
---

The Store API opens PST, OST, Outlook for Mac OLM, mbox, EMLX, Apple Mail, Maildir, and mailbox-directory sources without Outlook or COM. Store items are returned as ordinary `EmailDocument` objects while folder paths, store metadata, associated items, diagnostics, and bounded attachment access remain available.

## Enumerate and read a PST or OST

```csharp
using OfficeIMO.Email;
using OfficeIMO.Email.Store;

using EmailStoreSession session = EmailStoreSession.Open("archive.pst");

foreach (EmailStoreItemReference reference in session.EnumerateItems(
    new EmailStoreEnumerationOptions(maxItems: 1000))) {
    EmailStoreItem item = session.ReadItem(reference);
    Console.WriteLine($"{reference.FolderPath}: {item.Document.Subject}");
}
```

Use reader options to set item, attachment, and total attachment limits before opening untrusted or very large stores.

## Convert OST or another supported store to PST

`ExportToPst` and `EmailStoreConverter.ConvertToPst` write a new PST; they do not modify the source store.

```csharp
using OfficeIMO.Email.Store;

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

This is a source-to-new-PST conversion, not an in-place OST rewrite. The report distinguishes converted and skipped items and includes verification results when verification is enabled.

For a path-based conversion with explicit input limits:

```csharp
var readerOptions = new EmailStoreReaderOptions(
    maxItemCount: 100_000,
    maxAttachmentBytes: 64L * 1024L * 1024L,
    maxTotalAttachmentBytes: 2L * 1024L * 1024L * 1024L,
    retainAttachmentContent: false);

EmailStorePstConversionReport report = EmailStoreConverter.ConvertToPst(
    "archive.olm",
    "archive.pst",
    readerOptions,
    new EmailStorePstConversionOptions(continueOnItemError: true));
```

## Extract messages

Read each store item and save it in the message format your workflow requires:

```csharp
Directory.CreateDirectory("messages");

using EmailStoreSession session = EmailStoreSession.Open("archive.pst");
int index = 0;

foreach (EmailStoreItemReference reference in session.EnumerateItems()) {
    EmailDocument message = session.ReadItem(reference).Document;
    message.Save(Path.Combine("messages", $"{index++:D6}.eml"));
}
```

Choose file names and directory layout from stable identifiers or sanitized metadata in production. Do not use untrusted subjects directly as paths.

## Verify a migration without storing private values

PST conversion can write a verification manifest containing keyed semantic fingerprints, match status, paths, and lengths without message subjects, addresses, attachment names, or body values:

```csharp
using EmailStoreSession source = EmailStoreSession.Open("source.ost");

EmailStorePstConversionReport report = source.ExportToPst(
    "destination.pst",
    new EmailStorePstConversionOptions(
        verificationManifestPath: "destination.verification.tsv"));

if (report.Verification is { IsSuccessful: false } verification) {
    foreach (EmailStorePstVerificationIssue issue in verification.Issues) {
        Console.WriteLine(issue.SourceItemId);
    }
}
```

The caller should supply a private digest key when fingerprints will be persisted for private mail. See [safety and verification](/docs/email/safety-and-verification/) for comparison profiles and limits.
