---
title: "Read and Convert PST, OST, and Mailbox Stores"
description: "Open PST, OST, OLM, mbox, EMLX, Apple Mail, and Maildir stores, export their messages, and convert a store to a new PST with verification."
meta.seo_title: "Read PST, OST, OLM, and export stores to PST in .NET"
order: 38
packages:
  - OfficeIMO.Email
cmdlets:
  - Open-OfficeEmailMailbox
  - Save-OfficeEmailMailbox
  - New-OfficeEmailMailboxReaderOptions
commands:
  - officeimo agent search-email
limits: "Store conversion always writes a new PST; it never rewrites the source PST or OST in place."
search_terms:
  - read pst
  - ost to pst
  - olm to pst
  - mbox
  - extract emails from pst
verified: 2026-11-14
---

`EmailStoreSession` (namespace `OfficeIMO.Email.Store`) opens PST, OST, Outlook for Mac OLM, mbox, EMLX, Apple Mail, Maildir, and mailbox directories without Outlook. Items come back as ordinary `EmailDocument` objects, with folder, store metadata, and diagnostics alongside. Set `EmailStoreReaderOptions` limits before opening large or untrusted stores.

## List folders and messages in a PST or OST

### C#

```csharp
using OfficeIMO.Email.Store;

using EmailStoreSession session = EmailStoreSession.Open("archive.pst");

Dictionary<string, string> folderNames = session.Folders.ToDictionary(folder => folder.Id, folder => folder.Name);
foreach (EmailStoreFolderInfo folder in session.Folders) {
    Console.WriteLine($"{folder.Name}: {folder.ItemCount} items");
}

foreach (EmailStoreItemReference reference in session.EnumerateItems(
    new EmailStoreEnumerationOptions(maxItems: 1000))) {
    EmailStoreItem item = session.ReadItem(reference);
    Console.WriteLine($"{folderNames[item.FolderId]}: {item.Document.Subject}");
}
```

### CLI

Search a mailbox and print bounded results:

```sh
officeimo agent search-email archive.pst --query invoice --take 10
```

## Export every message as EML

### C#

```csharp
using OfficeIMO.Email;
using OfficeIMO.Email.Store;

Directory.CreateDirectory("messages");

using EmailStoreSession session = EmailStoreSession.Open("archive.pst");
int index = 0;

foreach (EmailStoreItemReference reference in session.EnumerateItems()) {
    EmailDocument message = session.ReadItem(reference).Document;
    message.Save(Path.Combine("messages", $"{index++:D6}.eml"));
}
```

Build file names from stable identifiers, not from subjects: subjects are untrusted input.

## Read an mbox file

`EmailMailbox` loads a whole mbox file into memory and can save it again. For very large files, open it with `EmailStoreSession` as shown above.

### C#

```csharp
using OfficeIMO.Email;

EmailMailbox mailbox = EmailMailbox.Load("archive.mbox");

foreach (EmailMailboxEntry entry in mailbox.Messages) {
    Console.WriteLine($"{entry.Document.Date:u} {entry.Document.Subject}");
}
```

### PowerShell

```powershell
$options = New-OfficeEmailMailboxReaderOptions -MaxMessageCount 5000
$mailbox = Open-OfficeEmailMailbox -Path .\archive.mbox -Options $options
$mailbox.Messages | ForEach-Object { $_.Document.Subject }
$mailbox | Save-OfficeEmailMailbox -Path .\copy.mbox
```

## Convert an OST, OLM, or mbox store to PST

`ExportToPst` and `EmailStoreConverter.ConvertToPst` write a new PST and, by default, verify it against the source after writing.

### C#

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

For a path-to-path conversion with explicit input limits:

```csharp
using OfficeIMO.Email.Store;

var readerOptions = new EmailStoreReaderOptions(
    maxItemCount: 100_000,
    maxAttachmentBytes: 64L * 1024 * 1024,
    maxTotalAttachmentBytes: 2L * 1024 * 1024 * 1024);

EmailStorePstConversionReport report = EmailStoreConverter.ConvertToPst(
    "archive.olm",
    "archive.pst",
    readerOptions,
    new EmailStorePstConversionOptions(continueOnItemError: true));

Console.WriteLine($"{report.ConvertedItems} converted, {report.SkippedItems} skipped");
```

## Verify a migration without storing message content

Conversion can write a verification manifest of keyed fingerprints, match status, paths, and lengths. It holds no subjects, addresses, attachment names, or body text.

### C#

```csharp
using OfficeIMO.Email;
using OfficeIMO.Email.Store;

string encodedKey = Environment.GetEnvironmentVariable("OFFICEIMO_EMAIL_DIGEST_KEY")
    ?? throw new InvalidOperationException("Set OFFICEIMO_EMAIL_DIGEST_KEY to a Base64-encoded key of at least 16 bytes.");

var verificationOptions = new EmailSemanticComparisonOptions(digestKey: Convert.FromBase64String(encodedKey));

using EmailStoreSession source = EmailStoreSession.Open("source.ost");

EmailStorePstConversionReport report = source.ExportToPst(
    "destination.pst",
    new EmailStorePstConversionOptions(
        verificationOptions: verificationOptions,
        verificationManifestPath: "destination.verification.tsv"));

if (report.Verification is { IsSuccessful: false } verification) {
    foreach (EmailStorePstVerificationIssue issue in verification.Issues) {
        Console.WriteLine($"{issue.SourceItemId}: {issue.Code}");
    }
}
```

Keep the digest key as a secret outside the application. The same key is needed to compare saved fingerprints later.
