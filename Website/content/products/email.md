---
title: "OfficeIMO.Email"
description: "Read and write email, Outlook items, mailbox stores, calendars, contacts, and address books without Outlook or third-party message parsers."
layout: product
meta.seo_title: "OfficeIMO.Email for .NET applications"
meta.head_html: '<link rel="alternate" hreflang="en" href="https://officeimo.com/products/email/" /><link rel="alternate" hreflang="x-default" href="https://officeimo.com/products/email/" />'
product_label: "Email and Outlook data"
product_color: "#0284c7"
install: "dotnet add package OfficeIMO.Email"
nuget: "OfficeIMO.Email"
docs_url: "/docs/email/"
api_url: "/api/email/"
---

## One managed API for email files and mailbox data

`OfficeIMO.Email` reads, writes, converts, searches, and verifies persisted email and Outlook data through managed .NET APIs. It does not require Outlook, COM, MailKit, MimeKit, MsgKit, or a native storage library.

```csharp
using OfficeIMO.Email;

EmailDocument message = EmailDocument.Load("message.msg");
Console.WriteLine(message.Subject);
message.Save("message.eml");
```

## Formats and workflows

| Artifact family | Read | Write or export | Useful for |
|---|---|---|---|
| EML and MIME | Yes | Yes | Standards-based messages, attachments, inline resources, and embedded messages |
| Outlook MSG and OFT | Yes | Yes | Messages, templates, appointments, contacts, tasks, journals, notes, and MAPI properties |
| TNEF and `winmail.dat` | Yes | Yes | Outlook transport payloads and typed item projections |
| mboxo and mboxrd | Yes | Yes | Portable mailbox archives and streaming message enumeration |
| PST and OST | Yes | New PST export | Folder-aware extraction, store conversion, and migration verification |
| Outlook for Mac OLM | Yes | New PST export | Archive discovery and cross-platform migration workflows |
| EMLX, Apple Mail, Maildir, and mailbox directories | Yes | Message or PST export | File-based archive ingestion and normalization |
| ICS, vCalendar, and vCard | Yes | Yes | Calendars, recurrence, reminders, contacts, and standards-based interchange |
| Outlook Offline Address Book | Yes | Search and validation results | Offline identity discovery and resolution |

## Convert a mailbox store to a new PST

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

The source is not modified. The conversion report records converted and skipped items and can compare source and destination semantics.

## Safety and verification

- immutable parsing limits for source, header, attachment, MIME, compound-file, MAPI, and TNEF boundaries
- streaming and reopenable attachment sources for large messages and store-backed content
- conversion analysis that blocks known semantic loss by default
- versioned semantic fingerprints for migration verification and deduplication
- comparison reports that expose paths and lengths without copying private message values
- optional keyed, value-free PST verification manifests

## Guides

| Guide | Use it for |
|---|---|
| [Messages and conversion](/docs/email/messages-and-conversion/) | EML, MIME, MSG, OFT, TNEF, attachments, streaming, and loss analysis |
| [Mailbox stores](/docs/email/mailbox-stores/) | PST, OST, OLM, mbox, EMLX, Apple Mail, Maildir, extraction, and PST export |
| [Calendars and contacts](/docs/email/calendars-and-contacts/) | ICS, vCard, appointments, recurrence, tasks, and Outlook personal information |
| [Safety and verification](/docs/email/safety-and-verification/) | Bounded input, diagnostics, protected messages, comparison, and migration proof |
| [Email API reference](/api/email/) | Complete types, methods, parameters, and return models |
