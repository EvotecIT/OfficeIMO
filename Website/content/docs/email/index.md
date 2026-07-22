---
title: "Email and Outlook data"
description: "Read and write EML, MSG, OFT, TNEF, mailbox stores, calendars, contacts, and address books."
order: 36
---

## Install

```shell
dotnet add package OfficeIMO.Email
```

## Read and convert a message

```csharp
using OfficeIMO.Email;

EmailDocument message = EmailDocument.Load("message.msg");
Console.WriteLine(message.Subject);
Console.WriteLine(message.Body.Text);
message.Save("message.eml");
```

`Save` infers EML, MSG, OFT, or TNEF from the destination extension. Convenience methods fail when an error diagnostic is produced instead of silently writing a partial result.

## Supported workflow families

- Messages: EML/MIME, MSG, OFT, TNEF, attachments, embedded messages, and inline resources.
- Personal information: appointments, contacts, tasks, journals, sticky notes, iCalendar, and vCard.
- Stores: mbox, PST, OST, OLM, EMLX, Apple Mail, Maildir, and mailbox directories.
- Address books: Outlook Offline Address Book discovery, search, validation, and identity resolution.
- Pipelines: bounded I/O, cancellation, structured diagnostics, semantic fingerprints, and explicit conversion-loss policy.

Use the [Reader extraction guide](/docs/reader/) when the goal is normalized extraction rather than format-specific editing. Browse the [Email API reference](/api/email/) for the complete type surface.
