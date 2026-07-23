---
title: "Email and Outlook data for .NET"
description: "Read, write, convert, search, and verify EML, MIME, MSG, OFT, TNEF, PST, OST, OLM, mbox, EMLX, calendars, contacts, and address books."
meta.seo_title: "Read and convert email and Outlook files in .NET"
order: 36
---

## Install

```shell
dotnet add package OfficeIMO.Email
```

`OfficeIMO.Email` works with persisted email and Outlook data without Outlook, COM automation, or third-party message parsers. Use the message API for EML/MIME, MSG, OFT, and TNEF; the store API for mailbox archives; and the calendar/contact APIs for standards-based personal information.

## Choose your workflow

| You need to | Start here |
|---|---|
| Read or write EML, MIME, MSG, OFT, TNEF, or `winmail.dat` | [Messages and conversion](/docs/email/messages-and-conversion/) |
| Extract messages or convert PST, OST, OLM, mbox, EMLX, Apple Mail, or Maildir | [Mailbox stores](/docs/email/mailbox-stores/) |
| Read and update ICS, vCalendar, vCard, appointments, contacts, tasks, or recurrence | [Calendars and contacts](/docs/email/calendars-and-contacts/) |
| Bound untrusted input, inspect diagnostics, or verify a migration | [Safety and verification](/docs/email/safety-and-verification/) |

## Read and convert one message

```csharp
using OfficeIMO.Email;

EmailDocument message = EmailDocument.Load("message.msg");
Console.WriteLine(message.Subject);
Console.WriteLine(message.Body.Text);
message.Save("message.eml");
```

`Save` infers EML, MSG, OFT, or TNEF from the destination extension. A conversion that would lose protected content, recurrence data, or another known semantic is blocked unless the caller explicitly accepts the reported loss.

## Supported artifact families

| Family | Formats and operations |
|---|---|
| Messages | EML/MIME, MSG, OFT, TNEF, attachments, embedded messages, inline resources, HTML, plain text, and RTF bodies |
| Outlook items | Messages, appointments, contacts, tasks, journals, sticky notes, meeting lifecycle items, and distribution lists |
| Stores | mboxo, mboxrd, PST, OST, OLM, EMLX, Apple Mail, Maildir, and mailbox directories |
| Personal information | ICS, vCalendar, vCard 2.1/3.0/4.0, recurrence, time zones, and reminders |
| Address books | Outlook Offline Address Book discovery, v4 entry decoding, search, validation, and offline identity resolution |
| Verification | Structured diagnostics, bounded I/O, conversion analysis, semantic fingerprints, comparison reports, and optional value-free manifests |

Use the [Reader extraction guide](/docs/reader/) when the goal is normalized extraction rather than format-specific editing. Browse the [Email API reference](/api/email/) for the complete type surface.
