---
title: "Read, Write, and Convert Email and Outlook Files"
description: "OfficeIMO.Email reads and writes EML, MSG, OFT, TNEF, PST, OST, OLM, mbox, calendar and contact files in .NET, with no Outlook installed."
meta.seo_title: "Read and convert email and Outlook files in .NET"
order: 36
packages:
  - OfficeIMO.Email
cmdlets:
  - Open-OfficeEmail
  - Save-OfficeEmail
commands:
  - officeimo convert
search_terms:
  - outlook msg
  - eml
  - pst
  - email parser
  - winmail.dat
verified: 2026-11-14
---

`OfficeIMO.Email` works with saved email and Outlook data without Outlook, COM, or third-party message parsers. Use `EmailDocument` for single messages (EML/MIME, MSG, OFT, TNEF), `EmailStoreSession` for mailbox stores, and `IcsDocument` or `VCardDocument` for calendars and contacts.

| You want to | Page |
|---|---|
| Read or convert EML, MSG, OFT, TNEF, or `winmail.dat` | [Read and Convert Email Messages](/docs/email/messages-and-conversion/) |
| Read PST, OST, OLM, mbox, EMLX, Apple Mail, or Maildir, or convert a store to PST | [Read and Convert PST, OST, and Mailbox Stores](/docs/email/mailbox-stores/) |
| Read or update ICS, vCalendar, vCard, or Outlook recurrence | [Read and Edit Calendars, Contacts, and Outlook Items](/docs/email/calendars-and-contacts/) |
| Limit untrusted input, read diagnostics, or verify a migration | [Read Untrusted Email Safely and Verify Conversions](/docs/email/safety-and-verification/) |

For text extraction across many formats rather than format-specific editing, use the [Reader](/docs/reader/).

## Convert one message

`Save` picks the output format from the extension. A conversion that would lose protected content, recurrence data, or another known semantic is blocked unless you accept the loss explicitly.

### C#

```csharp
using OfficeIMO.Email;

EmailDocument message = EmailDocument.Load("message.msg");
Console.WriteLine(message.Subject);
Console.WriteLine(message.Body.Text);
message.Save("message.eml");
```

### PowerShell

```powershell
Open-OfficeEmail -Path .\message.msg | Save-OfficeEmail -Path .\message.eml
```

### CLI

```sh
officeimo convert message.msg message.eml
```

## What the package covers

| Family | Formats and operations |
|---|---|
| Messages | EML/MIME, MSG, OFT, TNEF, attachments, embedded messages, inline resources, HTML, plain text, and RTF bodies |
| Outlook items | Messages, appointments, contacts, tasks, journals, sticky notes, meeting items, and distribution lists |
| Stores | mboxo, mboxrd, PST, OST, OLM, EMLX, Apple Mail, Maildir, and mailbox directories |
| Calendars and contacts | ICS, vCalendar, vCard 2.1/3.0/4.0, recurrence, time zones, and reminders |
| Address books | Outlook Offline Address Book discovery, v4 entry decoding, search, and offline identity resolution |
| Verification | Structured diagnostics, bounded I/O, conversion analysis, semantic fingerprints, and comparison reports |
