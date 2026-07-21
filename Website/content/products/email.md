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

## One model for messages and Outlook artifacts

`OfficeIMO.Email` reads and writes EML/MIME, MSG, OFT, TNEF, mbox, calendar, contact, and Outlook store artifacts through managed .NET APIs. It does not require Outlook, COM, MailKit, MimeKit, MsgKit, or a native storage library.

```csharp
using OfficeIMO.Email;

EmailDocument message = EmailDocument.Load("message.msg");
Console.WriteLine(message.Subject);
message.Save("message.eml");
```

## Workflows

| Need | Capability |
|---|---|
| Message migration | Deterministic EML, MSG, OFT, TNEF, and mbox read/write with explicit conversion-loss policy |
| Archive discovery | Mixed-artifact discovery across messages, calendars, contacts, stores, and address books |
| Large stores | File-backed and streaming reads for PST, OST, OLM, EMLX, Maildir, and mailbox directories |
| Verification | Structured diagnostics, safety limits, semantic fingerprints, and comparison reports |
| Content pipelines | HTML, RTF, attachments, inline resources, and Reader adapters owned by the corresponding OfficeIMO engines |
