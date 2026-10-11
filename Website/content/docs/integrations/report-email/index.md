---
title: "Prepare a report email for review"
description: "Use Mailozaurr to save report attachments in a local EML for review, with an optional PSWriteOffice sample PDF. Sending requires an explicit operation."
layout: docs
---

Prepare the message as a file so a reviewer can inspect recipients, text and attachments. The default recipe creates this small service PDF and embeds it in an EML with example.invalid recipients and a unique example.invalid message identity.

![Generated service report with space between its heading, table and next action](/downloads/integrations/mail-default/preview.png)

[Service PDF](/downloads/integrations/mail-default/Daily-Service-Report.pdf) · [Local EML package](/downloads/integrations/mail-default/ReportReview.eml) · [Complete script](https://github.com/EvotecIT/PSWriteOffice/blob/87729715cabfeda884b46a3bb7a2994c97ec8093/Examples/Integrations/Recipe-Mailozaurr-PdfDelivery.ps1)

## Prepare your own attachments

```powershell
Install-Module Mailozaurr -Scope CurrentUser
$attachments = @(
    './Output/EventBrief/EventIncidentBrief.pdf'
    './Output/EventBrief/EventRegister.xlsx'
)
./Examples/Integrations/Recipe-Mailozaurr-PdfDelivery.ps1 -AttachmentPath $attachments -OutputDirectory ./Output/Mail
```

Run from a PSWriteOffice checkout, or download the [script](/downloads/integrations/source/Recipe-Mailozaurr-PdfDelivery.ps1). Existing attachments need Mailozaurr 3.0.2 or later. Omitting `-AttachmentPath` also requires PSWriteOffice 3.0.8 or later and creates the sample service PDF.

Open `ReportReview.eml` in a mail client to review it. The default command saves locally and needs no SMTP server, OAuth session or credentials. Mailozaurr embeds the ordinary report files; the document library does not own mail transport.

## Delivery options

The recipe's `-Send` switch is an explicit delivery operation. Set real `-From` and `-To` recipients, supply `-SmtpServer` and use `-Credential` when required. Port 587 and StartTLS are the recipe defaults. Choose settings appropriate to your server; the example does not configure it or authenticate an account for you.

Create the source reports with the [supplier review](../supplier-review/) or [event brief](../event-brief/) and keep their editable workbooks alongside the PDF review copies.
