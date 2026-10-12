---
title: "Prepare a report email for review"
description: "Use Mailozaurr to save report attachments in a local EML for review, with an optional PSWriteOffice sample PDF. Sending requires an explicit operation."
layout: docs
---

Prepare the message as a file so a reviewer can inspect recipients, text and attachments. The default recipe creates this small service PDF and embeds it in an EML with example.invalid recipients and a unique example.invalid message identity.

![Generated service report with space between its heading, table and next action](/downloads/integrations/mail-default/preview.png)

[Service PDF](/downloads/integrations/mail-default/Daily-Service-Report.pdf) · [Local EML package](/downloads/integrations/mail-default/ReportReview.eml)

## Try the sample report

Download the [report email example ZIP](/downloads/integrations/report-email.zip) to your working folder. It contains one standalone script. Open PowerShell in that folder and run:

```powershell
Install-Module PSWriteOffice, Mailozaurr -Scope CurrentUser
Expand-Archive -LiteralPath ./report-email.zip -DestinationPath ./ReportEmailExample
./ReportEmailExample/Recipe-Mailozaurr-PdfDelivery.ps1 -OutputDirectory ./Output/Mail
```

Open `Output/Mail/ReportReview.eml` in a mail client to inspect the message and its attached PDF. This sample requires PSWriteOffice 3.0.8 and Mailozaurr 3.0.2 or later. If older copies are installed, run `Update-Module PSWriteOffice, Mailozaurr` and start a fresh PowerShell session. You can also save the [standalone script](/downloads/integrations/source/Recipe-Mailozaurr-PdfDelivery.ps1) anywhere and run it from that location.

If Windows blocks the downloaded script, read it first, then unblock that file with `Unblock-File -LiteralPath ./ReportEmailExample/Recipe-Mailozaurr-PdfDelivery.ps1`. Keep your normal execution policy.

## Prepare your own attachments

```powershell
Install-Module Mailozaurr -Scope CurrentUser
$attachments = @(
    './Output/EventBrief/EventIncidentBrief.pdf'
    './Output/EventBrief/EventRegister.xlsx'
)
./ReportEmailExample/Recipe-Mailozaurr-PdfDelivery.ps1 -AttachmentPath $attachments -OutputDirectory ./Output/Mail
```

The attachment paths above match the [event brief example](/docs/integrations/event-brief/). Replace them with your own report paths. Packaging existing attachments requires only Mailozaurr 3.0.2 or later; PSWriteOffice is used when the script generates the sample PDF.

Open `ReportReview.eml` in a mail client to review it. The default command saves locally and needs no SMTP server, OAuth session or credentials. Mailozaurr embeds the ordinary report files; the document library does not own mail transport.

## Delivery options

The recipe's `-Send` switch is an explicit delivery operation. Set real `-From` and `-To` recipients, supply `-SmtpServer` and use `-Credential` when required. Port 587 and StartTLS are the recipe defaults. Choose settings appropriate to your server; the example does not configure it or authenticate an account for you.

Create the source reports with the [supplier review](/docs/integrations/supplier-review/) or [event brief](/docs/integrations/event-brief/) and keep their editable workbooks alongside the PDF review copies.

## Complete PowerShell script

The code below is the same script included in the ZIP and standalone download. The default run saves the message locally. Delivery requires the explicit `-Send` switch and your server settings.

{{< integration-example id="report-email" >}}
