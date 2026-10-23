---
title: "Create reports from HTML and Windows events"
description: "Combine PSParseHTML, PSEventViewer, PSWriteOffice and Mailozaurr to create editable reports and local email review packages."
layout: docs
---

Use a collector to obtain the data, then pass ordinary objects or files to the next module. The [integration recipes](https://github.com/EvotecIT/OfficeIMO/tree/81d91150c11e2bd4f24c8b8508e768b81a0063ce/PowerShell/PSWriteOffice/Examples/Integrations) provide complete scripts with sample inputs and explicit output directories.

| Task | Start with | Deliverables |
| --- | --- | --- |
| Compare supplier prices and delivery times from HTML | [HTML supplier review](https://github.com/EvotecIT/OfficeIMO/blob/81d91150c11e2bd4f24c8b8508e768b81a0063ce/PowerShell/PSWriteOffice/Examples/Integrations/Recipe-HtmlSupplierReview.ps1) | Spaced Word report, PDF review copy and numeric Excel table |
| Assign investigation from a Windows event snapshot | [Event incident brief recipe](https://github.com/EvotecIT/OfficeIMO/blob/81d91150c11e2bd4f24c8b8508e768b81a0063ce/PowerShell/PSWriteOffice/Examples/Integrations/Recipe-PSEventViewer-OfficeReport.ps1) | Word/PDF brief and Excel register with full messages |
| Review report attachments before distribution | [Mailozaurr report package](https://github.com/EvotecIT/OfficeIMO/blob/81d91150c11e2bd4f24c8b8508e768b81a0063ce/PowerShell/PSWriteOffice/Examples/Integrations/Recipe-Mailozaurr-PdfDelivery.ps1) | Local EML with attached reports |

The first two recipes run with synthetic data by default. The mail recipe saves an EML locally; sending is an explicit `-Send` operation with your SMTP settings.

Open the worked examples to inspect the generated files before adapting the scripts: [supplier offer review](/docs/integrations/supplier-review/), [event incident brief walkthrough](/docs/integrations/event-brief/) and [email review package](/docs/integrations/report-email/).

```powershell
./Examples/Integrations/Recipe-HtmlSupplierReview.ps1 -OutputDirectory ./Output/SupplierReview
./Examples/Integrations/Recipe-PSEventViewer-OfficeReport.ps1 -OutputDirectory ./Output/EventBrief

./Examples/Integrations/Recipe-Mailozaurr-PdfDelivery.ps1 `
    -AttachmentPath ./Output/EventBrief/EventIncidentBrief.pdf,./Output/EventBrief/EventRegister.xlsx `
    -OutputDirectory ./Output/Mail
```

The [recipe guide](https://github.com/EvotecIT/OfficeIMO/blob/81d91150c11e2bd4f24c8b8508e768b81a0063ce/PowerShell/PSWriteOffice/Examples/Integrations/README.md) explains module requirements, HTML headers, bounded live event collection, supplied snapshots and delivery options. The [document layout guide](/docs/pswriteoffice/document-layout/) covers the paragraph spacing and cell padding used by these reports.

PSParseHTML owns HTML extraction. [EventViewerX](https://github.com/EvotecIT/EventViewerX) owns the PSEventViewer module and its event queries. OfficeIMO owns document generation through PSWriteOffice, while Mailozaurr owns MIME and transport. These integrations add no required modules to PSWriteOffice itself.
