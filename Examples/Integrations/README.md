# Reports from HTML, events and email workflows

These recipes pass ordinary objects and files between modules. PSParseHTML extracts HTML, PSEventViewer queries Windows events, PSWriteOffice creates documents, and Mailozaurr packages or delivers attachments.

| Workflow | Recipe | Output |
| --- | --- | --- |
| Review supplier prices and lead times from an HTML table | [Supplier review](Recipe-HtmlSupplierReview.ps1) | DOCX, PDF and numeric XLSX |
| Hand off a bounded event snapshot for investigation | [Event incident brief](Recipe-PSEventViewer-OfficeReport.ps1) | DOCX, PDF and full XLSX register |
| Prepare attachments for a mail review | [Report email package](Recipe-Mailozaurr-PdfDelivery.ps1) | Local EML; a sample PDF when attachments are omitted |

Install the modules used by your chosen recipe:

```powershell
# Supplier review
Install-Module PSWriteOffice, PSParseHTML -Scope CurrentUser
# Offline event brief; add PSEventViewer only for live collection
Install-Module PSWriteOffice -Scope CurrentUser
# Email review package
Install-Module PSWriteOffice, Mailozaurr -Scope CurrentUser
```

If you already have older copies, use `Update-Module` for those modules before starting a fresh PowerShell session. The examples require PSWriteOffice 3.0.8, PSParseHTML 3.0.3, PSEventViewer 4.0.1 for live collection, and Mailozaurr 3.0.2 or later. The event module's source repository is [EventViewerX](https://github.com/EvotecIT/EventViewerX).

From the PSWriteOffice repository root:

```powershell
./Examples/Integrations/Recipe-HtmlSupplierReview.ps1 -OutputDirectory ./Output/SupplierReview

./Examples/Integrations/Recipe-PSEventViewer-OfficeReport.ps1 -OutputDirectory ./Output/EventBrief

$attachments = @(
    './Output/SupplierReview/SupplierReview.pdf'
    './Output/SupplierReview/SupplierOffers.xlsx'
)
./Examples/Integrations/Recipe-Mailozaurr-PdfDelivery.ps1 -AttachmentPath $attachments -OutputDirectory ./Output/Mail
```

The supplier and event defaults use synthetic inputs. The mail command saves `ReportReview.eml` with a unique `example.invalid` message identity and sends nothing. Open that file to review its recipients, text and attachments.

## Adapt the inputs

For supplier offers, copy [SupplierOffers.html](Input/SupplierOffers.html) and keep table `id="offers"` with the `Item`, `Supplier`, `UnitPriceUsd` and `LeadDays` headers. Prices are USD per unit, use a decimal point, and exclude tax and freight. Comma decimals and group separators reject before any documents are written. Sub-cent quotes such as `0.003` and `0.004` remain distinct in Word and Excel. Pass your file with `-HtmlPath`; the XLSX retains numeric cells, subject to Excel's normal numeric precision.

For events, use `-CollectLive` on Windows to query the selected local log for critical events, warnings and errors from the last 24 hours:

```powershell
./Examples/Integrations/Recipe-PSEventViewer-OfficeReport.ps1 -CollectLive -LogName Application -MaxEvents 100 -OutputDirectory ./Output/LiveBrief
```

Or query once yourself and pass raw or typed `Get-EVXEvent` results with `-Event`. Typed results use their `SourceEvent` snapshot. The script retains at most `MaxEvents` records, shows the top eight providers in Word, and keeps all retained records in Excel, including the source computer and log where supplied. The register displays seconds and milliseconds to distinguish events within the same minute. An empty snapshot produces an explicit empty report. A count of returned events is not the total size of the log or a root-cause conclusion. Review full messages before sharing the workbook.

For email, pass any existing PDF, DOCX or XLSX paths through `-AttachmentPath`. Omitting attachments creates a small sample service PDF. Delivery is a separate opt-in: the existing `-Send` switch requires `-SmtpServer`; `-Port` defaults to 587 and the recipe uses StartTLS. Supply your real recipients and `-Credential` where the server requires authentication. Local EML creation needs no mail credentials.
