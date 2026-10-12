---
title: "HTML supplier offer review"
description: "Extract supplier offers with PSParseHTML and create a padded Word review, PDF and numeric Excel workbook. Download the files and runnable PowerShell recipe."
layout: docs
---

Compare four supplier offers for two items. PSParseHTML extracts the HTML table; PSWriteOffice creates an editable decision brief and a workbook that keeps prices and lead times numeric.

![Generated supplier review with padded table rows and separate decision steps](/downloads/integrations/supplier/preview.png)

[Word document](/downloads/integrations/supplier/SupplierReview.docx) · [PDF review copy](/downloads/integrations/supplier/SupplierReview.pdf) · [Excel offers](/downloads/integrations/supplier/SupplierOffers.xlsx)

## Run the example

Download the [supplier review example ZIP](/downloads/integrations/supplier-review.zip) to your working folder. It contains the script and an `Input` folder with the sample HTML. Open PowerShell in that folder and run:

```powershell
Install-Module PSParseHTML, PSWriteOffice -Scope CurrentUser
Expand-Archive -LiteralPath ./supplier-review.zip -DestinationPath ./SupplierReviewExample
./SupplierReviewExample/Recipe-HtmlSupplierReview.ps1 -OutputDirectory ./Output/SupplierReview
```

Open `Output/SupplierReview` to find the DOCX, PDF and XLSX. The example requires PSWriteOffice 3.0.8 and PSParseHTML 3.0.3 or later. If older copies are installed, run `Update-Module PSWriteOffice, PSParseHTML` and start a fresh PowerShell session.

If Windows blocks the downloaded script, read it first, then unblock that file with `Unblock-File -LiteralPath ./SupplierReviewExample/Recipe-HtmlSupplierReview.ps1`. Keep your normal execution policy.

You can also download the [standalone script](/downloads/integrations/source/Recipe-HtmlSupplierReview.ps1) and [sample HTML](/downloads/integrations/source/Input/SupplierOffers.html) separately. Keep the HTML in an `Input` folder beside the script, or pass your own file with `-HtmlPath`:

```powershell
./SupplierReviewExample/Recipe-HtmlSupplierReview.ps1 -HtmlPath ./MyOffers.html -OutputDirectory ./Output/MyReview
```

## Adapt the data

Keep table `id="offers"` and the `Item`, `Supplier`, `UnitPriceUsd` and `LeadDays` headers. Pass your own file with `-HtmlPath`. Prices are USD per unit with a decimal point; freight and tax are excluded. Quotes support at most 15 significant digits and 28 fractional places: leading zeros and fractional padding zeros do not count; all integer digits count. Comma decimals, group separators and quotes beyond those limits reject before output. Sub-cent prices remain distinct in Word and in numeric Excel cells; each quote displays its own decimal places, and the workbook fits the price column to those values. The recipe rejects missing tables, empty item/supplier names and negative price or lead-time values before writing documents.

![Native Excel range preview with numeric price and lead-time columns](/downloads/integrations/supplier/workbook.png)

Word uses semantic paragraph spacing and cell padding to separate the heading, offer table and decision steps. Excel retains editable numbers for sorting, filtering and calculations. PDF pagination is produced by OfficeIMO and can differ from Microsoft Word.

Continue with the [email review package](/docs/integrations/report-email/) to attach the PDF and XLSX without sending them.

## Complete PowerShell script

The code below is the same script included in the ZIP and standalone download. PSParseHTML reads the table; PSWriteOffice formats the review and preserves numeric values in Excel.

{{< integration-example id="supplier-review" >}}
