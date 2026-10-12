---
title: "HTML supplier offer review"
description: "Extract supplier offers with PSParseHTML and create a padded Word review, PDF and numeric Excel workbook. Download the files and runnable PowerShell recipe."
layout: docs
---

Compare four supplier offers for two items. PSParseHTML extracts the HTML table; PSWriteOffice creates an editable decision brief and a workbook that keeps prices and lead times numeric.

![Generated supplier review with padded table rows and separate decision steps](/downloads/integrations/supplier/preview.png)

[Word document](/downloads/integrations/supplier/SupplierReview.docx) · [PDF review copy](/downloads/integrations/supplier/SupplierReview.pdf) · [Excel offers](/downloads/integrations/supplier/SupplierOffers.xlsx)

## Run the example

Install PSParseHTML and PSWriteOffice, then run from a PSWriteOffice checkout:

```powershell
Install-Module PSParseHTML, PSWriteOffice -Scope CurrentUser
./Examples/Integrations/Recipe-HtmlSupplierReview.ps1 -OutputDirectory ./Output/SupplierReview
```

Read the [complete script](https://github.com/EvotecIT/PSWriteOffice/blob/2ca77f6b772cd954f553124fb801bed087213986/Examples/Integrations/Recipe-HtmlSupplierReview.ps1) and [sample HTML](https://github.com/EvotecIT/PSWriteOffice/blob/2ca77f6b772cd954f553124fb801bed087213986/Examples/Integrations/Input/SupplierOffers.html). For a small standalone run, download the [script](/downloads/integrations/source/Recipe-HtmlSupplierReview.ps1) and save the [HTML input](/downloads/integrations/source/Input/SupplierOffers.html) as `Input/SupplierOffers.html` beside it.

## Adapt the data

Keep table `id="offers"` and the `Item`, `Supplier`, `UnitPriceUsd` and `LeadDays` headers. Pass your own file with `-HtmlPath`. Prices are USD per unit with a decimal point; freight and tax are excluded. Quotes support at most 15 significant digits and 28 fractional places: leading zeros and fractional padding zeros do not count; all integer digits count. Comma decimals, group separators and quotes beyond those limits reject before output. Sub-cent prices remain distinct in Word and in numeric Excel cells; each quote displays its own decimal places, and the workbook fits the price column to those values. The recipe rejects missing tables, empty item/supplier names and negative price or lead-time values before writing documents.

![Native Excel range preview with numeric price and lead-time columns](/downloads/integrations/supplier/workbook.png)

Word uses semantic paragraph spacing and cell padding to separate the heading, offer table and decision steps. Excel retains editable numbers for sorting, filtering and calculations. PDF pagination is produced by OfficeIMO and can differ from Microsoft Word.

Continue with the [email review package](../report-email/) to attach the PDF and XLSX without sending them.
