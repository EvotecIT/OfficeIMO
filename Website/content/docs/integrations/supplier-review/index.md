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

Read the [complete script](https://github.com/EvotecIT/PSWriteOffice/blob/76b1994c7a3a2fe23d5e57b180c94aa1c8e43d58/Examples/Integrations/Recipe-HtmlSupplierReview.ps1) and [sample HTML](https://github.com/EvotecIT/PSWriteOffice/blob/76b1994c7a3a2fe23d5e57b180c94aa1c8e43d58/Examples/Integrations/Input/SupplierOffers.html). For a small standalone run, download the [script](/downloads/integrations/source/Recipe-HtmlSupplierReview.ps1) and save the [HTML input](/downloads/integrations/source/Input/SupplierOffers.html) as `Input/SupplierOffers.html` beside it.

## Adapt the data

Keep table `id="offers"` and the `Item`, `Supplier`, `UnitPriceUsd` and `LeadDays` headers. Pass your own file with `-HtmlPath`. Prices are USD per unit with a decimal point; freight and tax are excluded. Comma decimals and group separators reject before output. Sub-cent prices remain distinct in Word and in numeric Excel cells; each quote displays its own decimal places. The recipe rejects missing tables, empty item/supplier names and negative price or lead-time values before writing documents.

![Native Excel range preview with numeric price and lead-time columns](/downloads/integrations/supplier/workbook.png)

Word uses semantic paragraph spacing and cell padding to separate the heading, offer table and decision steps. Excel retains editable numbers for sorting, filtering and calculations. PDF pagination is produced by OfficeIMO and can differ from Microsoft Word.

Continue with the [email review package](../report-email/) to attach the PDF and XLSX without sending them.
