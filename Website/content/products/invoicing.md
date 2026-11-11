---
title: "OfficeIMO.Invoicing"
description: "Create, read, convert and validate EN 16931, Factur-X/ZUGFeRD, XRechnung, Peppol BIS and Polish FA(3) electronic invoices from one typed .NET model."
layout: product
meta.seo_title: "Electronic invoicing for .NET: Factur-X, XRechnung, Peppol and KSeF"
product_color: "#b45309"
product_label: "Electronic invoice engine"
install: "dotnet add package OfficeIMO.Invoicing"
nuget: "OfficeIMO.Invoicing"
docs_url: "/docs/invoicing/"
api_url: "/api/invoicing/"
meta.software.name: "OfficeIMO.Invoicing"
meta.software.application_category: "DeveloperApplication"
meta.software.operating_system: "Windows, Linux, macOS"
meta.software.download_url: "https://www.nuget.org/packages/OfficeIMO.Invoicing"
meta.software.price: 0
meta.software.price_currency: "USD"
---

## One invoice model, several standards

`OfficeIMO.Invoicing` writes and reads CII and UBL invoices from a single typed `Invoice`. Each output names an exact specification release, syntax and profile, and projection rejects data loss unless you explicitly allow a profile's declared reduction.

```csharp
using OfficeIMO.Invoicing;

var parsed = InvoiceParser.Read(File.ReadAllBytes("invoice.xml"));
if (parsed.HasCompleteMapping) {
    parsed.Invoice.Lines[0].Quantity = 2m;
    InvoiceCalculator.UpdateDeclaredAmounts(parsed.Invoice);
    File.WriteAllBytes("invoice-ubl.xml", parsed.Write(new InvoiceXmlOptions(
        InvoiceSpecificationRelease.En16931_1_3_16, InvoiceSyntax.Ubl, InvoiceProfile.En16931)));
}
```

## Useful for

- Issuing Factur-X/ZUGFeRD 2.5.2 (all five profiles), EN 16931, XRechnung 3.0.2 and Peppol BIS Billing 3.0.21 invoices.
- Converting between CII and UBL with every unsupported field reported instead of dropped.
- Writing and reading Polish FA(3) invoices with national amounts kept separate from EN totals.
- Producing a visible PDF and a hybrid PDF/A-3 Factur-X file from one snapshot with `OfficeIMO.Invoicing.Pdf`.
- Checking exact bytes against pinned XSD and Schematron releases with `OfficeIMO.Invoicing.Validation`.
- Submitting FA(3) invoices to KSeF online sessions and retrieving bound UPO receipts with `OfficeIMO.Invoicing.KSeF`.

The core has no external runtime dependencies and targets .NET Standard 2.0, .NET 8, .NET 10 and .NET Framework 4.7.2. Validation and KSeF target .NET 8 and .NET 10; business-rule validation needs a Java runtime and the rule archives you download. Validation results describe structure and pinned rules; they do not certify tax treatment, network delivery or KSeF acceptance.

Start with the [invoicing guide](/docs/invoicing/) or read the [package README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Invoicing/README.md).
