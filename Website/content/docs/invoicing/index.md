---
title: "Electronic Invoices"
description: "Create, read and validate EN 16931, Factur-X, XRechnung, Peppol BIS and Polish FA(3) invoices in .NET and submit FA(3) to KSeF."
order: 56
meta.seo_title: "Create and validate e-invoices in C#"
---

`OfficeIMO.Invoicing` keeps one typed `Invoice` model for authoring, parsing, editing and converting electronic invoices. Optional packages add PDF presentation, pinned standards validation and Polish KSeF submission. The core has no external runtime dependencies and does not need a PDF engine or Java.

| Package | Purpose | Targets |
|---|---|---|
| `OfficeIMO.Invoicing` | Typed model, CII/UBL and FA(3) writing, parsing, conversion, model checks | .NET Standard 2.0, .NET 8, .NET 10, .NET Framework 4.7.2 |
| `OfficeIMO.Invoicing.Pdf` | Visible PDF and hybrid Factur-X/ZUGFeRD PDF from one invoice snapshot | .NET Standard 2.0, .NET 8, .NET 10, .NET Framework 4.7.2 |
| `OfficeIMO.Invoicing.Validation` | XSD and Schematron validation against pinned releases; offline FA(3) XSD validation | .NET 8, .NET 10 |
| `OfficeIMO.Invoicing.KSeF` | KSeF v2 token authentication, encrypted online FA(3) submission, status and UPO retrieval | .NET 8, .NET 10 |

## Supported contracts

| Release | Syntax | Profiles |
|---|---|---|
| Factur-X 1.09.2 / ZUGFeRD 2.5.2 | CII D22B | MINIMUM, BASIC WL, BASIC, EN 16931, EXTENDED |
| EN 16931 1.3.16 | CII D16B or UBL 2.1 | EN 16931 |
| XRechnung 3.0.2 (2026-08-31 rules) | CII D16B or UBL 2.1 | XRechnung |
| Peppol BIS Billing 3.0.21 | UBL 2.1 | Peppol BIS |
| Polish FA(3), schema 1-0E | National XML | VAT, KOR, ZAL, KOR_ZAL, ROZ, KOR_ROZ, UPR |

Document types 380, 381, 326, 384, 386 and 389 are mapped; Peppol BIS rejects self-billed invoices (389). `InvoiceXmlOptions.GetSupportedTargets()` lists every valid release, syntax and profile combination.

## Create CII and UBL

```shell
dotnet add package OfficeIMO.Invoicing
```

```csharp
using OfficeIMO.Invoicing;

var invoice = new Invoice {
    Number = "INV-001",
    IssueDate = new DateTime(2026, 9, 10),
    DueDate = new DateTime(2026, 10, 10),
    Currency = "EUR",
    Seller = new InvoiceParty { Name = "Example Seller", Address = new InvoiceAddress { CountryCode = "DE" } },
    Buyer = new InvoiceParty { Name = "Example Buyer", Address = new InvoiceAddress { CountryCode = "DE" } }
};
invoice.Seller.TaxRegistrations.Add(new InvoiceTaxRegistration("DE123456789", InvoiceTaxRegistration.VatScheme));
invoice.Payments.Add(new InvoicePayment {
    MeansCode = "58",
    Reference = invoice.Number,
    Account = new InvoiceBankAccount { Identifier = "DE79000000001234567890" }
});
invoice.Lines.Add(new InvoiceLine {
    Id = "1", Name = "Consulting", Quantity = 1m, UnitPrice = 100m,
    Tax = new InvoiceTaxCategory { Code = "S", Rate = 19m }
});

byte[] cii = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(
    InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2, InvoiceSyntax.Cii, InvoiceProfile.En16931));
byte[] ubl = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(
    InvoiceSpecificationRelease.En16931_1_3_16, InvoiceSyntax.Ubl, InvoiceProfile.En16931));
```

Writing checks model arithmetic and the target mapping, and identical input produces identical bytes. Lower Factur-X profiles carry less data: the default `RejectDataLoss` policy reports every populated field the profile would omit. Successful serialization is not proof of standards compliance; validate the bytes.

## Read, edit and convert

```csharp
var parsed = InvoiceParser.Read(File.ReadAllBytes("invoice.xml"));
foreach (var item in parsed.UnmappedData)
    Console.WriteLine($"{item.Location}: {item.Message}");

if (parsed.HasCompleteMapping) {
    parsed.Invoice.Lines[0].Quantity = 2m;
    InvoiceCalculator.UpdateDeclaredAmounts(parsed.Invoice);
    File.WriteAllBytes("edited.xml", parsed.Write(new InvoiceXmlOptions(
        InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2, InvoiceSyntax.Cii, InvoiceProfile.En16931)));
}

var conversion = InvoiceConverter.Convert(File.ReadAllBytes("invoice.xml"), new InvoiceXmlOptions(
    InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31, InvoiceSyntax.Ubl, InvoiceProfile.XRechnung));
if (conversion.Succeeded)
    File.WriteAllBytes("converted.xml", conversion.Xml!);
```

Unknown elements and unsupported variations are reported; rewriting blocks unless you pass `allowUnmappedDataLoss`. `GetOriginalBytes()` always returns the source unchanged. Conversion produces no bytes when observed data cannot be represented in the target, and every diagnostic names the source path. For header changes on a source with extensions, `InvoiceSourceEditor.Apply` replaces the number, dates, buyer reference or payment reference in place and keeps everything else.

## Polish FA(3)

`Fa3InvoiceWriter` writes FA(3) from the same model with explicit national options; `Fa3InvoiceReader` reads it back with national kinds, P_13/P_14 buckets and P_15 retained.

```csharp
// polish is an Invoice in PLN with Polish addresses and NIP fiscal registrations.
var national = new Fa3InvoiceWriteOptions(Fa3InvoiceKind.TaxInvoice,
    new DateTimeOffset(2026, 9, 30, 12, 0, 0, TimeSpan.Zero), new Fa3InvoiceAnnotations());
foreach (var finding in Fa3InvoiceWriter.Inspect(polish, national))
    Console.WriteLine($"{finding.Code}: {finding.Message}");
byte[] fa3Xml = Fa3InvoiceWriter.Write(polish, national);
```

Foreign-currency and correction kinds need explicit national amounts (`Fa3FiscalAmounts`); exchange rates are never guessed. Additional parties, JST/GV roles, new means of transport, previous-state correction rows, settlement charges and attachments are reported rather than written.

## PDF and hybrid Factur-X

```csharp
using OfficeIMO.Invoicing.Pdf;
using OfficeIMO.Pdf;

var contract = new InvoiceXmlOptions(
    InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2, InvoiceSyntax.Cii, InvoiceProfile.En16931);
var snapshot = PdfInvoiceDocument.Create(invoice, contract, InvoicePdfLayoutOptions.ForCultures("pl-PL", "en-GB"));
byte[] font = File.ReadAllBytes("invoice-font.ttf");
var options = new PdfCreateOptions()
    .EmbedStandardFont(PdfStandardFont.Helvetica, font, "Invoice font")
    .EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "Invoice font");
File.WriteAllBytes("invoice.pdf", snapshot.ToPdfBytes(options));
```

The PDF and embedded `factur-x.xml` come from one captured snapshot, so later model edits cannot make them disagree. Labels ship in ten languages; themes, logos, column selection and printed approval blocks are optional. Hybrid output supports PDF/A-3a/b/u, optionally with PDF/UA-1. Approval blocks are printed text, not digital signatures. Peppol BIS needs UBL and cannot be embedded through this CII adapter.

## Validate against pinned releases

```csharp
using OfficeIMO.Invoicing.Validation;

var bundle = InvoiceRuleBundle.Load("xrechnung-3.0.2-validator-configuration-2026-08-31.zip",
    facturXArchivePath: "validator-configuration-zugferd.zip");
var validator = new InvoiceValidator(bundle, new SaxonInvoiceRulesRunner("tools/saxon/saxon-he-12.10.jar"));
var report = await validator.ValidateAsync(File.ReadAllBytes("invoice.xml"),
    InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31);
Console.WriteLine($"{report.IsValid}: {report.Sha256}");
```

`IsValid` requires both XSD and business rules to pass. Rule archives and SaxonJ-HE are downloaded by you, verified by SHA-256 and never fetched by the library; business rules need a Java runtime. `Fa3SchemaValidator` validates FA(3) against the official 1-0E XSD without Java.

## Submit FA(3) to KSeF

`KsefClient` authenticates with a KSeF token, opens an online session, encrypts the invoice with AES-256 and runs the FA(3) XSD gate before submission. It defaults to the `Test` environment; `Demo` and `Production` must be selected explicitly. Mutations are never retried automatically: an uncertain outcome raises `KsefMutationAmbiguousException`, and `ResumeSubmission` reattaches to a known reference. Receipts (UPO) are bound to the session, KSeF number and invoice hash and checked against the pinned v4-3 schema. See the [KSeF guide](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Invoicing.KSeF/README.md) for the full polling flow.

## Limits

- Validation and XSD success establish structure and pinned rule results, not tax treatment, Peppol network delivery or KSeF acceptance.
- XRechnung is the only national CIUS with an authoring and validation contract. EXTENDED-CTC-FR is recognized for inspection only.
- KSeF supports token authentication and online submission. XAdES certificate authentication, batch and offline submission, permission administration and UPO signature verification are not included.
- Inputs and outputs are bounded (16 MiB XML, 50,000 collection items, 999 detailed diagnostics).

See the [invoicing README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Invoicing/README.md), [PDF adapter](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Invoicing.Pdf/README.md), [validation guide](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Invoicing.Validation/README.md) and the [PDF documentation](/docs/pdf/).
