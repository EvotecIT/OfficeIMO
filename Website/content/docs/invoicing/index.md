---
title: "Create, Read, and Validate Electronic Invoices"
description: "Create, read, convert, render, and validate EN 16931, Factur-X/ZUGFeRD, XRechnung, Peppol BIS, and Polish FA(3) invoices with OfficeIMO.Invoicing."
order: 56
meta.seo_title: "Create and validate e-invoices in C#"
packages:
  - OfficeIMO.Invoicing
  - OfficeIMO.Invoicing.Pdf
  - OfficeIMO.Invoicing.Validation
commands:
  - officeimo invoice validate
  - officeimo invoice inspect
  - officeimo invoice convert
  - officeimo invoice edit
  - officeimo invoice render
  - officeimo invoice hybrid
limits: "Passing XSD and Schematron validation proves structure and pinned rule results, not tax treatment, Peppol network delivery, or KSeF acceptance."
search_terms:
  - e-invoice
  - factur-x
  - zugferd
  - xrechnung
  - peppol
  - ksef fa(3)
related:
  - /docs/cli/invoices/
  - /docs/pdf/
verified: 2026-11-14
---

`OfficeIMO.Invoicing` has one typed `Invoice` model for writing, reading, editing, and converting electronic invoices. `OfficeIMO.Invoicing.Pdf` renders a visible PDF or a hybrid Factur-X/ZUGFeRD PDF, and `OfficeIMO.Invoicing.Validation` runs XSD and Schematron checks against pinned rule releases. `OfficeIMO.Invoicing.KSeF` submits Polish FA(3) invoices to KSeF. The core package needs no PDF engine and no Java.

| Release | Syntax | Profiles |
|---|---|---|
| Factur-X 1.09.2 / ZUGFeRD 2.5.2 | CII D22B | MINIMUM, BASIC WL, BASIC, EN 16931, EXTENDED |
| EN 16931 1.3.16 | CII D16B or UBL 2.1 | EN 16931 |
| XRechnung 3.0.2 (2026-08-31 rules) | CII D16B or UBL 2.1 | XRechnung |
| Peppol BIS Billing 3.0.21 | UBL 2.1 | Peppol BIS |
| Polish FA(3), schema 1-0E | National XML | VAT, KOR, ZAL, KOR_ZAL, ROZ, KOR_ROZ, UPR |

Document types 380, 381, 326, 384, 386, and 389 are mapped; Peppol BIS rejects self-billed invoices (389). `InvoiceXmlOptions.GetSupportedTargets()` lists every valid release, syntax, and profile combination. XML input and output are limited to 16 MiB and 50,000 collection items.

## Create a CII or UBL invoice

Writing checks the model arithmetic and the target mapping, and the same input always produces the same bytes. Lower Factur-X profiles carry less data: by default every populated field the profile would drop is reported as an error. Writing successfully is not proof of standards compliance; validate the bytes.

### C#

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
    Id = "1",
    Name = "Consulting",
    Quantity = 1m,
    UnitPrice = 100m,
    Tax = new InvoiceTaxCategory { Code = "S", Rate = 19m }
});

byte[] cii = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(
    InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2, InvoiceSyntax.Cii, InvoiceProfile.En16931));
byte[] ubl = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(
    InvoiceSpecificationRelease.En16931_1_3_16, InvoiceSyntax.Ubl, InvoiceProfile.En16931));

File.WriteAllBytes("invoice.cii.xml", cii);
File.WriteAllBytes("invoice.ubl.xml", ubl);
```

## Read and edit an invoice

Unknown elements and unsupported variations are reported as unmapped data, and rewriting such a source is blocked unless you accept the loss. `GetOriginalBytes()` always returns the source unchanged.

### C#

```csharp
using OfficeIMO.Invoicing;

InvoiceReadResult parsed = InvoiceParser.Read(File.ReadAllBytes("invoice.xml"));

foreach (InvoiceDiagnostic item in parsed.UnmappedData) {
    Console.WriteLine($"{item.Location}: {item.Message}");
}

if (parsed.HasCompleteMapping && parsed.Invoice.Lines.Count > 0) {
    parsed.Invoice.Lines[0].Quantity = 2m;
    InvoiceCalculator.UpdateDeclaredAmounts(parsed.Invoice);

    byte[] edited = parsed.Write(new InvoiceXmlOptions(
        InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2, InvoiceSyntax.Cii, InvoiceProfile.En16931));
    File.WriteAllBytes("edited.xml", edited);
}
```

To change only header fields of a source that has extensions, `InvoiceSourceEditor.Apply` replaces the number, dates, buyer reference, or payment reference in place and keeps everything else. The CLI `edit` command does the same.

### CLI

```sh
officeimo invoice inspect invoice.xml
officeimo invoice edit invoice.xml --output invoice.edited.xml --number INV-002 --issue-date 2026-09-30
```

## Convert between CII, UBL, and profiles

Conversion produces no bytes when data in the source cannot be represented in the target, and each diagnostic names the source path.

### C#

```csharp
using OfficeIMO.Invoicing;

InvoiceConversionResult conversion = InvoiceConverter.Convert(
    File.ReadAllBytes("invoice.xml"),
    new InvoiceXmlOptions(
        InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31, InvoiceSyntax.Ubl, InvoiceProfile.XRechnung));

if (conversion.Succeeded) {
    File.WriteAllBytes("converted.xml", conversion.Xml!);
}
```

### CLI

```sh
officeimo invoice convert invoice.xml --output invoice.ubl.xml --release En16931_1_3_16 --syntax Ubl --profile En16931
```

## Create a Polish FA(3) invoice

`Fa3InvoiceWriter` writes FA(3) from the same `Invoice` model with explicit national options. Creation time and annotations are the issuer's declarations; nothing is inferred from names or descriptions. Foreign-currency and correction kinds need explicit national amounts (`Fa3FiscalAmounts`), and exchange rates are never guessed. `Fa3InvoiceReader` reads FA(3) back with national kinds and P_13/P_14 buckets kept.

### C#

```csharp
using OfficeIMO.Invoicing;

var polish = new Invoice {
    Number = "FA-2026-001",
    IssueDate = new DateTime(2026, 9, 30),
    Currency = "PLN",
    Seller = new InvoiceParty {
        Name = "Example Seller",
        Address = new InvoiceAddress { CountryCode = "PL", Line1 = "Example Street 1", Line2 = "00-001 Warszawa" }
    },
    Buyer = new InvoiceParty {
        Name = "Example Buyer",
        Address = new InvoiceAddress { CountryCode = "PL", Line1 = "Example Street 2", Line2 = "00-001 Warszawa" }
    }
};
// Fictional identifiers; use the real NIP numbers of the parties.
polish.Seller.TaxRegistrations.Add(new InvoiceTaxRegistration("9999999999", "NIP", InvoiceTaxRegistrationKind.Fiscal));
polish.Buyer.TaxRegistrations.Add(new InvoiceTaxRegistration("1111111111", "NIP", InvoiceTaxRegistrationKind.Fiscal));
polish.Lines.Add(new InvoiceLine {
    Id = "1",
    Name = "Consulting",
    Quantity = 1,
    UnitPrice = 100,
    UnitCode = "HUR",
    Tax = new InvoiceTaxCategory { Code = "S", Rate = 23 }
});

var national = new Fa3InvoiceWriteOptions(
    Fa3InvoiceKind.TaxInvoice,
    new DateTimeOffset(2026, 9, 30, 12, 0, 0, TimeSpan.Zero),
    new Fa3InvoiceAnnotations());

foreach (InvoiceDiagnostic finding in Fa3InvoiceWriter.Inspect(polish, national)) {
    Console.WriteLine($"{finding.Code}: {finding.Message}");
}

File.WriteAllBytes("fa3.xml", Fa3InvoiceWriter.Write(polish, national));
```

## Render a PDF or a hybrid Factur-X PDF

`PdfInvoiceDocument.Create` captures one snapshot of the invoice, so the visible PDF and the embedded `factur-x.xml` cannot disagree. `ToPdfBytes` writes the hybrid PDF (PDF/A-3a, 3b, or 3u, optionally with PDF/UA-1); `ToPresentationPdfBytes` writes a visible PDF without the XML. Labels ship in ten languages. Approval blocks are printed text, not digital signatures. Peppol BIS needs UBL, so it cannot be embedded through this CII adapter.

### C#

```csharp
using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Pdf;
using OfficeIMO.Pdf;

Invoice invoice = InvoiceParser.Read(File.ReadAllBytes("invoice.cii.xml")).Invoice;

var contract = new InvoiceXmlOptions(
    InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2, InvoiceSyntax.Cii, InvoiceProfile.En16931);
PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice, contract,
    InvoicePdfLayoutOptions.ForCultures("pl-PL", "en-GB"));

byte[] font = File.ReadAllBytes("invoice-font.ttf");
PdfCreateOptions options = new PdfCreateOptions()
    .EmbedStandardFont(PdfStandardFont.Helvetica, font, "Invoice font")
    .EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "Invoice font");

File.WriteAllBytes("invoice.pdf", snapshot.ToPdfBytes(options));
```

### CLI

```sh
officeimo invoice render invoice.xml --output invoice-view.pdf --release En16931_1_3_16 --syntax Cii --profile En16931
officeimo invoice hybrid invoice.xml --output invoice.pdf --release FacturX_1_09_2_Zugferd_2_5_2 --syntax Cii --profile En16931 --language pl-PL --font ./fonts/InvoiceFont.ttf
```

## Validate against pinned rule releases

`IsValid` requires both the XSD and the business rules to pass. You download the rule archives and SaxonJ-HE yourself; they are checked by SHA-256 and never fetched by the library. Business rules need a Java runtime. `Fa3SchemaValidator` checks FA(3) against the official 1-0E XSD without Java.

### C#

```csharp
using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Validation;

InvoiceRuleBundle bundle = InvoiceRuleBundle.Load(
    "xrechnung-3.0.2-validator-configuration-2026-08-31.zip",
    facturXArchivePath: "validator-configuration-zugferd.zip");

var validator = new InvoiceValidator(bundle, new SaxonInvoiceRulesRunner("tools/saxon/saxon-he-12.10.jar"));
InvoiceValidationReport report = await validator.ValidateAsync(
    File.ReadAllBytes("invoice.xml"),
    InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31);

Console.WriteLine($"{report.IsValid}: {report.Sha256}");
```

### CLI

Without rule options, `validate` checks the model and mapping and returns exit code 1 on errors:

```sh
officeimo invoice validate invoice.xml
officeimo invoice validate invoice.xml --standards-release En16931_1_3_16 --rule-bundle ./rules/xrechnung.zip --saxon-jar ./saxon/saxon-he-12.10.jar
```

## Submit FA(3) to KSeF

`KsefClient` in `OfficeIMO.Invoicing.KSeF` authenticates with a KSeF token, opens an online session, encrypts the invoice with AES-256, and checks it against the FA(3) XSD before submission. It defaults to the `Test` environment; `Demo` and `Production` must be selected explicitly. Submissions are never retried automatically: an uncertain outcome raises `KsefMutationAmbiguousException`, and `ResumeSubmission` reattaches to a known reference. Receipts (UPO) are bound to the session, KSeF number, and invoice hash, and checked against the pinned v4-3 schema. XAdES certificate authentication, batch and offline submission, permission administration, and UPO signature verification are not included. The [KSeF README](https://github.com/EvotecIT/OfficeIMO/blob/master/OfficeIMO.Invoicing.KSeF/README.md) has the polling flow.
