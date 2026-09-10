# OfficeIMO.Invoicing

Create, read, edit and convert electronic invoices with one typed .NET model.
The core has no external runtime dependencies and supports .NET Standard 2.0,
.NET 8, .NET 10 and .NET Framework 4.7.2.

## Create invoice XML

```csharp
using OfficeIMO.Invoicing;

var invoice = new Invoice {
    Number = "INV-001",
    IssueDate = new DateTime(2026, 9, 10),
    DueDate = new DateTime(2026, 10, 10),
    Currency = "EUR",
    Seller = new InvoiceParty {
        Name = "Example Seller", VatIdentifier = "DE123456789",
        Address = new InvoiceAddress { CountryCode = "DE" }
    },
    Buyer = new InvoiceParty {
        Name = "Example Buyer",
        Address = new InvoiceAddress { CountryCode = "DE" }
    },
    Payment = new InvoicePayment { MeansCode = "58" }
};
invoice.Payment.Accounts.Add(new InvoiceBankAccount { Identifier = "DE79000000001234567890" });
invoice.Lines.Add(new InvoiceLine {
    Id = "1", Name = "Consulting", Quantity = 1m, UnitPrice = 100m,
    Tax = new InvoiceTaxCategory { Code = "S", Rate = 19m }
});
byte[] cii = InvoiceSerializer.Write(invoice);
byte[] ubl = InvoiceSerializer.Write(invoice,
    new InvoiceXmlOptions(InvoiceSyntax.Ubl, InvoiceProfile.En16931));
```

Serialization checks model arithmetic and the supported target mapping. Identical
models and options produce identical UTF-8 bytes. Use
[OfficeIMO.Invoicing.Validation](../OfficeIMO.Invoicing.Validation/README.md) to
check the resulting bytes against an explicit XSD and Schematron release.
Successful serialization alone does not establish standards compliance.

## Supported authoring contract

| Area | Supported data |
| --- | --- |
| Syntax and profiles | CII D16B and UBL 2.1; EN 16931 and XRechnung 3.0; Peppol BIS on UBL |
| Documents | Invoice and credit note; document currency and accounting-currency VAT |
| Parties | Seller, buyer, payee, tax representative, addresses, identifiers and contacts within each semantic role |
| Lines | Quantities, price base quantities, net/gross prices, discounts, allowances, charges, item identifiers, classifications and attributes |
| VAT and totals | Category/rate breakdowns, exemptions, document adjustments, prepayments and payable rounding |
| Payments | Transfer accounts, payment references, direct-debit mandate and creditor details, masked card details |
| References | Orders, preceding invoices, contracts, projects, delivery, periods, accounting and supporting documents |

Profile recognition also covers Factur-X MINIMUM, BASIC WL, BASIC and EXTENDED.
Authoring those profiles is not supported by this engine. National CIUS rules
other than XRechnung and Peppol require separate mappings and validation.

## Read and edit safely

```csharp
var parsed = InvoiceParser.Read(File.ReadAllBytes("invoice.xml"));
foreach (var item in parsed.UnmappedData)
    Console.WriteLine($"{item.Location}: {item.Message}");

if (parsed.HasCompleteMapping) {
    parsed.Invoice.Lines[0].Quantity = 2m;
    InvoiceCalculator.UpdateDeclaredAmounts(parsed.Invoice);
    File.WriteAllBytes("edited.xml", parsed.Write());
}
```

Unknown business elements, attributes and unsupported semantic variations are
reported. Rewriting blocks by default when it would discard them. The explicit
`allowUnmappedDataLoss` option accepts the reported loss; it does not preserve
unknown extensions. `GetOriginalBytes()` always returns the original bytes,
even after model edits.

Declared source line and VAT amounts are preserved. Model checks permit up to
0.02 difference from the unrounded line formula and a conservative 0.01 VAT
rounding difference; totals must match the resulting amounts exactly. These
checks do not replace release-specific rules. `UpdateDeclaredAmounts` explicitly
recalculates lines, VAT and totals after financial edits.

## Convert CII and UBL

```csharp
var conversion = InvoiceConverter.Convert(File.ReadAllBytes("invoice.xml"),
    new InvoiceXmlOptions(InvoiceSyntax.Ubl, InvoiceProfile.XRechnung));
if (conversion.Succeeded)
    File.WriteAllBytes("converted.xml", conversion.Xml!);
else
    foreach (var diagnostic in conversion.Diagnostics)
        Console.WriteLine($"{diagnostic.Location}: {diagnostic.Message}");
```

Conversion produces no bytes if observed source data cannot be represented in
the target. Examples include arbitrary tax-registration schemes, conflicting
payment descriptions, CII sales-order-only references and unsupported card
network metadata. A changed guideline is reported and requires validation
against the target release. This is a bounded semantic conversion, with explicit
loss reporting for fields outside the supported contract.

## Inspect profile declarations

```csharp
using OfficeIMO.Invoicing;

var declaration = InvoiceProfileDeclaration.Read(File.ReadAllBytes("invoice.xml"));
Console.WriteLine(declaration.GuidelineId);
if (declaration.Profile is InvoiceProfile profile && profile != InvoiceProfile.PeppolBis)
    Console.WriteLine(InvoiceProfiles.GetXmpConformanceLevel(profile));
```

The catalogue distinguishes MINIMUM, BASIC WL, BASIC, EN 16931, EXTENDED,
XRechnung 3.0, and Peppol BIS Billing 3.0. Unknown identifiers are returned with a
null profile; missing, ambiguous, malformed, or oversized input is rejected.
Peppol BIS does not define Factur-X XMP metadata.
