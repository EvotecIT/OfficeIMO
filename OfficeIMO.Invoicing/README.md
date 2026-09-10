# OfficeIMO.Invoicing

Create, read, edit and convert electronic invoices with one typed .NET model.
The core has no external runtime dependencies and supports .NET Standard 2.0,
.NET 8, .NET 10 and .NET Framework 4.7.2.

## Build and install locally

From the repository root, pack the project into a local feed, then add that
package to your application:

```powershell
dotnet pack OfficeIMO.Invoicing/OfficeIMO.Invoicing.csproj -c Release -o artifacts/invoice-feed
dotnet add path/to/Application.csproj package OfficeIMO.Invoicing --source artifacts/invoice-feed
```

This package does not depend on a PDF engine or an external validation runtime.

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
| Documents | Invoice and credit note; document currency and accounting-currency VAT. UBL credit notes cannot carry a due date or project reference; writing or converting those fields reports unsupported target data. |
| Parties | Seller, buyer, payee, tax representative, addresses, identifiers and contacts within each semantic role |
| Lines | Quantities, price base quantities, net/gross prices, discounts, allowances, charges, item identifiers, classifications and attributes |
| VAT and totals | Category/rate breakdowns, exemptions, document adjustments, prepayments and payable rounding |
| Payments | Transfer accounts, payment references, direct-debit mandate and creditor details, masked card details |
| References | Orders, preceding invoices, contracts, projects, delivery, periods, accounting and supporting documents; external locations preserve well-formed absolute URIs, including FTP and URN schemes, without fetching them |

Profile recognition also covers Factur-X MINIMUM, BASIC WL, BASIC, EXTENDED and EXTENDED-CTC-FR.
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
rounding difference for taxable categories; zero-tax categories require exactly
zero VAT. Totals must match the resulting amounts exactly. These
checks do not replace release-specific rules. `UpdateDeclaredAmounts` explicitly
recalculates lines, VAT and totals after financial edits.

Model validation requires a seller business, legal or VAT identifier, and an
account for credit-transfer payment codes 30 and 58. VAT checks cover category-specific
registration requirements, exemption reasons on the resulting breakdown, and
intra-community delivery details. Imported header-only exemption reasons remain
valid through recalculation. Outside-scope VAT cannot mix with other categories
or carry party VAT identifiers. Official code lists and release-specific rules
remain the responsibility of the optional standards validator.

Quantity, price, base-quantity and percentage calculations retain intermediate
precision until monetary rounding. Monetary results use two decimal places with
ties towards positive infinity; values that exceed decimal capacity are rejected
instead of silently losing cents. Caller text must contain valid XML characters.
The model permits up to 4 MiB of combined UTF-8 text and 8 MiB of embedded bytes.
Parsing and authoring share a 50,000-item budget across collections, including
declared or calculated VAT breakdowns. Parsing rejects inputs with more than
1,000 mapping diagnostics. Diagnostic codes are bounded to 256 characters;
messages and locations to 4,096, with an explicit truncation marker when needed.
Serialized XML is limited to 16 MiB while it is written.

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
XRechnung 3.0, EXTENDED-CTC-FR, and Peppol BIS Billing 3.0. Unknown identifiers are returned with a
null profile; missing, ambiguous, malformed, or oversized input is rejected.
Peppol BIS does not define Factur-X XMP metadata.

The French profile identifier follows [AFNOR XP Z12-012](https://www.impots.gouv.fr/sites/default/files/media/1_metier/2_professionnel/EV/2_gestion/290_facturation_electronique/specification_externes_b2b/afnor/norme-afnor-factures.pdf).
Recognition does not establish French business-rule coverage.
