# OfficeIMO.Invoicing.Pdf

Generate a visible PDF and an embedded CII invoice from one captured typed invoice.
This optional adapter depends on `OfficeIMO.Invoicing` and `OfficeIMO.Pdf`.
XML-only applications can use `OfficeIMO.Invoicing`; PDF-only applications can use
`OfficeIMO.Pdf`. Neither engine depends on this adapter or on the other engine.

## Build and install locally

From the repository root, pack the adapter and its first-party dependencies into
a local feed, then add the adapter package to your application:

```powershell
dotnet pack OfficeIMO.Core/OfficeIMO.Core.csproj -c Release -o artifacts/invoice-feed
dotnet pack OfficeIMO.Pdf/OfficeIMO.Pdf.csproj -c Release -o artifacts/invoice-feed
dotnet pack OfficeIMO.Invoicing/OfficeIMO.Invoicing.csproj -c Release -o artifacts/invoice-feed
dotnet pack OfficeIMO.Invoicing.Pdf/OfficeIMO.Invoicing.Pdf.csproj -c Release -o artifacts/invoice-feed
dotnet add path/to/Application.csproj package OfficeIMO.Invoicing.Pdf --source artifacts/invoice-feed
```

Targets: .NET Standard 2.0, .NET 8, .NET 10, and .NET Framework 4.7.2.
No external PDF renderer or Java runtime is required to generate the documents.

## Generate a PDF and its XML from one snapshot

Create a populated invoice using the [invoice model example](../OfficeIMO.Invoicing/README.md#create-invoice-xml),
then capture prices, quantities, VAT breakdowns, adjustments, payments and totals:

```csharp
using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Pdf;
using OfficeIMO.Pdf;

// invoice is a populated OfficeIMO.Invoicing.Invoice.
var snapshot = PdfInvoiceDocument.Create(invoice);
byte[] font = File.ReadAllBytes("invoice-font.ttf");
var options = new PdfOptions()
    .EmbedStandardFont(PdfStandardFont.Helvetica, font, "Invoice font")
    .EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "Invoice font");
File.WriteAllBytes("invoice.xml", snapshot.ToXmlBytes());
File.WriteAllBytes("invoice.pdf", snapshot.ToPdfBytes(options));
```

Later edits to `invoice` cannot change the snapshot. `ToInvoice()` returns an
independent editable model; create a new snapshot after edits. The PDF uses the
same declared amounts and calculation as its XML, includes `factur-x.xml` as an
alternative representation, and derives its XMP profile from that attachment.
PDF presentation accepts document type 380 (invoice) and 381 (credit note), and
rejects other document types before creating the snapshot. The layout includes
invoice and credit-note headings, repeated line-table headers,
VAT and payable totals, party details, payment instructions and references.
Totals stay together when page space permits.

Use [pinned XML validation](../OfficeIMO.Invoicing.Validation/README.md) and run
veraPDF plus an invoice validator against the exact generated PDF. Font coverage,
page layout and external validation remain necessary for the selected content and
profile. This API generates English labels and ISO dates; localized templates and
additional national profiles are separate capabilities.

The adapter accepts CII authoring profiles EN 16931 (the default) and XRechnung. Other Factur-X levels are inspection-only contracts in the invoice engine, and Peppol BIS authoring requires UBL.
