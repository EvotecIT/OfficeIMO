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
var contract = new InvoiceXmlOptions(
    InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2,
    InvoiceSyntax.Cii,
    InvoiceProfile.En16931);
var layout = InvoicePdfLayoutOptions.ForCultures("pl-PL", "en-GB");
var snapshot = PdfInvoiceDocument.Create(invoice, contract, layout);
byte[] font = File.ReadAllBytes("invoice-font.ttf");
var options = new PdfOptions()
    .EmbedStandardFont(PdfStandardFont.Helvetica, font, "Invoice font")
    .EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "Invoice font");
File.WriteAllBytes("invoice.xml", snapshot.ToXmlBytes());
File.WriteAllBytes("invoice.pdf", snapshot.ToPdfBytes(options));
```

Add a logo, a color theme, rounded cards, and visible approval details without
rebuilding the invoice as low-level PDF primitives:

```csharp
var layout = InvoicePdfLayoutOptions.ForCultures("en-GB");
layout.Theme = InvoicePdfTheme.Modern(PdfColor.FromRgb(63, 92, 255));
layout.LogoBytes = File.ReadAllBytes("company-logo.png");
layout.LogoAlternativeText = "Company logo";
layout.Approvals.Add(new InvoicePdfApproval(
    "Prepared by", "Marta Nowak", "Finance", invoice.IssueDate));
layout.Approvals.Add(new InvoicePdfApproval(
    "Approved by", "Daniel Reed", "Delivery lead", invoice.IssueDate));

var branded = PdfInvoiceDocument.Create(invoice, contract, layout);
File.WriteAllBytes("invoice.pdf", branded.ToPdfBytes(options));
```

The modern theme is opt-in; existing layouts keep their established appearance.
Approval blocks are printed identity and date fields. They are not PDF digital
signatures and do not provide cryptographic proof. For a visible PDF without the
CII attachment, such as a presentation-only comparison, call
`ToPresentationPdfBytes`. Presentation-only rendering rejects `PdfOptions` already configured for Factur-X/ZUGFeRD so it cannot silently retain an unrelated electronic-invoice attachment. Use `ToPdfBytes` for the hybrid Factur-X/ZUGFeRD output.

Later edits to `invoice` cannot change the snapshot. `ToInvoice()` returns an
independent editable model. Reuse the captured `Release` and `Profile` when
creating a new snapshot after edits:

```csharp
Invoice edited = snapshot.ToInvoice();
edited.Number = "INV-2026-002";
var updatedContract = new InvoiceXmlOptions(snapshot.Release, InvoiceSyntax.Cii, snapshot.Profile);
var updated = PdfInvoiceDocument.Create(edited, updatedContract, layout);
```

The editable model contains business data; the snapshot's release and profile
identify its exact CII authoring contract. There is no implicit release overload.
The PDF uses the
same declared amounts and calculation as its XML, includes `factur-x.xml` as an
alternative representation, and derives its XMP profile from that attachment.
PDF presentation accepts document type 380 (invoice) and 381 (credit note), and
rejects other document types before creating the snapshot. The layout includes
invoice and credit-note headings, repeated line-table headers,
VAT and payable totals, party details, payment instructions and references.
Totals stay together when page space permits. Generated labels are available in
English, German, Polish, French, Spanish, Italian, Dutch, Portuguese, Czech and
Slovak. `ForCultures` combines packs in order for a
bilingual or multilingual layout and formats values with the first culture.
`InvoicePdfLanguagePack.Create` supports partial custom translations with explicit
English fallback. Layout options are captured with the invoice, so later caller
changes cannot alter an existing snapshot.

`ToPdfBytes` enables the shared multilingual font-fallback planner and uses the
managed Arabic joining and bidirectional shaping provider unless the caller supplies
a different shaping provider. For portable, repeatable output, pass embedded
fonts that cover the actual scripts through
`PdfOptions.RegisterEmbeddedFontFallbacks`; system font discovery is suitable only
when the deployment owns and verifies the installed families. Long descriptions
flow across pages instead of being truncated.

Use [pinned XML validation](../OfficeIMO.Invoicing.Validation/README.md) and run
veraPDF plus an invoice validator against the exact generated PDF. Font coverage,
page layout and external validation remain necessary for the selected content and
profile. Tests compare the exact embedded XML bytes, parsed semantic values and
visible mixed-script PDF text; external validation and human visual inspection
remain required for a release artifact.

The adapter accepts explicit CII contracts. Factur-X EN 16931 and EXTENDED retain
the complete authored line model. Lower Factur-X profiles use their declared
projection policy and therefore show only the business data retained by their XML
contract. Peppol BIS requires UBL and cannot be embedded through this CII adapter.
