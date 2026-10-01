# OfficeIMO.Invoicing.Pdf

Generate a visible PDF and an embedded CII invoice from one captured typed invoice.
This optional adapter depends on `OfficeIMO.Invoicing` and `OfficeIMO.Pdf`.
XML-only applications can use `OfficeIMO.Invoicing`; PDF-only applications can use
`OfficeIMO.Pdf`. Neither engine depends on this adapter or on the other engine.

## Install

```powershell
dotnet add package OfficeIMO.Invoicing.Pdf --version 3.4.4
```

## Build from source

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
PDF presentation accepts document types 380 (invoice), 381 (credit note), 326
(partial invoice), 384 (corrected invoice), 386 (prepayment invoice) and 389
(self-billed invoice), and rejects other types before creating the snapshot.
The layout uses a translated heading for each type and includes repeated line-table headers,
VAT and payable totals, party details, payment instructions and references.
Totals stay together when page space permits. Generated labels are available in
English, German, Polish, French, Spanish, Italian, Dutch, Portuguese, Czech and
Slovak. `ForCultures` combines packs in order for a
bilingual or multilingual layout and formats values with the first culture.
`InvoicePdfLanguagePack.Create` supports partial custom translations with explicit
English fallback. Layout options are captured with the invoice, so later caller
changes cannot alter an existing snapshot.

## Choose columns and continuation-page details

Both layouts share ordered column selection, unit/payment descriptions and page
identity settings. Keep `Item` and `NetAmount`, choose two to eight distinct
columns, and use a wider page when the selected columns need more space:

```csharp
var layout = InvoicePdfLayoutOptions.ForCultures("pl-PL", "en-GB");
layout.Theme = InvoicePdfTheme.Modern(PdfColor.FromRgb(63, 92, 255));
layout.LineColumns.Clear();
foreach (var column in new[] {
    InvoicePdfLineColumn.Item, InvoicePdfLineColumn.Quantity,
    InvoicePdfLineColumn.Unit, InvoicePdfLineColumn.NetPrice,
    InvoicePdfLineColumn.Vat, InvoicePdfLineColumn.NetAmount
}) layout.LineColumns.Add(column);
layout.UnitCodeDisplay = InvoicePdfCodeDisplay.CodeAndDescription;
layout.PaymentCodeDisplay = InvoicePdfCodeDisplay.Description;
layout.CompactDetails = true;
layout.IncludePageIdentity = true;
var snapshot = PdfInvoiceDocument.Create(invoice, contract, layout);
File.WriteAllBytes("invoice.pdf", snapshot.ToPdfBytes(options));
```

Optional columns also include line ID, description, service period, seller/buyer
item IDs, standard item ID, accounting reference, gross price and price discount.
Business details assigned to those columns appear there once; the `Item` column
retains the remaining item metadata. Column selection controls the visible table
and does not change the embedded XML.

The default code display preserves raw codes. Translated descriptions cover
common UN/ECE units (`C62`, `HUR`, `DAY`, `WEE`, `MON`, `KGM`, `MTR`, `MTK`, `LTR`)
and UNCL 4461 payment codes (`10`, `30`, `48`, `49`, `58`, `59`) in the ten built-in
languages. Unknown codes remain visible, and authored payment text is retained.
Use `InvoicePdfLanguagePack.WithCodeDescriptions` to supply additional descriptions:

```csharp
var pack = InvoicePdfLanguagePack.ForCulture("en-GB").WithCodeDescriptions(
    units: new Dictionary<string, string> { ["XBX"] = "box" });
layout.Languages.Clear();
layout.Languages.Add(pack);
```

Compact details reduce heading spacing and table padding while keeping all
values. Page identity is opt-in and requires a single-line invoice number up to
80 characters. It replaces footer text on all pages with the invoice number and
localized page count, while retaining configured first/even headers and footer
graphics. Invoice numbers are literal text, including braces. Leave it disabled
to keep a caller-supplied footer. Defaults retain the five established line
columns and existing spacing.

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
