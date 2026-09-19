using OfficeIMO.Drawing;
using OfficeIMO.Invoicing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Builds the comparison invoice directly from OfficeIMO.Pdf composition primitives.</summary>
internal static class OfficeImoPdfInvoiceGenerator {
    private static readonly PdfColor Navy = PdfColor.FromRgb(1, 21, 52);
    private static readonly PdfColor Accent = PdfColor.FromRgb(63, 92, 255);
    private static readonly PdfColor Surface = PdfColor.FromRgb(245, 248, 252);
    private static readonly PdfColor Border = PdfColor.FromRgb(214, 223, 235);
    private static readonly PdfColor Muted = PdfColor.FromRgb(82, 99, 122);

    internal static byte[] Generate(InvoiceComparisonScenario scenario) {
        Invoice invoice = scenario.Invoice;
        PdfDocument document = PdfDocument.Create(pdf => pdf.Page(page => {
            page.Size(PageSizes.A4).Margin(42D);
            page.DefaultTextStyle(style => style.FontSize(9.5D).Color(Navy));
            page.Content(content => content.Column(column => {
                column.Item().Row(row => row
                    .Style(new PdfRowStyle { Gap = 20D, SpacingAfter = 14D })
                    .PercentColumn(44D, left => left.Image(
                        scenario.LogoBytes,
                        180D,
                        42D,
                        align: null,
                        clipPath: null,
                        fit: OfficeImageFit.Contain,
                        spacingBefore: null,
                        spacingAfter: null,
                        style: null,
                        linkUri: null,
                        linkContents: null,
                        alternativeText: "Evotec"))
                    .PercentColumn(56D, right => {
                        right.Paragraph(p => p.FontSize(18D).Bold("Invoice", Navy), PdfAlign.Right,
                            style: Tight(2D));
                        right.Paragraph(p => p.FontSize(20D).Bold(invoice.Number, Navy), PdfAlign.Right,
                            style: Tight(4D));
                        right.Paragraph(p => p.Color(Muted).Text($"Issued {invoice.IssueDate:dd/MM/yyyy} · Due {invoice.DueDate:dd/MM/yyyy}"),
                            PdfAlign.Right, style: Tight(0D));
                    }));

                column.Item().Row(row => row
                    .Style(new PdfRowStyle { Gap = 12D, SpacingAfter = 14D })
                    .PercentColumn(50D, cell => PartyCard(cell, "SELLER", invoice.Seller))
                    .PercentColumn(50D, cell => PartyCard(cell, "BUYER", invoice.Buyer)));

                column.Item().Table(new[] {
                    new[] { "Currency", invoice.Currency, "Buyer reference", invoice.BuyerReference ?? string.Empty }
                }, style: IdentityStyle());

                var lines = new List<string[]> {
                    new[] { "Item", "Quantity", "Net price", "VAT", "Net amount" }
                };
                for (int index = 0; index < invoice.Lines.Count; index++) {
                    InvoiceLine line = invoice.Lines[index];
                    lines.Add(new[] {
                        $"{line.Id}. {line.Name}\n{line.Description}",
                        InvoiceComparisonScenario.Number(line.Quantity) + " " + line.UnitCode,
                        InvoiceComparisonScenario.Number(line.UnitPrice) + " / " + InvoiceComparisonScenario.Number(line.PriceBaseQuantity) + " " + line.UnitCode,
                        line.Tax.Code + " " + InvoiceComparisonScenario.Number(line.Tax.Rate!.Value) + "%",
                        InvoiceComparisonScenario.Money(scenario.Calculation.Lines[index].NetAmount, invoice.Currency)
                    });
                }
                column.Item().Table(lines, style: LinesStyle());
                column.Item().Table(new[] {
                    new[] { "VAT category", "Taxable amount", "VAT amount", "Exemption" },
                    new[] {
                        "S 23%",
                        InvoiceComparisonScenario.Money(scenario.Calculation.TaxExclusiveTotal, invoice.Currency),
                        InvoiceComparisonScenario.Money(scenario.Calculation.TaxTotal, invoice.Currency),
                        string.Empty
                    }
                }, style: TaxStyle());

                column.Item().PageBreak();
                column.Item().Row(row => row
                    .Style(new PdfRowStyle { Gap = 18D, SpacingAfter = 16D, KeepTogether = true })
                    .PercentColumn(55D, left => left.PanelParagraph(p => p
                        .Bold("Payment reference\n", Navy)
                        .Color(Muted).Text(invoice.Number + "\n" + invoice.PaymentTerms), CardStyle()))
                    .PercentColumn(45D, right => right.Table(TotalRows(scenario), style: TotalsStyle())));

                column.Item().H2("References", PdfAlign.Left, Navy);
                column.Item().Table(new[] {
                    new[] { "Purchase order", invoice.PurchaseOrderReference ?? string.Empty },
                    new[] { "Document type", invoice.TypeCode }
                }, style: DetailsStyle());
                column.Item().H2("Payment", PdfAlign.Left, Navy);
                column.Item().Table(new[] {
                    new[] { "Payment means", invoice.Payments[0].MeansCode + " " + invoice.Payments[0].MeansText },
                    new[] { "Reference", invoice.Payments[0].Reference ?? string.Empty },
                    new[] { "IBAN", invoice.Payments[0].Account!.Identifier + "\n" + invoice.Payments[0].Account!.Name }
                }, style: DetailsStyle());
                column.Item().Paragraph(p => p.Text(invoice.PaymentTerms ?? string.Empty));
                column.Item().H2("Notes", PdfAlign.Left, Navy);
                column.Item().Paragraph(p => p.Text(invoice.Notes[0].Text));
                column.Item().H2("Approvals", PdfAlign.Left, Navy);
                column.Item().Row(row => row
                    .Style(new PdfRowStyle { Gap = 12D })
                    .PercentColumn(50D, cell => Approval(cell, "PREPARED BY", "Marta Nowak", "Finance · 17/09/2026"))
                    .PercentColumn(50D, cell => Approval(cell, "APPROVED BY", "Daniel Reed", "Delivery lead · 17/09/2026")));
            }));
        }), scenario.PdfOptions).Meta(title: "Evotec invoice PDF renderer comparison");
        return document.ToBytes();
    }

    private static void PartyCard(PdfContentBuilder content, string label, InvoiceParty party) => content.Panel(card => {
        card.Paragraph(p => p.Bold(label, Accent), style: Tight(4D));
        card.Paragraph(p => p.Text(party.Name + "\n" + InvoiceComparisonScenario.PartyDetails(party)), style: Tight(0D));
    }, CardStyle());

    private static void Approval(PdfContentBuilder content, string label, string name, string detail) => content.Panel(card => {
        card.Paragraph(p => p.Color(Muted).Text(label), style: Tight(9D));
        card.HR(0.8D, Border, 0D, 6D);
        card.Paragraph(p => p.Bold(name, Navy), style: Tight(3D));
        card.Paragraph(p => p.Color(Muted).Text(detail), style: Tight(0D));
    }, CardStyle());

    private static string[][] TotalRows(InvoiceComparisonScenario scenario) => new[] {
        new[] { "Line net total", InvoiceComparisonScenario.Money(scenario.Calculation.LineNetTotal, scenario.Invoice.Currency) },
        new[] { "Total excluding VAT", InvoiceComparisonScenario.Money(scenario.Calculation.TaxExclusiveTotal, scenario.Invoice.Currency) },
        new[] { "VAT total", InvoiceComparisonScenario.Money(scenario.Calculation.TaxTotal, scenario.Invoice.Currency) },
        new[] { "Total including VAT", InvoiceComparisonScenario.Money(scenario.Calculation.TaxInclusiveTotal, scenario.Invoice.Currency) },
        new[] { "Prepaid", InvoiceComparisonScenario.Money(scenario.Calculation.PrepaidAmount, scenario.Invoice.Currency) },
        new[] { "Amount due", InvoiceComparisonScenario.Money(scenario.Calculation.PayableAmount, scenario.Invoice.Currency) }
    };

    private static PdfParagraphStyle Tight(double after) => new() { SpacingBefore = 0D, SpacingAfter = after };

    private static PdfPanelStyle CardStyle() => new() {
        Background = Surface, BorderColor = Border, BorderWidth = 0.6D, CornerRadius = 7D,
        PaddingX = 10D, PaddingY = 9D, SpacingAfter = 0D
    };

    private static PdfTableStyle IdentityStyle() => new() {
        HeaderRowCount = 0, RowStripeFill = null, BorderColor = Border, BorderWidth = 0.7D,
        CornerRadius = 7D, CellPaddingX = 10D, CellPaddingY = 7D, FontSize = 9D, SpacingAfter = 14D,
        BodyColumnFills = new List<PdfColor?> { Surface, null, Surface, null },
        ColumnWidthWeights = new List<double> { 1.1D, 1.4D, 1.4D, 2.1D }
    };

    private static PdfTableStyle LinesStyle() => new() {
        HeaderRowCount = 1, HeaderFill = Navy, HeaderTextColor = PdfColor.White, HeaderFontSize = 9D,
        FontSize = 9D, BorderColor = Border, BorderWidth = 0.5D, CornerRadius = 7D, RowStripeFill = Surface,
        CellPaddingX = 7D, CellPaddingY = 7D, SpacingAfter = 12D,
        ColumnWidthWeights = new List<double> { 3.7D, 1.15D, 1.55D, 1D, 1.6D },
        Alignments = new List<PdfColumnAlign> { PdfColumnAlign.Left, PdfColumnAlign.Right, PdfColumnAlign.Right, PdfColumnAlign.Right, PdfColumnAlign.Right }
    };

    private static PdfTableStyle TaxStyle() => new() {
        HeaderRowCount = 1, HeaderFill = Navy, HeaderTextColor = PdfColor.White, HeaderFontSize = 8.5D,
        FontSize = 8.5D, BorderColor = Border, BorderWidth = 0.5D, CornerRadius = 7D, RowStripeFill = Surface,
        CellPaddingX = 7D, CellPaddingY = 6D, SpacingAfter = 12D,
        ColumnWidthWeights = new List<double> { 1.2D, 1.5D, 1.4D, 2D },
        Alignments = new List<PdfColumnAlign> { PdfColumnAlign.Left, PdfColumnAlign.Right, PdfColumnAlign.Right, PdfColumnAlign.Left }
    };

    private static PdfTableStyle TotalsStyle() => new() {
        HeaderRowCount = 0, FontSize = 10D, BorderColor = Border, BorderWidth = 0.6D, CornerRadius = 7D,
        FooterRowCount = 1, FooterFill = Accent, FooterTextColor = PdfColor.White, RowStripeFill = Surface,
        CellPaddingX = 9D, CellPaddingY = 7D,
        Alignments = new List<PdfColumnAlign> { PdfColumnAlign.Left, PdfColumnAlign.Right }
    };

    private static PdfTableStyle DetailsStyle() => new() {
        HeaderRowCount = 0, FontSize = 9D, BorderColor = Border, BorderWidth = 0.5D, CornerRadius = 7D,
        RowStripeFill = Surface, CellPaddingX = 7D, CellPaddingY = 5D, SpacingAfter = 6D,
        ColumnWidthWeights = new List<double> { 1D, 1D }
    };
}
