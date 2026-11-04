using OfficeIMO.Drawing;
using OfficeIMO.Invoicing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Builds the comparison invoice directly from OfficeIMO.Pdf composition primitives.</summary>
internal static class OfficeImoPdfInvoiceGenerator {
    private static readonly global::OfficeIMO.OfficeColor Navy = global::OfficeIMO.OfficeColor.FromRgb(1, 21, 52);
    private static readonly global::OfficeIMO.OfficeColor Accent = global::OfficeIMO.OfficeColor.FromRgb(63, 92, 255);
    private static readonly global::OfficeIMO.OfficeColor Surface = global::OfficeIMO.OfficeColor.FromRgb(245, 248, 252);
    private static readonly global::OfficeIMO.OfficeColor Border = global::OfficeIMO.OfficeColor.FromRgb(214, 223, 235);
    private static readonly global::OfficeIMO.OfficeColor Muted = global::OfficeIMO.OfficeColor.FromRgb(82, 99, 122);

    internal static byte[] Generate(InvoiceComparisonScenario scenario) {
        Invoice invoice = scenario.Invoice;
        PdfDocument document = PdfDocument.Create(pdf => pdf.Page(page => {
            page.Size(PdfPageSizes.A4).Margin(global::OfficeIMO.OfficeLength.Points(42D));
            page.DefaultTextStyle(style => style.FontSize(global::OfficeIMO.OfficeLength.Points(9.5D)).Color(Navy));
            page.Content(content => content.Column(column => {
                column.Item().Row(row => row
                    .Style(new PdfRowStyle { Gap = global::OfficeIMO.OfficeLength.Points(20D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(14D) })
                    .PercentColumn(44D, left => left.AddImage(
                        scenario.LogoBytes,
                        global::OfficeIMO.OfficeLength.Points(180D),
                        global::OfficeIMO.OfficeLength.Points(42D),
                        align: null,
                        clipPath: null,
                        fit: OfficeImageFit.Contain,
                        spacingBefore: null,
                        spacingAfter: null,
                        style: null,
                        linkUri: null,
                        linkContents: null,
                        altText: "Evotec"))
                    .PercentColumn(56D, right => {
                        right.AddParagraph(p => p.AddText("Invoice").SetBold().SetColor(Navy).SetFontSize(global::OfficeIMO.OfficeLength.Points(18D)), OfficeAlignment.Right,
                            style: Tight(2D));
                        right.AddParagraph(p => p.AddText(invoice.Number).SetBold().SetColor(Navy).SetFontSize(global::OfficeIMO.OfficeLength.Points(20D)), OfficeAlignment.Right,
                            style: Tight(4D));
                        right.AddParagraph(p => p.AddText($"Issued {invoice.IssueDate:dd/MM/yyyy} · Due {invoice.DueDate:dd/MM/yyyy}").SetColor(Muted),
                            OfficeAlignment.Right, style: Tight(0D));
                    }));

                column.Item().Row(row => row
                    .Style(new PdfRowStyle { Gap = global::OfficeIMO.OfficeLength.Points(12D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(14D) })
                    .PercentColumn(50D, cell => PartyCard(cell, "SELLER", invoice.Seller))
                    .PercentColumn(50D, cell => PartyCard(cell, "BUYER", invoice.Buyer)));

                column.Item().AddTable(new[] {
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
                column.Item().AddTable(lines, style: LinesStyle());
                column.Item().AddTable(new[] {
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
                    .Style(new PdfRowStyle { Gap = global::OfficeIMO.OfficeLength.Points(18D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(16D), KeepTogether = true })
                    .PercentColumn(55D, left => left.AddPanelParagraph(p => {
                        p.AddText("Payment reference\n").SetBold().SetColor(Navy);
                        p.AddText(invoice.Number + "\n" + invoice.PaymentTerms).SetColor(Muted);
                    }, CardStyle()))
                    .PercentColumn(45D, right => right.AddTable(TotalRows(scenario), style: TotalsStyle())));

                column.Item().H2("References", OfficeAlignment.Left, Navy);
                column.Item().AddTable(new[] {
                    new[] { "Purchase order", invoice.PurchaseOrderReference ?? string.Empty },
                    new[] { "Document type", invoice.TypeCode }
                }, style: DetailsStyle());
                column.Item().H2("Payment", OfficeAlignment.Left, Navy);
                column.Item().AddTable(new[] {
                    new[] { "Payment means", invoice.Payments[0].MeansCode + " " + invoice.Payments[0].MeansText },
                    new[] { "Reference", invoice.Payments[0].Reference ?? string.Empty },
                    new[] { "IBAN", invoice.Payments[0].Account!.Identifier + "\n" + invoice.Payments[0].Account!.Name }
                }, style: DetailsStyle());
                column.Item().AddParagraph(p => p.AddText(invoice.PaymentTerms ?? string.Empty));
                column.Item().H2("Notes", OfficeAlignment.Left, Navy);
                column.Item().AddParagraph(p => p.AddText(invoice.Notes[0].Text));
                column.Item().H2("Approvals", OfficeAlignment.Left, Navy);
                column.Item().Row(row => row
                    .Style(new PdfRowStyle { Gap = global::OfficeIMO.OfficeLength.Points(12D) })
                    .PercentColumn(50D, cell => Approval(cell, "PREPARED BY", "Marta Nowak", "Finance · 17/09/2026"))
                    .PercentColumn(50D, cell => Approval(cell, "APPROVED BY", "Daniel Reed", "Delivery lead · 17/09/2026")));
            }));
        }), scenario.PdfOptions).Meta(title: "Evotec invoice PDF renderer comparison");
        return document.ToBytes();
    }

    private static void PartyCard(PdfContentBuilder content, string label, InvoiceParty party) => content.Panel(card => {
        card.AddParagraph(p => p.AddText(label).SetBold().SetColor(Accent), style: Tight(4D));
        card.AddParagraph(p => p.AddText(party.Name + "\n" + InvoiceComparisonScenario.PartyDetails(party)), style: Tight(0D));
    }, CardStyle());

    private static void Approval(PdfContentBuilder content, string label, string name, string detail) => content.Panel(card => {
        card.AddParagraph(p => p.AddText(label).SetColor(Muted), style: Tight(9D));
        card.HR(global::OfficeIMO.OfficeLength.Points(0.8D), Border, global::OfficeIMO.OfficeLength.Points(0D), global::OfficeIMO.OfficeLength.Points(6D));
        card.AddParagraph(p => p.AddText(name).SetBold().SetColor(Navy), style: Tight(3D));
        card.AddParagraph(p => p.AddText(detail).SetColor(Muted), style: Tight(0D));
    }, CardStyle());

    private static string[][] TotalRows(InvoiceComparisonScenario scenario) => new[] {
        new[] { "Line net total", InvoiceComparisonScenario.Money(scenario.Calculation.LineNetTotal, scenario.Invoice.Currency) },
        new[] { "Total excluding VAT", InvoiceComparisonScenario.Money(scenario.Calculation.TaxExclusiveTotal, scenario.Invoice.Currency) },
        new[] { "VAT total", InvoiceComparisonScenario.Money(scenario.Calculation.TaxTotal, scenario.Invoice.Currency) },
        new[] { "Total including VAT", InvoiceComparisonScenario.Money(scenario.Calculation.TaxInclusiveTotal, scenario.Invoice.Currency) },
        new[] { "Prepaid", InvoiceComparisonScenario.Money(scenario.Calculation.PrepaidAmount, scenario.Invoice.Currency) },
        new[] { "Amount due", InvoiceComparisonScenario.Money(scenario.Calculation.PayableAmount, scenario.Invoice.Currency) }
    };

    private static PdfParagraphStyle Tight(double after) => new() { SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(after) };

    private static PdfPanelStyle CardStyle() => new() {
        Background = Surface, BorderColor = Border, BorderWidth = global::OfficeIMO.OfficeLength.Points(0.6D), CornerRadius = global::OfficeIMO.OfficeLength.Points(7D),
        PaddingX = global::OfficeIMO.OfficeLength.Points(10D), PaddingY = global::OfficeIMO.OfficeLength.Points(9D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D)
    };

    private static PdfTableStyle IdentityStyle() => new() {
        HeaderRowCount = 0, RowStripeFill = null, BorderColor = Border, BorderWidth = global::OfficeIMO.OfficeLength.Points(0.7D),
        CornerRadius = global::OfficeIMO.OfficeLength.Points(7D), CellPaddingX = global::OfficeIMO.OfficeLength.Points(10D), CellPaddingY = global::OfficeIMO.OfficeLength.Points(7D), FontSize = global::OfficeIMO.OfficeLength.Points(9D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(14D),
        BodyColumnFills = new List<global::OfficeIMO.OfficeColor?> { Surface, null, Surface, null },
        ColumnWidthWeights = new List<double> { 1.1D, 1.4D, 1.4D, 2.1D }
    };

    private static PdfTableStyle LinesStyle() => new() {
        HeaderRowCount = 1, HeaderFill = Navy, HeaderTextColor = global::OfficeIMO.OfficeColor.White, HeaderFontSize = global::OfficeIMO.OfficeLength.Points(9D),
        FontSize = global::OfficeIMO.OfficeLength.Points(9D), BorderColor = Border, BorderWidth = global::OfficeIMO.OfficeLength.Points(0.5D), CornerRadius = global::OfficeIMO.OfficeLength.Points(7D), RowStripeFill = Surface,
        CellPaddingX = global::OfficeIMO.OfficeLength.Points(7D), CellPaddingY = global::OfficeIMO.OfficeLength.Points(7D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(12D),
        ColumnWidthWeights = new List<double> { 3.7D, 1.15D, 1.55D, 1D, 1.6D },
        Alignments = new List<OfficeAlignment> { OfficeAlignment.Left, OfficeAlignment.Right, OfficeAlignment.Right, OfficeAlignment.Right, OfficeAlignment.Right }
    };

    private static PdfTableStyle TaxStyle() => new() {
        HeaderRowCount = 1, HeaderFill = Navy, HeaderTextColor = global::OfficeIMO.OfficeColor.White, HeaderFontSize = global::OfficeIMO.OfficeLength.Points(8.5D),
        FontSize = global::OfficeIMO.OfficeLength.Points(8.5D), BorderColor = Border, BorderWidth = global::OfficeIMO.OfficeLength.Points(0.5D), CornerRadius = global::OfficeIMO.OfficeLength.Points(7D), RowStripeFill = Surface,
        CellPaddingX = global::OfficeIMO.OfficeLength.Points(7D), CellPaddingY = global::OfficeIMO.OfficeLength.Points(6D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(12D),
        ColumnWidthWeights = new List<double> { 1.2D, 1.5D, 1.4D, 2D },
        Alignments = new List<OfficeAlignment> { OfficeAlignment.Left, OfficeAlignment.Right, OfficeAlignment.Right, OfficeAlignment.Left }
    };

    private static PdfTableStyle TotalsStyle() => new() {
        HeaderRowCount = 0, FontSize = global::OfficeIMO.OfficeLength.Points(10D), BorderColor = Border, BorderWidth = global::OfficeIMO.OfficeLength.Points(0.6D), CornerRadius = global::OfficeIMO.OfficeLength.Points(7D),
        FooterRowCount = 1, FooterFill = Accent, FooterTextColor = global::OfficeIMO.OfficeColor.White, RowStripeFill = Surface,
        CellPaddingX = global::OfficeIMO.OfficeLength.Points(9D), CellPaddingY = global::OfficeIMO.OfficeLength.Points(7D),
        Alignments = new List<OfficeAlignment> { OfficeAlignment.Left, OfficeAlignment.Right }
    };

    private static PdfTableStyle DetailsStyle() => new() {
        HeaderRowCount = 0, FontSize = global::OfficeIMO.OfficeLength.Points(9D), BorderColor = Border, BorderWidth = global::OfficeIMO.OfficeLength.Points(0.5D), CornerRadius = global::OfficeIMO.OfficeLength.Points(7D),
        RowStripeFill = Surface, CellPaddingX = global::OfficeIMO.OfficeLength.Points(7D), CellPaddingY = global::OfficeIMO.OfficeLength.Points(5D), SpacingAfter = global::OfficeIMO.OfficeLength.Points(6D),
        ColumnWidthWeights = new List<double> { 1D, 1D }
    };
}
