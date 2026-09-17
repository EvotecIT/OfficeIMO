using iText.IO.Font;
using iText.IO.Image;
using iText.Kernel.Colors;
using iText.Kernel.Font;
using iText.Kernel.Geom;
using iText.Kernel.Pdf;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using OfficeIMO.Invoicing;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class ITextPdfInvoiceGenerator {
    private static readonly DeviceRgb Navy = new(1, 21, 52);
    private static readonly DeviceRgb Accent = new(63, 92, 255);
    private static readonly DeviceRgb Surface = new(245, 248, 252);
    private static readonly DeviceRgb Border = new(214, 223, 235);

    internal static byte[] Generate(InvoiceComparisonScenario scenario) {
        using var output = new MemoryStream();
        var writer = new PdfWriter(output, new WriterProperties().SetCompressionLevel(6));
        var pdf = new iText.Kernel.Pdf.PdfDocument(writer);
        var document = new Document(pdf, iText.Kernel.Geom.PageSize.A4);
        document.SetMargins(42, 42, 42, 42);
        PdfFont regular = PdfFontFactory.CreateFont(scenario.RegularFont, PdfEncodings.IDENTITY_H, PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED);
        PdfFont bold = PdfFontFactory.CreateFont(scenario.BoldFont, PdfEncodings.IDENTITY_H, PdfFontFactory.EmbeddingStrategy.PREFER_EMBEDDED);
        document.SetFont(regular).SetFontSize(9.5f).SetFontColor(Navy);

        Invoice invoice = scenario.Invoice;
        var header = new Table(UnitValue.CreatePercentArray(new float[] { 1, 1 })).UseAllAvailableWidth().SetBorder(null);
        header.AddCell(NoBorder(new Cell()).Add(new Image(ImageDataFactory.Create(scenario.LogoBytes)).ScaleToFit(180, 42)));
        var title = new Paragraph().SetTextAlignment(TextAlignment.RIGHT)
            .Add(new Text("Invoice\n").SetFont(bold).SetFontSize(18))
            .Add(new Text(invoice.Number + "\n").SetFont(bold).SetFontSize(20))
            .Add(new Text($"Issued {invoice.IssueDate:dd/MM/yyyy} · Due {invoice.DueDate:dd/MM/yyyy}").SetFontColor(new DeviceRgb(82, 99, 122)));
        header.AddCell(NoBorder(new Cell()).Add(title));
        document.Add(header.SetMarginBottom(14));

        var parties = new Table(UnitValue.CreatePercentArray(new float[] { 1, 1 })).UseAllAvailableWidth().SetHorizontalBorderSpacing(12);
        parties.AddCell(PartyCell("SELLER", invoice.Seller, bold));
        parties.AddCell(PartyCell("BUYER", invoice.Buyer, bold));
        document.Add(parties.SetMarginBottom(14));

        var identity = new Table(UnitValue.CreatePercentArray(new float[] { 1, 1.25f, 1.25f, 1.9f })).UseAllAvailableWidth();
        AddBody(identity, "Currency", Surface, TextAlignment.LEFT);
        AddBody(identity, invoice.Currency, ColorConstants.WHITE, TextAlignment.LEFT);
        AddBody(identity, "Buyer reference", Surface, TextAlignment.LEFT);
        AddBody(identity, invoice.BuyerReference ?? string.Empty, ColorConstants.WHITE, TextAlignment.LEFT);
        document.Add(identity.SetMarginBottom(14));

        var lines = new Table(UnitValue.CreatePercentArray(new float[] { 3.7f, 1.15f, 1.55f, 1f, 1.6f })).UseAllAvailableWidth();
        foreach (string value in new[] { "Item", "Quantity", "Net price", "VAT", "Net amount" })
            lines.AddHeaderCell(new Cell().SetBackgroundColor(Navy).SetPadding(7).Add(new Paragraph(value).SetFont(bold).SetFontColor(ColorConstants.WHITE)));
        for (int index = 0; index < invoice.Lines.Count; index++) {
            InvoiceLine line = invoice.Lines[index];
            Color fill = index % 2 == 1 ? Surface : ColorConstants.WHITE;
            AddBody(lines, $"{line.Id}. {line.Name}\n{line.Description}", fill, TextAlignment.LEFT);
            AddBody(lines, InvoiceComparisonScenario.Number(line.Quantity) + " " + line.UnitCode, fill, TextAlignment.RIGHT);
            AddBody(lines, InvoiceComparisonScenario.Number(line.UnitPrice) + " / " + InvoiceComparisonScenario.Number(line.PriceBaseQuantity) + " " + line.UnitCode, fill, TextAlignment.RIGHT);
            AddBody(lines, line.Tax.Code + " " + InvoiceComparisonScenario.Number(line.Tax.Rate!.Value) + "%", fill, TextAlignment.RIGHT);
            AddBody(lines, InvoiceComparisonScenario.Money(scenario.Calculation.Lines[index].NetAmount, invoice.Currency), fill, TextAlignment.RIGHT);
        }
        document.Add(lines);

        var taxes = new Table(UnitValue.CreatePercentArray(new float[] { 1.2f, 1.5f, 1.4f, 2f })).UseAllAvailableWidth().SetMarginTop(12);
        foreach (string value in new[] { "VAT category", "Taxable amount", "VAT amount", "Exemption" })
            taxes.AddHeaderCell(new Cell().SetBackgroundColor(Navy).SetPadding(7).Add(new Paragraph(value).SetFont(bold).SetFontColor(ColorConstants.WHITE)));
        AddBody(taxes, "S 23%", ColorConstants.WHITE, TextAlignment.LEFT);
        AddBody(taxes, InvoiceComparisonScenario.Money(scenario.Calculation.TaxExclusiveTotal, invoice.Currency), ColorConstants.WHITE, TextAlignment.RIGHT);
        AddBody(taxes, InvoiceComparisonScenario.Money(scenario.Calculation.TaxTotal, invoice.Currency), ColorConstants.WHITE, TextAlignment.RIGHT);
        AddBody(taxes, string.Empty, ColorConstants.WHITE, TextAlignment.LEFT);
        document.Add(taxes);

        document.Add(new AreaBreak(AreaBreakType.NEXT_PAGE));
        var second = new Table(UnitValue.CreatePercentArray(new float[] { 1.25f, 1f })).UseAllAvailableWidth().SetHorizontalBorderSpacing(18);
        second.AddCell(new Cell().SetBackgroundColor(Surface).SetBorder(new iText.Layout.Borders.SolidBorder(Border, 0.6f)).SetPadding(10)
            .Add(new Paragraph("Payment reference").SetFont(bold))
            .Add(new Paragraph(invoice.Number).SetFontColor(new DeviceRgb(82, 99, 122)))
            .Add(new Paragraph(invoice.PaymentTerms).SetFontColor(new DeviceRgb(82, 99, 122))));
        var totals = new Table(UnitValue.CreatePercentArray(new float[] { 1, 1 })).UseAllAvailableWidth();
        AddTotal(totals, "Line net total", scenario.Calculation.LineNetTotal, invoice.Currency, bold, false);
        AddTotal(totals, "Total excluding VAT", scenario.Calculation.TaxExclusiveTotal, invoice.Currency, bold, false);
        AddTotal(totals, "VAT total", scenario.Calculation.TaxTotal, invoice.Currency, bold, false);
        AddTotal(totals, "Total including VAT", scenario.Calculation.TaxInclusiveTotal, invoice.Currency, bold, false);
        AddTotal(totals, "Prepaid", scenario.Calculation.PrepaidAmount, invoice.Currency, bold, false);
        AddTotal(totals, "Amount due", scenario.Calculation.PayableAmount, invoice.Currency, bold, true);
        second.AddCell(NoBorder(new Cell()).Add(totals));
        document.Add(second.SetMarginBottom(16));
        document.Add(new Paragraph("References").SetFont(bold).SetFontSize(15));
        var references = new Table(UnitValue.CreatePercentArray(new float[] { 1, 1 })).UseAllAvailableWidth();
        AddBody(references, "Purchase order", ColorConstants.WHITE, TextAlignment.LEFT);
        AddBody(references, invoice.PurchaseOrderReference ?? string.Empty, ColorConstants.WHITE, TextAlignment.LEFT);
        AddBody(references, "Document type", ColorConstants.WHITE, TextAlignment.LEFT);
        AddBody(references, invoice.TypeCode, ColorConstants.WHITE, TextAlignment.LEFT);
        document.Add(references);
        document.Add(new Paragraph("Payment").SetFont(bold).SetFontSize(15));
        var payment = new Table(UnitValue.CreatePercentArray(new float[] { 1, 1 })).UseAllAvailableWidth();
        AddBody(payment, "Payment means", ColorConstants.WHITE, TextAlignment.LEFT);
        AddBody(payment, $"{invoice.Payments[0].MeansCode} {invoice.Payments[0].MeansText}", ColorConstants.WHITE, TextAlignment.LEFT);
        AddBody(payment, "Reference", ColorConstants.WHITE, TextAlignment.LEFT);
        AddBody(payment, invoice.Payments[0].Reference ?? string.Empty, ColorConstants.WHITE, TextAlignment.LEFT);
        AddBody(payment, "IBAN", ColorConstants.WHITE, TextAlignment.LEFT);
        AddBody(payment, $"{invoice.Payments[0].Account!.Identifier}\n{invoice.Payments[0].Account!.Name}", ColorConstants.WHITE, TextAlignment.LEFT);
        document.Add(payment);
        document.Add(new Paragraph(invoice.PaymentTerms));
        document.Add(new Paragraph("Notes").SetFont(bold).SetFontSize(15));
        document.Add(new Paragraph(invoice.Notes[0].Text));
        document.Add(new Paragraph("Approvals").SetFont(bold).SetFontSize(15));
        var approvals = new Table(UnitValue.CreatePercentArray(new float[] { 1, 1 })).UseAllAvailableWidth().SetHorizontalBorderSpacing(12);
        approvals.AddCell(ApprovalCell("PREPARED BY", "Marta Nowak", "Finance · 17/09/2026", bold));
        approvals.AddCell(ApprovalCell("APPROVED BY", "Daniel Reed", "Delivery lead · 17/09/2026", bold));
        document.Add(approvals);
        document.Close();
        return output.ToArray();
    }

    private static Cell NoBorder(Cell cell) => cell.SetBorder(null);

    private static Cell PartyCell(string label, InvoiceParty party, PdfFont bold) => new Cell()
        .SetBackgroundColor(Surface).SetBorder(new iText.Layout.Borders.SolidBorder(Border, 0.6f)).SetPadding(10)
        .Add(new Paragraph(label).SetFont(bold).SetFontColor(Accent))
        .Add(new Paragraph(party.Name).SetFontSize(11))
        .Add(new Paragraph(InvoiceComparisonScenario.PartyDetails(party)));

    private static void AddBody(Table table, string value, Color fill, TextAlignment alignment) => table.AddCell(
        new Cell().SetBackgroundColor(fill).SetBorderBottom(new iText.Layout.Borders.SolidBorder(Border, 0.5f)).SetPadding(7)
            .SetTextAlignment(alignment).Add(new Paragraph(value).SetMargin(0)));

    private static void AddTotal(Table table, string label, decimal value, string currency, PdfFont bold, bool accent) {
        Color fill = accent ? Accent : ColorConstants.WHITE;
        Color text = accent ? ColorConstants.WHITE : Navy;
        var labelParagraph = new Paragraph(label).SetFontColor(text);
        var valueParagraph = new Paragraph(InvoiceComparisonScenario.Money(value, currency)).SetFontColor(text);
        if (accent) {
            labelParagraph.SetFont(bold);
            valueParagraph.SetFont(bold);
        }

        table.AddCell(new Cell().SetBackgroundColor(fill).SetPadding(7).Add(labelParagraph));
        table.AddCell(new Cell().SetBackgroundColor(fill).SetPadding(7).SetTextAlignment(TextAlignment.RIGHT)
            .Add(valueParagraph));
    }

    private static Cell ApprovalCell(string label, string name, string detail, PdfFont bold) => new Cell()
        .SetBackgroundColor(Surface).SetBorder(new iText.Layout.Borders.SolidBorder(Border, 0.6f)).SetPadding(10)
        .Add(new Paragraph(label).SetFontColor(new DeviceRgb(82, 99, 122)))
        .Add(new Paragraph(name).SetFont(bold))
        .Add(new Paragraph(detail).SetFontColor(new DeviceRgb(82, 99, 122)));
}
