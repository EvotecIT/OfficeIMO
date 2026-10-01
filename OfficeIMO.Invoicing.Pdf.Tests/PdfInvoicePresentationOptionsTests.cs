using OfficeIMO.Invoicing.Tests;
using OfficeIMO.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Invoicing.Pdf.Tests;

public sealed class PdfInvoicePresentationOptionsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OrderedColumnsAndLocalizedCodesRetainBusinessDetailsInBothLayouts(bool modern) {
        Invoice invoice = InvoiceFixture.Rich();
        invoice.Lines[0].Description = "DescriptionMarker";
        var layout = InvoicePdfLayoutOptions.ForCultures("pl-PL", "en-GB");
        layout.Theme = modern ? new InvoicePdfTheme() : null;
        layout.UnitCodeDisplay = InvoicePdfCodeDisplay.CodeAndDescription;
        layout.PaymentCodeDisplay = InvoicePdfCodeDisplay.Description;
        layout.CompactDetails = true;
        layout.IncludePageIdentity = true;
        layout.LineColumns.Clear();
        foreach (var column in new[] { InvoicePdfLineColumn.LineIdentifier, InvoicePdfLineColumn.Item, InvoicePdfLineColumn.Description,
            InvoicePdfLineColumn.Quantity, InvoicePdfLineColumn.Unit, InvoicePdfLineColumn.NetPrice, InvoicePdfLineColumn.Vat, InvoicePdfLineColumn.NetAmount }) layout.LineColumns.Add(column);
        var snapshot = PdfInvoiceDocument.Create(invoice, Contract(), layout);
        layout.LineColumns.Clear(); // A deferred render uses its captured column settings.
        var options = Options(); options.PageWidth = 842; options.PageHeight = 595;
        byte[] pdf = snapshot.ToPdfBytes(options);
        WriteEvidence(modern ? "presentation-modern" : "presentation-classic", pdf);
        string text = PdfReadDocument.Open(pdf).ExtractText();
        foreach (string label in new[] { "Numer", "pozycji", "Line ID", "Opis", "Description", "C62", "sztuka", "piece" }) Assert.Contains(label, text);
        Assert.Contains("Przelew SEPA / SEPA credit transfer", text);
        foreach (string value in new[] { "DescriptionMarker", "project-cost", "service-1", "buyer-service-1", "0721-880X", "Line discount", "Pay within thirty days.", "prior-invoice" }) Assert.Contains(value, text);
        Assert.Equal(1, text.Split(new[] { "DescriptionMarker" }, StringSplitOptions.None).Length - 1);
        Assert.Equal(snapshot.ToXmlBytes(), Assert.Single(PdfDocument.Load(pdf).Attachments.Extract()).Bytes);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InvoiceIdentityAndPageCountsAppearOnEveryPageWithoutExpandingNumberTokens(bool modern) {
        Invoice invoice = InvoiceFixture.Create(); invoice.Number = "INV-{page}-{pages}-LITERAL";
        for (int index = 2; index <= 65; index++) invoice.Lines.Add(new InvoiceLine {
            Id = index.ToString(System.Globalization.CultureInfo.InvariantCulture), Name = "Consulting line " + index,
            Quantity = 1, UnitPrice = 100, Tax = new InvoiceTaxCategory { Code = "S", Rate = 19 }
        });
        var layout = InvoicePdfLayoutOptions.ForCultures("pl-PL");
        layout.Theme = modern ? new InvoicePdfTheme() : null;
        layout.IncludePageIdentity = true;
        var options = Options(); options.ShowHeader = true; options.HeaderFormat = "RUNNING-HEADER";
        options.DifferentFirstPageHeaderFooter = true; options.FirstPageHeaderFormat = "FIRST-HEADER";
        options.DifferentOddAndEvenPagesHeaderFooter = true; options.EvenPageHeaderFormat = "EVEN-HEADER";
        options.FirstPageFooterFormat = "OLD-FIRST-FOOTER"; options.EvenPageFooterFormat = "OLD-EVEN-FOOTER";
        byte[] pdf = PdfInvoiceDocument.Create(invoice, Contract(), layout).ToPdfBytes(options);
        PdfReadDocument document = PdfReadDocument.Open(pdf);
        Assert.True(document.Pages.Count > 2);
        for (int index = 0; index < document.Pages.Count; index++) {
            string text = document.Pages[index].ExtractText();
            Assert.Contains(invoice.Number, text);
            Assert.Contains("Strona " + (index + 1) + " / " + document.Pages.Count, text);
            Assert.Contains(index == 0 ? "FIRST-HEADER" : index % 2 == 1 ? "EVEN-HEADER" : "RUNNING-HEADER", text);
            Assert.DoesNotContain("OLD-FIRST-FOOTER", text);
            Assert.DoesNotContain("OLD-EVEN-FOOTER", text);
        }
        Assert.Equal("OLD-FIRST-FOOTER", options.FirstPageFooterFormat);
        WriteEvidence(modern ? "continuation-modern" : "continuation-classic", pdf);
    }

    [Fact]
    public void ColumnSelectionsRejectMissingAssociationDuplicateAndUndefinedValues() {
        var layout = new InvoicePdfLayoutOptions();
        layout.LineColumns.Remove(InvoicePdfLineColumn.Item);
        Assert.Throws<InvalidOperationException>(() => layout.Clone());
        layout.LineColumns.Insert(0, InvoicePdfLineColumn.Item);
        layout.LineColumns.Add(InvoicePdfLineColumn.Item);
        Assert.Throws<InvalidOperationException>(() => layout.Clone());
        layout.LineColumns.RemoveAt(layout.LineColumns.Count - 1);
        layout.LineColumns.Add((InvoicePdfLineColumn)999);
        Assert.Throws<InvalidOperationException>(() => layout.Clone());
    }

    [Theory]
    [InlineData(InvoicePdfLineColumn.GrossPrice, "110 / 1 C62")]
    [InlineData(InvoicePdfLineColumn.PriceDiscount, "10 / 1 C62")]
    public void OptionalBusinessColumnsRetainTheirOwnValues(InvoicePdfLineColumn priceColumn, string price) {
        Invoice invoice = InvoiceFixture.Rich();
        invoice.Lines[0].SellerItemIdentifier = "SELLERKEY"; invoice.Lines[0].BuyerItemIdentifier = "BUYERKEY";
        var layout = new InvoicePdfLayoutOptions(); layout.LineColumns.Clear();
        foreach (var column in new[] { InvoicePdfLineColumn.Item, InvoicePdfLineColumn.Period, InvoicePdfLineColumn.SellerItem,
            InvoicePdfLineColumn.BuyerItem, InvoicePdfLineColumn.StandardItem, InvoicePdfLineColumn.AccountingReference, priceColumn,
            InvoicePdfLineColumn.NetAmount }) layout.LineColumns.Add(column);
        var options = Options(); options.PageWidth = 1191; options.PageHeight = 842;
        string text = PdfReadDocument.Open(PdfInvoiceDocument.Create(invoice, Contract(), layout).ToPdfBytes(options)).ExtractText();
        foreach (string value in new[] { "SELLERKEY", "BUYERKEY", "1234567890128", "project-cost", "2026-08-31", "2026-09-10", price }) Assert.Contains(value, text);
        Assert.Equal(1, text.Split(new[] { "SELLERKEY" }, StringSplitOptions.None).Length - 1);
        Assert.Equal(1, text.Split(new[] { "BUYERKEY" }, StringSplitOptions.None).Length - 1);
    }

    [Fact]
    public void CustomCodeDescriptionsAreCapturedAndUnknownCodesRetainTheirIdentity() {
        var units = new Dictionary<string, string> { ["ZZ"] = "CustomUnit" };
        var pack = InvoicePdfLanguagePack.ForCulture("pl-PL").WithCodeDescriptions(units);
        units["ZZ"] = "Changed after capture";
        Assert.Equal("CustomUnit", pack.GetUnitDescription("ZZ"));
        Assert.Equal("sztuka", pack.GetUnitDescription("C62"));
        Assert.Null(pack.GetUnitDescription("UNKNOWN"));
        Invoice invoice = InvoiceFixture.Create(); invoice.Lines[0].UnitCode = "ZZ";
        invoice.Payments[0].MeansCode = "999"; invoice.Payments[0].MeansText = "AUTHORED-MEANS-MARKER";
        var layout = new InvoicePdfLayoutOptions { UnitCodeDisplay = InvoicePdfCodeDisplay.Description, PaymentCodeDisplay = InvoicePdfCodeDisplay.Description };
        layout.Languages.Clear(); layout.Languages.Add(pack);
        string text = PdfReadDocument.Open(PdfInvoiceDocument.Create(invoice, Contract(), layout).ToPdfBytes(Options())).ExtractText();
        Assert.Contains("CustomUnit", text);
        Assert.Contains("999 AUTHORED-MEANS-MARKER", text);
    }

    private static InvoiceXmlOptions Contract() => new(InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2, InvoiceSyntax.Cii, InvoiceProfile.En16931);
    private static PdfOptions Options() {
        byte[] font = File.ReadAllBytes(PdfComplianceTestFonts.FindBundledOpenTypeCffFont()!);
        return new PdfOptions().EmbedStandardFont(PdfStandardFont.Helvetica, font, "Source Serif").EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "Source Serif");
    }
    private static void WriteEvidence(string name, byte[] pdf) {
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_PDF_EVIDENCE");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output!); File.WriteAllBytes(Path.Combine(output!, name + ".pdf"), pdf);
    }
}
