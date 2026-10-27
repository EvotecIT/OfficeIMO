using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Tests;
using OfficeIMO.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Invoicing.Pdf.Tests;

public class PdfInvoiceReadabilityTests {
    private static InvoiceXmlOptions Contract() => new InvoiceXmlOptions(
        InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2, InvoiceSyntax.Cii, InvoiceProfile.En16931,
        InvoiceProjectionPolicy.RejectDataLoss);

    private static PdfCreateOptions Options() {
        string fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont()!;
        byte[] font = File.ReadAllBytes(fontPath);
        return new PdfCreateOptions { IncludeStandardFontToUnicodeMaps = true }
            .EmbedStandardFont(PdfStandardFont.Helvetica, font, "OfficeIMO Source Serif")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "OfficeIMO Source Serif");
    }

    private static string Text(byte[] pdf) => string.Join(" ",
        PdfDocument.Load(new MemoryStream(pdf)).ExtractUnwrappedText().Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    [Fact]
    public void DefaultPresentationShowsReadableLabelsInsteadOfCodes() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Lines[0].UnitCode = "HUR";
        invoice.Lines[0].Quantity = 1250m;
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice, Contract(), InvoicePdfLayoutOptions.ForCultures("en-GB"));

        string text = Text(snapshot.ToPdfBytes(Options()));

        Assert.Contains("1,250 hour", text, StringComparison.Ordinal);
        Assert.DoesNotContain(" HUR", text, StringComparison.Ordinal);
        Assert.Contains("Standard rate 19%", text, StringComparison.Ordinal);
        Assert.Contains("Commercial invoice", text, StringComparison.Ordinal);
        Assert.Contains("SEPA credit transfer", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Document type 380", text, StringComparison.Ordinal);
        Assert.Contains("125,000.00 EUR", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SellerEmailRepeatingTheElectronicAddressAppearsOnce() {
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(InvoiceFixture.Create(), Contract());

        string text = Text(snapshot.ToPdfBytes(Options()));

        int first = text.IndexOf("seller@example.test", StringComparison.Ordinal);
        Assert.True(first >= 0);
        Assert.Equal(-1, text.IndexOf("seller@example.test", first + 1, StringComparison.Ordinal));
    }

    [Fact]
    public void InvoicesCarryPageNumbersLanguageAndDisplayedTitle() {
        Invoice invoice = InvoiceFixture.Create();
        for (int index = 2; index <= 40; index++) invoice.Lines.Add(new InvoiceLine {
            Id = index.ToString(System.Globalization.CultureInfo.InvariantCulture), Name = "Service line " + index,
            Quantity = 1m, UnitPrice = 5m, Tax = new InvoiceTaxCategory { Code = "S", Rate = 19m }
        });
        byte[] pdf = PdfInvoiceDocument.Create(invoice, Contract(), InvoicePdfLayoutOptions.ForCultures("en-GB")).ToPdfBytes(Options());

        PdfDocument read = PdfDocument.Load(new MemoryStream(pdf));
        int pages = read.Pages.Count;
        Assert.True(pages > 1);
        string text = read.ExtractText();
        for (int page = 1; page <= pages; page++) Assert.Contains("Page " + page + " of " + pages, text, StringComparison.Ordinal);
        PdfDocumentInfo info = PdfDocument.Load(new MemoryStream(pdf)).Inspect();
        Assert.Equal("en-GB", info.CatalogLanguage);
        Assert.True(info.ViewerPreferences?.GetBoolean("DisplayDocTitle"));
    }
}
