using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Tests;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Invoicing.Pdf.Tests;

public sealed class PdfInvoiceDocumentTypeTests {
    [Theory]
    [InlineData("326", "Partial invoice", false)]
    [InlineData("384", "Corrected invoice", false)]
    [InlineData("386", "Prepayment invoice", false)]
    [InlineData("389", "Self-billed invoice", false)]
    [InlineData("326", "Partial invoice", true)]
    [InlineData("384", "Corrected invoice", true)]
    [InlineData("386", "Prepayment invoice", true)]
    [InlineData("389", "Self-billed invoice", true)]
    public void DocumentTypeHeadingMatchesCapturedXmlInBothLayouts(string code, string heading, bool modern) {
        var invoice = InvoiceFixture.Create(); invoice.TypeCode = code;
        if (code == InvoiceDocumentTypes.CorrectedInvoice) invoice.PrecedingInvoices.Add(new("ORIGINAL-001", invoice.IssueDate.AddDays(-1)));
        var layout = new InvoicePdfLayoutOptions { Theme = modern ? new InvoicePdfTheme() : null };
        var contract = new InvoiceXmlOptions(InvoiceSpecificationRelease.FacturX_1_09_2_Zugferd_2_5_2, InvoiceSyntax.Cii, InvoiceProfile.En16931);
        var snapshot = PdfInvoiceDocument.Create(invoice, contract, layout);
        byte[] font = File.ReadAllBytes(OfficeIMO.Tests.Pdf.PdfComplianceTestFonts.FindBundledOpenTypeCffFont()!);
        var options = new PdfOptions().EmbedStandardFont(PdfStandardFont.Helvetica, font, "Source Serif")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "Source Serif");
        byte[] pdf = snapshot.ToPdfBytes(options);
        string text = PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains(heading, text);
        Assert.Contains(invoice.Number, text);
        Assert.Equal(code, InvoiceParser.Read(Assert.Single(PdfDocument.Load(pdf).Attachments.Extract()).Bytes).Invoice.TypeCode);
        Assert.Equal(119m, snapshot.PayableAmount);
    }
}
