using OfficeIMO.Invoicing;
using OfficeIMO.Invoicing.Tests;
using OfficeIMO.Pdf;
using OfficeIMO.Reader;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfInvoiceDocumentTests {
    [Theory]
    [InlineData(false, false, true)]
    [InlineData(false, false, false)]
    [InlineData(true, false, true)]
    [InlineData(false, true, true)]
    public void VisiblePdfAndAttachmentUseTheCapturedInvoice(bool credit, bool longInvoice, bool rich) {
        Invoice invoice = rich ? InvoiceFixture.Rich(credit) : InvoiceFixture.Create();
        if (longInvoice) {
            for (int index = 3; index <= 35; index++) invoice.Lines.Add(new InvoiceLine {
                Id = index.ToString(System.Globalization.CultureInfo.InvariantCulture), Name = "Service line " + index,
                Quantity = index, UnitPrice = 5.25m, Tax = new InvoiceTaxCategory { Code = "S", Rate = 19m }
            });
        }
        PdfInvoiceDocument snapshot = PdfInvoiceDocument.Create(invoice);
        byte[] xml = snapshot.ToXmlBytes();
        decimal due = InvoiceCalculator.Calculate(invoice).PayableAmount;
        invoice.Number = "MUTATED"; invoice.Lines[0].UnitPrice = 999m;
        snapshot.ToInvoice().Number = "SECOND-MUTATION";
        byte[] returned = snapshot.ToXmlBytes(); returned[0] ^= 1;
        byte[] pdf = snapshot.ToPdfBytes(Options());
        Assert.Equal(xml, snapshot.ToXmlBytes());
        var attachment = Assert.Single(PdfDocument.Load(pdf).Attachments.Extract());
        Assert.Equal("factur-x.xml", attachment.FileName);
        Assert.Equal(xml, attachment.Bytes);
        string text = PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains("INV-2026-001", text, StringComparison.Ordinal);
        Assert.DoesNotContain("MUTATED", text, StringComparison.Ordinal);
        Assert.Contains(due.ToString("0.00", System.Globalization.CultureInfo.InvariantCulture) + " EUR", text, StringComparison.Ordinal);
        Assert.Contains("Consulting", text, StringComparison.Ordinal);
        Assert.Contains("DE79000000001234567890", text, StringComparison.Ordinal);
        if (longInvoice) Assert.Contains("Service line 35", text, StringComparison.Ordinal);
        string name = !rich ? "simple-invoice" : longInvoice ? "multipage" : credit ? "credit-note" : "invoice";
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_INVOICE_PDF_EVIDENCE") ?? Environment.GetEnvironmentVariable("OFFICEIMO_PDF_COMPLIANCE_PROOF_OUTPUT");
        if (!string.IsNullOrWhiteSpace(output)) {
            Directory.CreateDirectory(output!);
            File.WriteAllBytes(Path.Combine(output!, name + ".pdf"), pdf);
            File.WriteAllBytes(Path.Combine(output!, name + ".xml"), xml);
        }
        foreach (PdfExternalValidator validator in new[] { PdfExternalValidator.VeraPdf(), PdfExternalValidator.Mustang() }) {
            if (!validator.IsAvailable) { PdfExternalValidator.SkipUnlessRequired(validator); continue; }
            PdfExternalProcessResult result = validator.Run(pdf, name + ".pdf");
            if (!string.IsNullOrWhiteSpace(output)) File.WriteAllText(Path.Combine(output!, validator.Name + "-" + name + ".txt"), result.GetDiagnosticText());
            Assert.True(result.ExitCode == 0, result.GetDiagnosticText());
        }
    }
    private static PdfOptions Options() {
        string fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont()!;
        Assert.NotNull(fontPath);
        byte[] font = File.ReadAllBytes(fontPath);
        return new PdfOptions { IncludeStandardFontToUnicodeMaps = true }
            .EmbedStandardFont(PdfStandardFont.Helvetica, font, "OfficeIMO Source Serif")
            .EmbedStandardFont(PdfStandardFont.HelveticaBold, font, "OfficeIMO Source Serif");
    }
}
