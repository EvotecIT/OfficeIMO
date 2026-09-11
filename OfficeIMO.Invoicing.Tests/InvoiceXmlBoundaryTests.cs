using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Tests;

public class InvoiceXmlBoundaryTests {
    [Theory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public void SharedPaymentTextIsEmittedOnceRegardlessOfAccountCount(InvoiceSyntax syntax) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payment!.MeansText = new string('D', 65536);
        invoice.Payment.Reference = new string('R', 65536);
        for (int index = 1; index < 40; index++) invoice.Payment.Accounts.Add(new InvoiceBankAccount { Identifier = "DE79000000001234567890" });
        byte[] xml = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax));
        Assert.InRange(xml.Length, 1, 256 * 1024);
        InvoiceReadResult read = InvoiceParser.Read(xml);
        Assert.True(read.HasCompleteMapping);
        Assert.Equal(40, read.Invoice.Payment!.Accounts.Count);
        Assert.Equal(invoice.Payment.Reference, read.Invoice.Payment.Reference);
        Assert.Equal(invoice.Payment.MeansText, read.Invoice.Payment.MeansText);
    }

    [Theory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public void TextOnlyPaymentContainersCannotBeSilentlyDiscarded(InvoiceSyntax syntax) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.PaymentTerms = "Original terms";
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax))));
        document.Descendants().Single(e => e.Name.LocalName == (syntax == InvoiceSyntax.Cii ? "SpecifiedTradePaymentTerms" : "PaymentTerms"))
            .Value = "Important payment instructions";
        byte[] source = Encoding.UTF8.GetBytes(document.ToString());
        InvoiceReadResult read = InvoiceParser.Read(source);
        Assert.False(read.HasCompleteMapping);
        Assert.Contains(read.UnmappedData, d => d.Message.Contains("Text in an invoice container"));
        Assert.Throws<InvalidDataException>(() => read.Write());
        Assert.False(InvoiceConverter.Convert(source, new InvoiceXmlOptions(syntax == InvoiceSyntax.Cii ? InvoiceSyntax.Ubl : InvoiceSyntax.Cii)).Succeeded);
    }

    [Theory]
    [InlineData(1, false)]
    [InlineData(1, true)]
    [InlineData(0xd800, false)]
    [InlineData(0xd800, true)]
    public void IllegalXmlCharactersAreModelDiagnostics(int character, bool optional) {
        Invoice invoice = InvoiceFixture.Create();
        string invalid = "Text" + (char)character;
        if (optional) invoice.Lines[0].Description = invalid;
        else invoice.Number = invalid;
        InvoiceModelValidationResult validation = InvoiceModelValidator.Validate(invoice);
        Assert.False(validation.IsValid);
        Assert.Contains(validation.Diagnostics, d => d.Message.Contains("XML cannot represent"));
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl })
            Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax)));
    }

    [Fact]
    public void SupplementaryUnicodeCharactersRemainSupported() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Lines[0].Description = "Service \U0001F600";
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl })
            Assert.Equal(invoice.Lines[0].Description, InvoiceParser.Read(InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax))).Invoice.Lines[0].Description);
    }

    [Fact]
    public void XmlOutputBudgetStopsWritesBeforeBufferGrowth() {
        using var output = new InvoiceXmlOutputStream();
        byte[] block = new byte[65536];
        for (int index = 0; index < InvoiceProfileDeclaration.MaximumXmlBytes / block.Length; index++) output.Write(block, 0, block.Length);
        Assert.Throws<InvalidDataException>(() => output.Write(block, 0, block.Length));
        Assert.Throws<InvalidDataException>(() => output.WriteByte(1));
        Assert.Equal(InvoiceProfileDeclaration.MaximumXmlBytes, output.Length);
    }
}
