using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Tests;

public class InvoicePreservationTests {
    [Theory]
    [InlineData("100.00000000000000000000000000001")]
    [InlineData("0.00000000000000000000000000001")]
    [InlineData("7922816251426433759354395033.51")]
    public void InexactDecimalInputIsRejected(string value) {
        XDocument xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(InvoiceSerializer.Write(InvoiceFixture.Create())));
        XNamespace ram = "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100";
        xml.Descendants(ram + "NetPriceProductTradePrice").Single().Element(ram + "ChargeAmount")!.Value = value;
        Assert.Throws<InvalidDataException>(() => InvoiceParser.Read(System.Text.Encoding.UTF8.GetBytes(xml.ToString())));
    }

    [Theory]
    [InlineData("+000100.000000000000000000000000000000000000000", 100)]
    [InlineData("-0.000000000000000000000000000000000000000000", 0)]
    public void ExactDecimalLexicalVariantsRemainReadable(string value, int expected) {
        XDocument xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(InvoiceSerializer.Write(InvoiceFixture.Create())));
        XNamespace ram = "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100";
        xml.Descendants(ram + "NetPriceProductTradePrice").Single().Element(ram + "ChargeAmount")!.Value = value;
        Assert.Equal(expected, InvoiceParser.Read(System.Text.Encoding.UTF8.GetBytes(xml.ToString())).Invoice.Lines[0].UnitPrice);
    }

    [Fact]
    public void LiteralClassifiedNotePrefixCannotChangeMeaningDuringConversion() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Notes.Add(new InvoiceNote("#ADU#This is literal text"));
        var result = InvoiceConverter.Convert(InvoiceSerializer.Write(invoice), new InvoiceXmlOptions(InvoiceSyntax.Ubl));
        Assert.False(result.Succeeded);
        Assert.Contains(result.Diagnostics, d => d.Location == "Notes" && d.Code == "INV-TARGET-UNSUPPORTED");
    }

    [Theory]
    [InlineData("ab")]
    [InlineData("ABCD")]
    [InlineData("A#U")]
    public void InvalidSubjectCodeCannotProduceAmbiguousUbl(string subject) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Notes.Add(new InvoiceNote("A note", subject));
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(InvoiceSyntax.Ubl)));
    }
}
