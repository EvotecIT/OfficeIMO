using System.Globalization;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Tests;

public class InvoiceSerializationTests {
    [Theory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public void OutputIsDeterministicAndCultureIndependent(InvoiceSyntax syntax) {
        Invoice invoice = InvoiceCalculationTests.Example();
        invoice.Lines[0].UnitPrice = 12.345m;
        invoice.Lines[0].Name = "Parts & services <test> – Żółć";
        var options = new InvoiceXmlOptions(syntax);
        CultureInfo saved = CultureInfo.CurrentCulture;
        try {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("pl-PL");
            byte[] first = InvoiceSerializer.Write(invoice, options);
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
            Assert.Equal(first, InvoiceSerializer.Write(invoice, options));
            Assert.False(first.Take(3).SequenceEqual(new byte[] { 0xef, 0xbb, 0xbf }));
            Assert.DoesNotContain("\r", Encoding.UTF8.GetString(first));
            Assert.Contains("12.345", Encoding.UTF8.GetString(first));
            XDocument parsed = XDocument.Parse(Encoding.UTF8.GetString(first));
            Assert.Contains(parsed.Descendants(), e => e.Value == invoice.Lines[0].Name);
            Assert.Equal(InvoiceProfile.En16931, InvoiceProfileDeclaration.Read(first).Profile);
        } finally { CultureInfo.CurrentCulture = saved; }
    }

    [Fact]
    public void ConversionTargetDoesNotDiscardUnsupportedCreditNoteFields() {
        Invoice invoice = InvoiceCalculationTests.Example();
        invoice.TypeCode = "381";
        invoice.ProjectReference = "project";
        var options = new InvoiceXmlOptions(InvoiceSyntax.Ubl);
        Assert.Contains(InvoiceSerializer.InspectTarget(invoice, options), d => d.Location == "ProjectReference");
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, options));
    }
}
