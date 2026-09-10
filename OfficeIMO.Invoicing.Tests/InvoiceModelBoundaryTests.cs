using System.Globalization;
using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Tests;

public class InvoiceModelBoundaryTests {
    [Theory]
    [InlineData("https://example.test/invoice%20detail")]
    [InlineData("ftp://example.test/invoice.pdf")]
    [InlineData("urn:uuid:00112233-4455-6677-8899-aabbccddeeff")]
    public void AbsoluteSupportingDocumentUrisSurviveBothSyntaxes(string uri) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.SupportingDocuments.Add(new InvoiceSupportingDocument { Reference = "support", ExternalUri = uri });
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl }) {
            byte[] xml = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax));
            InvoiceReadResult read = InvoiceParser.Read(xml);
            Assert.True(read.HasCompleteMapping);
            Assert.Equal(uri, Assert.Single(read.Invoice.SupportingDocuments).ExternalUri);
            Assert.Equal(xml, read.Write());
            InvoiceConversionResult converted = InvoiceConverter.Convert(xml, new InvoiceXmlOptions(syntax == InvoiceSyntax.Cii ? InvoiceSyntax.Ubl : InvoiceSyntax.Cii));
            Assert.True(converted.Succeeded, string.Join("; ", converted.Diagnostics.Select(d => d.Message)));
            Assert.Equal(uri, Assert.Single(InvoiceParser.Read(converted.Xml!).Invoice.SupportingDocuments).ExternalUri);
        }
    }

    [Theory]
    [InlineData("relative/document.pdf")]
    [InlineData("https://example.test/unescaped space")]
    [InlineData(" ")]
    public void SupportingDocumentUrisRequireAWellFormedAbsoluteValue(string uri) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.SupportingDocuments.Add(new InvoiceSupportingDocument { Reference = "support", ExternalUri = uri });
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics, d => d.Code == "INV-ATTACHMENT-URI");
    }

    [Theory]
    [InlineData(null, "Blue", "Name")]
    [InlineData("", "Blue", "Name")]
    [InlineData(" ", "Blue", "Name")]
    [InlineData("Color", null, "Value")]
    [InlineData("Color", "", "Value")]
    [InlineData("Color", " ", "Value")]
    public void ItemAttributesRequireNameAndValue(string? name, string? value, string missing) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Lines[0].Attributes.Add(new InvoiceItemAttribute { Name = name!, Value = value! });
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics, d => d.Location == "Lines[0].Attributes." + missing);
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl })
            Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InvalidExtremeGrossPricesReturnDiagnosticsForModelsAndImportedXml(bool negativeGross) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Lines[0].GrossPrice = 110m;
        invoice.Lines[0].PriceDiscount = 10m;
        decimal gross = negativeGross ? decimal.MinValue : decimal.MaxValue;
        decimal discount = negativeGross ? decimal.MaxValue : -1m;
        XNamespace ram = "urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100";
        XDocument source = XDocument.Parse(Encoding.UTF8.GetString(InvoiceSerializer.Write(invoice)));
        XElement price = source.Descendants(ram + "GrossPriceProductTradePrice").Single();
        price.Element(ram + "ChargeAmount")!.Value = gross.ToString(CultureInfo.InvariantCulture);
        price.Descendants(ram + "ActualAmount").Single().Value = discount.ToString(CultureInfo.InvariantCulture);
        invoice.Lines[0].GrossPrice = gross;
        invoice.Lines[0].PriceDiscount = discount;
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics, d => d.Code == "INV-PRICE");
        InvoiceConversionResult result = InvoiceConverter.Convert(Encoding.UTF8.GetBytes(source.ToString()), new InvoiceXmlOptions(InvoiceSyntax.Ubl));
        Assert.False(result.Succeeded);
        Assert.Contains(result.Diagnostics, d => d.Code == "INV-PRICE");
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" ")]
    public void SalesOrderRequiresANonBlankPurchaseOrderForUbl(string? purchaseOrder) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.SalesOrderReference = "sales-1";
        invoice.PurchaseOrderReference = purchaseOrder;
        Assert.Contains(InvoiceSerializer.InspectTarget(invoice, new InvoiceXmlOptions(InvoiceSyntax.Ubl)), d => d.Location == "SalesOrderReference");
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(InvoiceSyntax.Ubl)));
    }
}
