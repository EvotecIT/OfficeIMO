using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Tests;

public class InvoiceTaxBoundaryTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReservedSepaIdentifiersCannotChangeRolesInUbl(bool seller) {
        Invoice invoice = InvoiceFixture.Create();
        (seller ? invoice.Seller : invoice.Buyer).Identifiers.Add(new InvoiceIdentifier("party-id", "SEPA"));
        string path = seller ? "Seller.Identifiers" : "Buyer.Identifiers";
        Assert.Contains(InvoiceSerializer.InspectTarget(invoice, new InvoiceXmlOptions(InvoiceSyntax.Ubl)), d => d.Location == path);
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(InvoiceSyntax.Ubl)));
        byte[] cii = InvoiceSerializer.Write(invoice);
        InvoiceReadResult read = InvoiceParser.Read(cii);
        Assert.True(read.HasCompleteMapping);
        Assert.Equal("party-id", Assert.Single((seller ? read.Invoice.Seller : read.Invoice.Buyer).Identifiers).Value);
        InvoiceConversionResult converted = InvoiceConverter.Convert(cii, new InvoiceXmlOptions(InvoiceSyntax.Ubl));
        Assert.False(converted.Succeeded);
        Assert.Contains(converted.Diagnostics, d => d.Location == path);
    }

    [Theory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public void ImportedRepresentativeWithoutVatIdentifierCannotBeRewrittenOrConverted(InvoiceSyntax syntax) {
        Invoice invoice = InvoiceFixture.Rich();
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax))));
        XElement representative = document.Descendants().Single(e => e.Name.LocalName ==
            (syntax == InvoiceSyntax.Cii ? "SellerTaxRepresentativeTradeParty" : "TaxRepresentativeParty"));
        representative.Elements().Single(e => e.Name.LocalName ==
            (syntax == InvoiceSyntax.Cii ? "SpecifiedTaxRegistration" : "PartyTaxScheme")).Remove();
        byte[] source = Encoding.UTF8.GetBytes(document.ToString());
        InvoiceReadResult read = InvoiceParser.Read(source);
        Assert.Throws<InvalidDataException>(() => read.Write(allowUnmappedDataLoss: true));
        InvoiceConversionResult converted = InvoiceConverter.Convert(source, new InvoiceXmlOptions(syntax == InvoiceSyntax.Cii ? InvoiceSyntax.Ubl : InvoiceSyntax.Cii));
        Assert.False(converted.Succeeded);
        Assert.Contains(converted.Diagnostics, d => d.Location == "TaxRepresentative.VatIdentifier");
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" ")]
    public void TaxRepresentativeRequiresVatIdentifier(string? identifier) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.TaxRepresentative = new InvoiceParty {
            Name = "Representative", VatIdentifier = identifier,
            Address = new InvoiceAddress { CountryCode = "DE" }
        };
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics,
            d => d.Location == "TaxRepresentative.VatIdentifier" && d.Code == "INV-REQUIRED");
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl })
            Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax)));
    }

    [Theory]
    [InlineData("USD")]
    [InlineData("GBP")]
    [InlineData(null)]
    public void NonInvoiceCurrencyBreakdownsRequireExplicitLossWithoutPollutingTheModel(string? currency) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.TaxCurrency = "USD";
        invoice.TaxAmountInAccountingCurrency = 25m;
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(InvoiceSyntax.Ubl))));
        XNamespace cac = "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2";
        XNamespace cbc = "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2";
        XElement primary = document.Root!.Elements(cac + "TaxTotal").First();
        XElement extra = new XElement(primary);
        foreach (XAttribute attribute in extra.Descendants().Attributes("currencyID")) attribute.Value = currency ?? "USD";
        if (currency == null) extra.Element(cbc + "TaxAmount")!.Attribute("currencyID")!.Remove();
        if (currency == "USD") document.Root.Elements(cac + "TaxTotal").Last().ReplaceWith(extra);
        else primary.AddAfterSelf(extra);
        byte[] source = Encoding.UTF8.GetBytes(document.ToString());
        InvoiceReadResult read = InvoiceParser.Read(source);
        Assert.False(read.HasCompleteMapping);
        Assert.Contains(read.UnmappedData, d => d.Message.Contains("breakdowns cannot be mapped"));
        Assert.Equal(primary.Elements(cac + "TaxSubtotal").Count(), read.Invoice.DeclaredTaxes.Count);
        Assert.True(InvoiceModelValidator.Validate(read.Invoice).IsValid);
        Assert.Throws<InvalidDataException>(() => read.Write());
        Assert.False(InvoiceConverter.Convert(source, new InvoiceXmlOptions(InvoiceSyntax.Cii)).Succeeded);
        byte[] rewritten = read.Write(allowUnmappedDataLoss: true);
        InvoiceReadResult result = InvoiceParser.Read(rewritten);
        Assert.True(result.HasCompleteMapping);
        Assert.Equal(read.Invoice.DeclaredTaxes.Count, result.Invoice.DeclaredTaxes.Count);
        Assert.Equal(read.Invoice.TaxAmountInAccountingCurrency, result.Invoice.TaxAmountInAccountingCurrency);
        InvoiceReadResult cii = InvoiceParser.Read(read.Write(new InvoiceXmlOptions(InvoiceSyntax.Cii), allowUnmappedDataLoss: true));
        Assert.True(cii.HasCompleteMapping);
        Assert.Equal(result.Invoice.TaxAmountInAccountingCurrency, cii.Invoice.TaxAmountInAccountingCurrency);
    }
}
