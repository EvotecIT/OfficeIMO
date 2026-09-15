using System.Text;
using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Tests;

public class InvoiceTaxBoundaryTests {
    [Fact]
    public void ConflictingExemptionDescriptionsRemainInTheModelAndReportExactTargetValues() {
        Invoice invoice = InvoiceFixture.WithTaxCategory("E");
        InvoiceCalculator.UpdateDeclaredAmounts(invoice);
        invoice.Lines[0].Tax.ExemptionReason = "Line exemption";
        invoice.DeclaredTaxes[0].Category.ExemptionReason = "Header exemption";
        Assert.True(InvoiceModelValidator.Validate(invoice).IsValid);
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl }) {
            IReadOnlyList<InvoiceDiagnostic> diagnostics = InvoiceSerializer.InspectTarget(invoice, InvoiceTestContracts.En16931(syntax));
            Assert.Contains(diagnostics, diagnostic => diagnostic.Location == "DeclaredTaxes[0].Category.ExemptionReason" &&
                diagnostic.Message.IndexOf("Header exemption", StringComparison.Ordinal) >= 0 &&
                diagnostic.Message.IndexOf("Line exemption", StringComparison.Ordinal) >= 0);
            Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(syntax)));
        }
        Assert.Equal("Line exemption", invoice.Lines[0].Tax.ExemptionReason);
        Assert.Equal("Header exemption", invoice.DeclaredTaxes[0].Category.ExemptionReason);
    }

    [Fact]
    public void ArbitrarySellerTaxSchemeRoundTripsInUblAndIsNotRelabeledForCii() {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Seller.TaxRegistrations.Add(new InvoiceTaxRegistration("NATIONAL-123", "PL-KRS"));
        byte[] ubl = InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(InvoiceSyntax.Ubl));
        InvoiceTaxRegistration registration = Assert.Single(InvoiceParser.Read(ubl).Invoice.Seller.TaxRegistrations,
            candidate => candidate.SchemeId == "PL-KRS");
        Assert.Equal("PL-KRS", registration.SchemeId);
        Assert.Equal("NATIONAL-123", registration.Identifier);
        InvoiceConversionResult cii = InvoiceConverter.Convert(ubl, InvoiceTestContracts.En16931());
        Assert.False(cii.Succeeded);
        Assert.Contains(cii.Diagnostics, diagnostic => diagnostic.Location == "Seller.TaxRegistrations[1]" &&
            diagnostic.Message.IndexOf("PL-KRS", StringComparison.Ordinal) >= 0);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ReservedSepaIdentifiersCannotChangeRolesInUbl(bool seller) {
        Invoice invoice = InvoiceFixture.Create();
        (seller ? invoice.Seller : invoice.Buyer).Identifiers.Add(new InvoiceIdentifier("party-id", "SEPA"));
        string path = seller ? "Seller.Identifiers" : "Buyer.Identifiers";
        Assert.Contains(InvoiceSerializer.InspectTarget(invoice, InvoiceTestContracts.En16931(InvoiceSyntax.Ubl)), d => d.Location == path);
        Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(InvoiceSyntax.Ubl)));
        byte[] cii = InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931());
        InvoiceReadResult read = InvoiceParser.Read(cii);
        Assert.True(read.HasCompleteMapping);
        Assert.Equal("party-id", Assert.Single((seller ? read.Invoice.Seller : read.Invoice.Buyer).Identifiers).Value);
        InvoiceConversionResult converted = InvoiceConverter.Convert(cii, InvoiceTestContracts.En16931(InvoiceSyntax.Ubl));
        Assert.False(converted.Succeeded);
        Assert.Contains(converted.Diagnostics, d => d.Location == path);
    }

    [Theory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public void ImportedRepresentativeWithoutVatIdentifierCannotBeRewrittenOrConverted(InvoiceSyntax syntax) {
        Invoice invoice = InvoiceFixture.Rich();
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(syntax))));
        XElement representative = document.Descendants().Single(e => e.Name.LocalName ==
            (syntax == InvoiceSyntax.Cii ? "SellerTaxRepresentativeTradeParty" : "TaxRepresentativeParty"));
        representative.Elements().Single(e => e.Name.LocalName ==
            (syntax == InvoiceSyntax.Cii ? "SpecifiedTaxRegistration" : "PartyTaxScheme")).Remove();
        byte[] source = Encoding.UTF8.GetBytes(document.ToString());
        InvoiceReadResult read = InvoiceParser.Read(source);
        Assert.Throws<InvalidDataException>(() => read.Write(InvoiceTestContracts.En16931(syntax), allowUnmappedDataLoss: true));
        InvoiceConversionResult converted = InvoiceConverter.Convert(source, InvoiceTestContracts.En16931(syntax == InvoiceSyntax.Cii ? InvoiceSyntax.Ubl : InvoiceSyntax.Cii));
        Assert.False(converted.Succeeded);
        Assert.Contains(converted.Diagnostics, d => d.Location == "TaxRepresentative.TaxRegistrations");
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" ")]
    public void TaxRepresentativeRequiresVatIdentifier(string? identifier) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.TaxRepresentative = new InvoiceParty { Name = "Representative", Address = new InvoiceAddress { CountryCode = "DE" } };
        if (identifier != null)
            invoice.TaxRepresentative.TaxRegistrations.Add(new InvoiceTaxRegistration(identifier, InvoiceTaxRegistration.VatScheme));
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics,
            d => d.Location == "TaxRepresentative.TaxRegistrations" && d.Code == "INV-REQUIRED");
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl })
            Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(syntax)));
    }

    [Theory]
    [InlineData("USD")]
    [InlineData("GBP")]
    [InlineData(null)]
    public void NonInvoiceCurrencyBreakdownsRequireExplicitLossWithoutPollutingTheModel(string? currency) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.TaxCurrency = "USD";
        invoice.TaxAmountInAccountingCurrency = 25m;
        XDocument document = XDocument.Parse(Encoding.UTF8.GetString(InvoiceSerializer.Write(invoice, InvoiceTestContracts.En16931(InvoiceSyntax.Ubl))));
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
        Assert.Throws<InvalidDataException>(() => read.Write(InvoiceTestContracts.En16931(InvoiceSyntax.Ubl)));
        Assert.False(InvoiceConverter.Convert(source, InvoiceTestContracts.En16931(InvoiceSyntax.Cii)).Succeeded);
        byte[] rewritten = read.Write(InvoiceTestContracts.En16931(InvoiceSyntax.Ubl), allowUnmappedDataLoss: true);
        InvoiceReadResult result = InvoiceParser.Read(rewritten);
        Assert.True(result.HasCompleteMapping);
        Assert.Equal(read.Invoice.DeclaredTaxes.Count, result.Invoice.DeclaredTaxes.Count);
        Assert.Equal(read.Invoice.TaxAmountInAccountingCurrency, result.Invoice.TaxAmountInAccountingCurrency);
        InvoiceReadResult cii = InvoiceParser.Read(read.Write(InvoiceTestContracts.En16931(InvoiceSyntax.Cii), allowUnmappedDataLoss: true));
        Assert.True(cii.HasCompleteMapping);
        Assert.Equal(result.Invoice.TaxAmountInAccountingCurrency, cii.Invoice.TaxAmountInAccountingCurrency);
    }
}
