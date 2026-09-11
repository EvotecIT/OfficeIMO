using System.Xml.Linq;

namespace OfficeIMO.Invoicing.Tests;

public class InvoiceRequiredTermsTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AlternativeSellerIdentifiersAndNonTransferPaymentsRemainSupported(bool legal) {
        Invoice invoice = InvoiceFixture.WithTaxCategory("E");
        invoice.Seller.VatIdentifier = null;
        invoice.Seller.TaxRegistration = "local-tax-id";
        if (legal) invoice.Seller.LegalRegistration = new InvoiceIdentifier("HRB 1");
        else invoice.Seller.Identifiers.Add(new InvoiceIdentifier("seller-1"));
        invoice.Payment!.MeansCode = "10";
        invoice.Payment.Accounts.Clear();
        invoice.Lines[0].Tax.ExemptionReason = null;
        invoice.Lines[0].Tax.ExemptionReasonCode = "VATEX-EU-132";
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl }) {
            InvoiceReadResult read = InvoiceParser.Read(InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax)));
            Assert.True(read.HasCompleteMapping);
            Assert.Equal("VATEX-EU-132", Assert.Single(read.Invoice.DeclaredTaxes).Category.ExemptionReasonCode);
            Assert.Empty(read.Invoice.Payment!.Accounts);
        }
    }

    [Theory]
    [InlineData("E")]
    [InlineData("AE")]
    [InlineData("G")]
    [InlineData("K")]
    [InlineData("O")]
    public void ExemptionCanBeSuppliedOnlyOnTheHeaderAndSurviveEditing(string code) {
        Invoice invoice = InvoiceFixture.WithTaxCategory(code);
        InvoiceCalculator.UpdateDeclaredAmounts(invoice);
        invoice.Lines[0].Tax.ExemptionReason = null;
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl }) {
            byte[] xml = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax));
            InvoiceReadResult read = InvoiceParser.Read(xml);
            Assert.True(read.HasCompleteMapping);
            Assert.Equal("Exemption applies", Assert.Single(read.Invoice.DeclaredTaxes).Category.ExemptionReason);
            read.Invoice.Lines[0].UnitPrice = 150m;
            InvoiceCalculator.UpdateDeclaredAmounts(read.Invoice);
            InvoiceReadResult edited = InvoiceParser.Read(read.Write());
            Assert.Equal("Exemption applies", Assert.Single(edited.Invoice.DeclaredTaxes).Category.ExemptionReason);
            Assert.Equal(150m, Assert.Single(edited.Invoice.DeclaredTaxes).TaxableAmount);
        }
    }

    [Theory]
    [InlineData("E")]
    [InlineData("AE")]
    [InlineData("G")]
    [InlineData("K")]
    [InlineData("O")]
    public void MissingHeaderExemptionBlocksRewriteAndConversion(string code) {
        Invoice invoice = InvoiceFixture.WithTaxCategory(code);
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl }) {
            var xml = System.Xml.Linq.XDocument.Parse(Encoding.UTF8.GetString(InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax))));
            xml.Descendants().Where(e => e.Name.LocalName == "ExemptionReason" || e.Name.LocalName == "TaxExemptionReason").Remove();
            byte[] bytes = Encoding.UTF8.GetBytes(xml.ToString());
            InvoiceReadResult read = InvoiceParser.Read(bytes);
            Assert.Contains(InvoiceModelValidator.Validate(read.Invoice).Diagnostics, d => d.Code == "INV-VAT-EXEMPTION");
            Assert.Throws<InvalidDataException>(() => read.Write(allowUnmappedDataLoss: true));
            Assert.False(InvoiceConverter.Convert(bytes, new InvoiceXmlOptions(syntax == InvoiceSyntax.Cii ? InvoiceSyntax.Ubl : InvoiceSyntax.Cii)).Succeeded);
        }
    }

    [Theory]
    [InlineData("S")]
    [InlineData("Z")]
    [InlineData("L")]
    [InlineData("M")]
    public void TaxableCategoriesRejectExemptionReasons(string code) {
        Invoice invoice = InvoiceFixture.WithTaxCategory(code);
        invoice.Lines[0].Tax.ExemptionReason = "Does not apply";
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics, d => d.Code == "INV-VAT-EXEMPTION");
    }

    [Theory]
    [InlineData("Z")]
    [InlineData("E")]
    [InlineData("AE")]
    [InlineData("G")]
    [InlineData("K")]
    [InlineData("O")]
    public void ZeroTaxCategoriesCannotUseRoundingToleranceToDeclareNonzeroTax(string code) {
        Invoice invoice = InvoiceFixture.WithTaxCategory(code);
        InvoiceCalculator.UpdateDeclaredAmounts(invoice);
        invoice.DeclaredTotals = null;
        invoice.DeclaredTaxes[0].TaxAmount = 0.01m;
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics, d => d.Code == "INV-VAT-ZERO");
    }

    [Theory]
    [InlineData("seller-tax")]
    [InlineData("reverse-buyer")]
    [InlineData("community-buyer")]
    [InlineData("community-date")]
    [InlineData("community-country")]
    [InlineData("outside-id")]
    [InlineData("outside-mix")]
    public void RelatedVatPrerequisitesAreEnforced(string violation) {
        Invoice invoice = InvoiceFixture.WithTaxCategory(violation.StartsWith("community", StringComparison.Ordinal) ? "K" :
            violation.StartsWith("outside", StringComparison.Ordinal) ? "O" : violation == "reverse-buyer" ? "AE" : "S");
        switch (violation) {
            case "seller-tax": invoice.Seller.VatIdentifier = null; invoice.Seller.LegalRegistration = new InvoiceIdentifier("HRB 1"); break;
            case "reverse-buyer":
            case "community-buyer": invoice.Buyer.VatIdentifier = null; break;
            case "community-date": invoice.Delivery!.Date = null; break;
            case "community-country": invoice.Delivery!.Address = null; break;
            case "outside-id": invoice.Buyer.VatIdentifier = "DE987654321"; break;
            case "outside-mix": invoice.Lines.Add(InvoiceFixture.Create().Lines[0]); invoice.Lines[1].Id = "2"; break;
        }
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics, d => d.Code == "INV-VAT-IDENTIFIER" || d.Code == "INV-VAT-DELIVERY" || d.Code == "INV-VAT-MIX");
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl })
            Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax)));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" ")]
    public void SellerMustHaveABusinessLegalOrVatIdentifier(string? vat) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Seller.Identifiers.Clear();
        invoice.Seller.LegalRegistration = null;
        invoice.Seller.VatIdentifier = vat;
        invoice.Seller.TaxRegistration = "other-tax-id";
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics, d => d.Code == "INV-SELLER-ID");
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl })
            Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax)));
    }

    [Theory]
    [InlineData("30")]
    [InlineData("58")]
    public void CreditTransferRequiresAnAccount(string means) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payment!.MeansCode = means;
        invoice.Payment.Accounts.Clear();
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics, d => d.Code == "INV-PAYMENT-ACCOUNT");
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl })
            Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax)));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" ")]
    public void ExemptVatRequiresAReasonOnTheResultingBreakdown(string? reason) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Lines[0].Tax = new InvoiceTaxCategory { Code = "E", Rate = 0m, ExemptionReason = reason, ExemptionReasonCode = reason };
        Assert.Contains(InvoiceModelValidator.Validate(invoice).Diagnostics, d => d.Code == "INV-VAT-EXEMPTION");
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl })
            Assert.Throws<InvalidDataException>(() => InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax)));
    }
}
