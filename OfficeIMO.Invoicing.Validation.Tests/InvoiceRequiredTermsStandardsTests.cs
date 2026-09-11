using System.Text;
using System.Xml.Linq;
using OfficeIMO.Invoicing.Tests;

namespace OfficeIMO.Invoicing.Validation.Tests;

public partial class InvoiceStandardsTests {
    [InvoiceStandardsTheory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public async Task PinnedRulesExposeTheSyntaxSpecificZeroIgicRestriction(InvoiceSyntax syntax) {
        Invoice invoice = InvoiceFixture.WithTaxCategory("L");
        invoice.Lines[0].Tax.Rate = 0m;
        byte[] xml = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax));
        InvoiceValidationReport report = await new InvoiceValidator(Bundle(), Runner()).ValidateAsync(xml, InvoiceRulesRelease.En16931_1_3_16);
        Assert.Equal(InvoiceValidationStatus.Passed, report.SchemaStatus);
        if (syntax == InvoiceSyntax.Cii) {
            Assert.Equal(InvoiceValidationStatus.Invalid, report.BusinessRulesStatus);
            Assert.Contains(report.Diagnostics, d => d.Code == "BR-AF-05");
        } else Assert.True(report.IsValid, Report(report));
    }

    [InvoiceStandardsTheory]
    [InlineData("S")]
    [InlineData("Z")]
    [InlineData("E")]
    [InlineData("AE")]
    [InlineData("G")]
    [InlineData("K")]
    [InlineData("O")]
    [InlineData("L")]
    [InlineData("M")]
    public async Task SupportedVatCategoriesAndHeaderReasonsPassBothAuthorityMappings(string code) {
        Invoice invoice = InvoiceFixture.WithTaxCategory(code);
        InvoiceCalculator.UpdateDeclaredAmounts(invoice);
        invoice.Lines[0].Tax.ExemptionReason = null;
        var validator = new InvoiceValidator(Bundle(), Runner());
        foreach (InvoiceSyntax syntax in new[] { InvoiceSyntax.Cii, InvoiceSyntax.Ubl }) {
            byte[] xml = InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax));
            InvoiceValidationReport report = await validator.ValidateAsync(xml, InvoiceRulesRelease.En16931_1_3_16);
            Assert.True(report.IsValid, syntax + ": " + Report(report));
            InvoiceReadResult read = InvoiceParser.Read(xml);
            read.Invoice.Lines[0].UnitPrice = 150m;
            InvoiceCalculator.UpdateDeclaredAmounts(read.Invoice);
            InvoiceValidationReport edited = await validator.ValidateAsync(read.Write(), InvoiceRulesRelease.En16931_1_3_16);
            Assert.True(edited.IsValid, syntax + ": " + Report(edited));
        }
    }

    [InvoiceStandardsTheory]
    [InlineData(InvoiceSyntax.Cii)]
    [InlineData(InvoiceSyntax.Ubl)]
    public async Task MissingMandatoryTermsMatchTheAuthorityDiagnostics(InvoiceSyntax syntax) {
        Invoice invoice = InvoiceFixture.WithTaxCategory("E");
        XDocument xml = XDocument.Parse(Encoding.UTF8.GetString(InvoiceSerializer.Write(invoice, new InvoiceXmlOptions(syntax))));
        xml.Descendants().Where(e => e.Name.LocalName == "ExemptionReason" || e.Name.LocalName == "TaxExemptionReason").Remove();
        XElement seller = xml.Descendants().Single(e => e.Name.LocalName == (syntax == InvoiceSyntax.Cii ? "SellerTradeParty" : "AccountingSupplierParty"));
        seller.Descendants().Where(e => e.Name.LocalName == "SpecifiedTaxRegistration" || e.Name.LocalName == "PartyTaxScheme").Remove();
        xml.Descendants().Where(e => e.Name.LocalName == "PayeePartyCreditorFinancialAccount" || e.Name.LocalName == "PayeeFinancialAccount").Remove();
        byte[] bytes = Encoding.UTF8.GetBytes(xml.ToString());
        InvoiceValidationReport report = await new InvoiceValidator(Bundle(), Runner()).ValidateAsync(bytes, InvoiceRulesRelease.En16931_1_3_16);
        Assert.Equal(InvoiceValidationStatus.Passed, report.SchemaStatus);
        Assert.Equal(InvoiceValidationStatus.Invalid, report.BusinessRulesStatus);
        foreach (string rule in new[] { "BR-CO-26", "BR-E-10", syntax == InvoiceSyntax.Cii ? "CII-SR-470" : "BR-61" })
            Assert.Contains(report.Diagnostics, d => d.Code == rule);
        InvoiceReadResult read = InvoiceParser.Read(bytes);
        Assert.Throws<InvalidDataException>(() => read.Write(allowUnmappedDataLoss: true));
        Assert.False(InvoiceConverter.Convert(bytes, new InvoiceXmlOptions(syntax == InvoiceSyntax.Cii ? InvoiceSyntax.Ubl : InvoiceSyntax.Cii)).Succeeded);
    }
}
