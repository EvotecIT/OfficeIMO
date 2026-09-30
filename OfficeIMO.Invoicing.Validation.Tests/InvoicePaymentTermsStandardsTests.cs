using OfficeIMO.Invoicing.Tests;

namespace OfficeIMO.Invoicing.Validation.Tests;

public partial class InvoiceStandardsTests {
    [InvoiceStandardsTheory]
    [InlineData(InvoiceSyntax.Cii, InvoiceProfile.XRechnung)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.XRechnung)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.PeppolBis)]
    public async Task InvoiceLevelPaymentFallbacksPassTheirActualProfileRules(InvoiceSyntax syntax, InvoiceProfile profile) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payments[0].MeansCode = "59";
        invoice.Payments[0].Account = null;
        invoice.Payments[0].DebitedAccount = "DE89370400440532013000";
        invoice.DirectDebitMandateReference = "header-mandate";
        invoice.CreditorIdentifier = "DE98ZZZ09999999999";
        if (profile == InvoiceProfile.PeppolBis) {
            invoice.Seller.ElectronicAddress = new InvoiceIdentifier("1234567890128", "0088");
            invoice.Buyer.ElectronicAddress = new InvoiceIdentifier("1234567890135", "0088");
        }
        InvoiceXmlOptions options = InvoiceTestContracts.For(syntax, profile);
        byte[] xml = InvoiceSerializer.Write(invoice, options);
        InvoiceValidationReport report = await new InvoiceValidator(Bundle(), Runner()).ValidateAsync(xml, options.Release);
        Assert.True(report.IsValid, Report(report));
        Assert.True(InvoiceParser.Read(xml).HasCompleteMapping);
    }
    [InvoiceStandardsTheory]
    [InlineData("reference")]
    [InlineData("creditor")]
    [InlineData("mandate")]
    public async Task FacturXPermitsPaymentDataWithoutInventedPaymentMeans(string field) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payments.Clear();
        if (field == "reference") invoice.PaymentReference = "standalone-reference";
        if (field == "creditor") invoice.CreditorIdentifier = "standalone-creditor";
        if (field == "mandate") invoice.DirectDebitMandateReference = "standalone-mandate";
        InvoiceXmlOptions options = InvoiceTestContracts.FacturX(InvoiceProfile.En16931);
        byte[] xml = InvoiceSerializer.Write(invoice, options);
        var validator = new InvoiceValidator(Bundle(), Runner());
        InvoiceValidationReport report = await validator.ValidateAsync(xml, options.Release);
        Assert.True(report.IsValid, Report(report));
        InvoiceReadResult read = InvoiceParser.Read(xml);
        Assert.Empty(read.Invoice.Payments);
        Assert.True(read.HasCompleteMapping);
        Assert.True(InvoiceConverter.Convert(xml, options).Succeeded);
        InvoiceValidationReport rewrite = await validator.ValidateAsync(read.Write(options), options.Release);
        Assert.True(rewrite.IsValid, Report(rewrite));
    }

    [InvoiceStandardsTheory]
    [InlineData(InvoiceSyntax.Ubl)]
    public async Task En16931UblPreservesStandaloneSellerCreditorIdentifier(InvoiceSyntax syntax) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.Payments.Clear();
        invoice.CreditorIdentifier = "standalone-creditor";
        InvoiceXmlOptions options = InvoiceTestContracts.En16931(syntax);
        byte[] xml = InvoiceSerializer.Write(invoice, options);
        InvoiceValidationReport report = await new InvoiceValidator(Bundle(), Runner()).ValidateAsync(xml, options.Release);
        Assert.True(report.IsValid, Report(report));
        InvoiceReadResult read = InvoiceParser.Read(xml);
        Assert.Empty(read.Invoice.Payments);
        Assert.Equal(xml, read.Write(options));
    }
    [InvoiceStandardsTheory]
    [InlineData(InvoiceSyntax.Cii, InvoiceProfile.En16931, InvoiceSpecificationRelease.En16931_1_3_16)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.En16931, InvoiceSpecificationRelease.En16931_1_3_16)]
    [InlineData(InvoiceSyntax.Cii, InvoiceProfile.XRechnung, InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.XRechnung, InvoiceSpecificationRelease.XRechnung_3_0_2_2026_08_31)]
    [InlineData(InvoiceSyntax.Ubl, InvoiceProfile.PeppolBis, InvoiceSpecificationRelease.PeppolBis_3_0_21)]
    public async Task CurrentRuleReleasesPermitPositiveInvoicesWithoutDueDateOrTerms(InvoiceSyntax syntax, InvoiceProfile profile, InvoiceSpecificationRelease release) {
        Invoice invoice = InvoiceFixture.Create();
        invoice.DueDate = null;
        invoice.PaymentTerms = null;
        if (profile == InvoiceProfile.PeppolBis) {
            invoice.Seller.ElectronicAddress = new InvoiceIdentifier("1234567890128", "0088");
            invoice.Buyer.ElectronicAddress = new InvoiceIdentifier("1234567890135", "0088");
        }
        Assert.True(InvoiceModelValidator.Validate(invoice).Calculation!.PayableAmount > 0m);
        byte[] xml = InvoiceSerializer.Write(invoice, InvoiceTestContracts.For(syntax, profile));
        InvoiceValidationReport report = await new InvoiceValidator(Bundle(), Runner()).ValidateAsync(xml, release);
        Assert.True(report.IsValid, Report(report));
        InvoiceReadResult read = InvoiceParser.Read(xml);
        Assert.Null(read.Invoice.DueDate);
        Assert.Null(read.Invoice.PaymentTerms);
    }
}
