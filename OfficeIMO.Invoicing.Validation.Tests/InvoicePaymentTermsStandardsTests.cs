using OfficeIMO.Invoicing.Tests;

namespace OfficeIMO.Invoicing.Validation.Tests;

public partial class InvoiceStandardsTests {
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
