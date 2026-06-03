using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfComplianceAnalyzerTests {
    [Fact]
    public void FacturXReadinessRequiresCiiSettlementSummaryEssentials() {
        var missingCurrencyOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeInvoiceCurrencyCode: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingTaxOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeApplicableTradeTax: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingTaxTotalsOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeTaxTotals: false), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingCurrency = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingCurrencyOptions),
            "einvoice-xml-settlement-summary",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingTax = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingTaxOptions),
            "einvoice-xml-settlement-summary",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingTaxTotals = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingTaxTotalsOptions),
            "einvoice-xml-settlement-summary",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("InvoiceCurrencyCode", missingCurrency.Diagnostic);
        Assert.Contains("ApplicableTradeTax", missingTax.Diagnostic);
        Assert.Contains("TaxBasisTotalAmount", missingTaxTotals.Diagnostic);
        Assert.Contains("TaxTotalAmount", missingTaxTotals.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiCurrencyConsistency() {
        var missingInvoiceCurrencyOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeInvoiceCurrencyCode: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingAmountCurrencyOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(amountCurrencyId: null), "application/xml", PdfAssociatedFileRelationship.Data);
        var mismatchedAmountCurrencyOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(amountCurrencyId: "USD"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingInvoiceCurrency = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingInvoiceCurrencyOptions),
            "einvoice-xml-currency-consistency",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingAmountCurrency = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingAmountCurrencyOptions),
            "einvoice-xml-currency-consistency",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement mismatchedAmountCurrency = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, mismatchedAmountCurrencyOptions),
            "einvoice-xml-currency-consistency",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("InvoiceCurrencyCode", missingInvoiceCurrency.Diagnostic);
        Assert.Contains("when present", missingAmountCurrency.Diagnostic);
        Assert.Contains("InvoiceCurrencyCode EUR", mismatchedAmountCurrency.Diagnostic);
        Assert.Contains("currencyID USD", mismatchedAmountCurrency.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiCurrencyCodeListValue() {
        var invalidInvoiceCurrencyOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(invoiceCurrencyCodeValue: "EURO", amountCurrencyId: "EURO"), "application/xml", PdfAssociatedFileRelationship.Data);
        var invalidAmountCurrencyOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(amountCurrencyId: "ZZZ"), "application/xml", PdfAssociatedFileRelationship.Data);
        var yenCurrencyOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(invoiceCurrencyCodeValue: "JPY", amountCurrencyId: "JPY"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement invalidInvoiceCurrency = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidInvoiceCurrencyOptions),
            "einvoice-xml-currency-code",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement invalidAmountCurrency = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidAmountCurrencyOptions),
            "einvoice-xml-currency-code",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement yenCurrency = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, yenCurrencyOptions),
            "einvoice-xml-currency-code",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("ISO 4217", invalidInvoiceCurrency.Diagnostic);
        Assert.Contains("EURO", invalidInvoiceCurrency.Diagnostic);
        Assert.Contains("ZZZ", invalidAmountCurrency.Diagnostic);
        Assert.Contains("ISO 4217", yenCurrency.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiPaymentInstructionEssentials() {
        var missingPaymentMeansOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includePaymentMeans: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingPaymentTypeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includePaymentMeansTypeCode: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingAccountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeCreditorAccountId: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var cashWithoutAccountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(paymentMeansTypeCodeValue: "10", includeCreditorAccount: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var transferAccountOnDifferentPaymentMeansOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXmlWithTransferPaymentMeansMissingOwnAccount(), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingPaymentMeans = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingPaymentMeansOptions),
            "einvoice-xml-payment-instructions",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingPaymentType = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingPaymentTypeOptions),
            "einvoice-xml-payment-instructions",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingAccount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingAccountOptions),
            "einvoice-xml-payment-instructions",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement cashWithoutAccount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, cashWithoutAccountOptions),
            "einvoice-xml-payment-instructions",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement transferAccountOnDifferentPaymentMeans = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, transferAccountOnDifferentPaymentMeansOptions),
            "einvoice-xml-payment-instructions",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("SpecifiedTradeSettlementPaymentMeans", missingPaymentMeans.Diagnostic);
        Assert.Contains("SpecifiedTradeSettlementPaymentMeans TypeCode", missingPaymentType.Diagnostic);
        Assert.Contains("PayeePartyCreditorFinancialAccount IBANID or ProprietaryID", missingAccount.Diagnostic);
        Assert.Contains("does not require creditor account identifiers", cashWithoutAccount.Diagnostic);
        Assert.Contains("PayeePartyCreditorFinancialAccount on SpecifiedTradeSettlementPaymentMeans #1", transferAccountOnDifferentPaymentMeans.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiPaymentMeansCodeListValue() {
        var invalidTypeCodeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(paymentMeansTypeCodeValue: "999"), "application/xml", PdfAssociatedFileRelationship.Data);
        var creditTransferOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(paymentMeansTypeCodeValue: "30"), "application/xml", PdfAssociatedFileRelationship.Data);
        var mutuallyDefinedOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(paymentMeansTypeCodeValue: "ZZZ"), "application/xml", PdfAssociatedFileRelationship.Data);
        var secondPaymentMeansMissingTypeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXmlWithSecondPaymentMeansMissingTypeCode(), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement invalidTypeCode = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidTypeCodeOptions),
            "einvoice-xml-payment-means-code",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement creditTransfer = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, creditTransferOptions),
            "einvoice-xml-payment-means-code",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement mutuallyDefined = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, mutuallyDefinedOptions),
            "einvoice-xml-payment-means-code",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement secondPaymentMeansMissingType = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, secondPaymentMeansMissingTypeOptions),
            "einvoice-xml-payment-means-code",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("SpecifiedTradeSettlementPaymentMeans TypeCode", invalidTypeCode.Diagnostic);
        Assert.Contains("UNCL4461", invalidTypeCode.Diagnostic);
        Assert.Contains("999", invalidTypeCode.Diagnostic);
        Assert.Contains("UNCL4461", creditTransfer.Diagnostic);
        Assert.Contains("UNCL4461", mutuallyDefined.Diagnostic);
        Assert.Contains("SpecifiedTradeSettlementPaymentMeans TypeCode on SpecifiedTradeSettlementPaymentMeans #2", secondPaymentMeansMissingType.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiPaymentAccountFormat() {
        var invalidIbanOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(creditorAccountIban: "PL61109010140000071219812875"), "application/xml", PdfAssociatedFileRelationship.Data);
        var proprietaryAccountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(useCreditorProprietaryAccountId: true, creditorProprietaryAccountId: "ACCOUNT-001"), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingAccountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeCreditorAccountId: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var cashWithoutAccountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(paymentMeansTypeCodeValue: "10", includeCreditorAccount: false), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement invalidIban = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidIbanOptions),
            "einvoice-xml-payment-account-format",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement proprietaryAccount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, proprietaryAccountOptions),
            "einvoice-xml-payment-account-format",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement missingAccount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingAccountOptions),
            "einvoice-xml-payment-account-format",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement cashWithoutAccount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, cashWithoutAccountOptions),
            "einvoice-xml-payment-account-format",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("IBANID", invalidIban.Diagnostic);
        Assert.Contains("checksum", invalidIban.Diagnostic);
        Assert.Contains("PL61109010140000071219812875", invalidIban.Diagnostic);
        Assert.Contains("creditor account identifiers are present", proprietaryAccount.Diagnostic);
        Assert.Contains("PayeePartyCreditorFinancialAccount IBANID or ProprietaryID", missingAccount.Diagnostic);
        Assert.Contains("does not require creditor account identifiers", cashWithoutAccount.Diagnostic);
    }


}
