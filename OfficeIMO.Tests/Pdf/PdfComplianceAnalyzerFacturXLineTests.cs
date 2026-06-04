using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfComplianceAnalyzerTests {
    [Fact]
    public void FacturXReadinessRequiresCiiLineItemEssentials() {
        var missingLineOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeLineItem: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingProductOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeLineItemProductName: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingTotalOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeLineTotalAmount: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingUnitCodeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeLineBilledQuantityUnitCode: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var secondLineMissingProductOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateTwoLineCiiXmlWithSecondLineMissingProductName(), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingLine = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingLineOptions),
            "einvoice-xml-line-item",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingProduct = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingProductOptions),
            "einvoice-xml-line-item",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingTotal = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingTotalOptions),
            "einvoice-xml-line-item",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingUnitCode = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingUnitCodeOptions),
            "einvoice-xml-line-item",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement secondLineMissingProduct = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, secondLineMissingProductOptions),
            "einvoice-xml-line-item",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("IncludedSupplyChainTradeLineItem", missingLine.Diagnostic);
        Assert.Contains("AssociatedDocumentLineDocument LineID", missingLine.Diagnostic);
        Assert.Contains("SpecifiedTradeProduct Name", missingLine.Diagnostic);
        Assert.Contains("SpecifiedLineTradeDelivery BilledQuantity", missingLine.Diagnostic);
        Assert.Contains("SpecifiedLineTradeDelivery BilledQuantity unitCode", missingLine.Diagnostic);
        Assert.Contains("SpecifiedTradeSettlementLineMonetarySummation LineTotalAmount", missingLine.Diagnostic);
        Assert.Contains("SpecifiedTradeProduct Name", missingProduct.Diagnostic);
        Assert.Contains("SpecifiedTradeSettlementLineMonetarySummation LineTotalAmount", missingTotal.Diagnostic);
        Assert.Contains("SpecifiedLineTradeDelivery BilledQuantity unitCode", missingUnitCode.Diagnostic);
        Assert.Contains("line 2 SpecifiedTradeProduct Name", secondLineMissingProduct.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiLinePricingEssentials() {
        var missingAgreementOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeLineTradeAgreement: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingPriceOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeLinePriceChargeAmount: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var grossOnlyPriceOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateGrossLinePriceCiiXml(), "application/xml", PdfAssociatedFileRelationship.Data);
        var secondLineMissingPriceOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateTwoLineCiiXmlWithSecondLineMissingPriceCharge(), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingAgreement = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingAgreementOptions),
            "einvoice-xml-line-pricing",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingPrice = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingPriceOptions),
            "einvoice-xml-line-pricing",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement grossOnlyPrice = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, grossOnlyPriceOptions),
            "einvoice-xml-line-pricing",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement secondLineMissingPrice = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, secondLineMissingPriceOptions),
            "einvoice-xml-line-pricing",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("SpecifiedLineTradeAgreement", missingAgreement.Diagnostic);
        Assert.Contains("NetPriceProductTradePrice", missingAgreement.Diagnostic);
        Assert.Contains("ChargeAmount", missingPrice.Diagnostic);
        Assert.Contains("NetPriceProductTradePrice", grossOnlyPrice.Diagnostic);
        Assert.Contains("line 2 NetPriceProductTradePrice ChargeAmount", secondLineMissingPrice.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiUnitCodeListValue() {
        var invalidUnitCodeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(lineBilledQuantityUnitCodeValue: "QQQ"), "application/xml", PdfAssociatedFileRelationship.Data);
        var packagingUnitCodeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(lineBilledQuantityUnitCodeValue: "XBX"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement invalidUnitCode = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidUnitCodeOptions),
            "einvoice-xml-unit-code",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement packagingUnitCode = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, packagingUnitCodeOptions),
            "einvoice-xml-unit-code",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("UN/ECE Recommendation 20", invalidUnitCode.Diagnostic);
        Assert.Contains("QQQ", invalidUnitCode.Diagnostic);
        Assert.Contains("Rec 21 X-prefixed packaging codes", packagingUnitCode.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiLineAmountConsistency() {
        var mismatchedLineAmountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(linePriceChargeAmountValue: "90.00"), "application/xml", PdfAssociatedFileRelationship.Data);
        var normalizedLineAmountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(linePriceChargeAmountValue: "50.00", lineBilledQuantityValue: "2.00"), "application/xml", PdfAssociatedFileRelationship.Data);
        var basisQuantityLineAmountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(lineTotalAmount: "5.00", linePriceChargeAmountValue: "250.00", linePriceBasisQuantityValue: "100", lineBilledQuantityValue: "2.00"), "application/xml", PdfAssociatedFileRelationship.Data);
        var allowanceLineAmountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXmlWithLineAllowance(), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement mismatchedLineAmount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, mismatchedLineAmountOptions),
            "einvoice-xml-line-amount-consistency",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement normalizedLineAmount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, normalizedLineAmountOptions),
            "einvoice-xml-line-amount-consistency",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement basisQuantityLineAmount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, basisQuantityLineAmountOptions),
            "einvoice-xml-line-amount-consistency",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement allowanceLineAmount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, allowanceLineAmountOptions),
            "einvoice-xml-line-amount-consistency",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("BilledQuantity times ProductTradePrice ChargeAmount divided by BasisQuantity", mismatchedLineAmount.Diagnostic);
        Assert.Contains("1", mismatchedLineAmount.Diagnostic);
        Assert.Contains("line quantity, price, and line total amount", normalizedLineAmount.Diagnostic);
        Assert.Contains("line quantity, price, and line total amount", basisQuantityLineAmount.Diagnostic);
        Assert.Contains("line quantity, price, and line total amount", allowanceLineAmount.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiLineTaxEssentials() {
        var missingLineTaxOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeLineTradeTax: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingLineTaxCategoryOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeLineTradeTaxCategoryCode: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingLineTaxRateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeLineTradeTaxRate: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var secondLineMissingLineTaxOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateTwoLineCiiXmlWithSecondLineMissingLineTax(), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingLineTax = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingLineTaxOptions),
            "einvoice-xml-line-tax",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingLineTaxCategory = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingLineTaxCategoryOptions),
            "einvoice-xml-line-tax",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingLineTaxRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingLineTaxRateOptions),
            "einvoice-xml-line-tax",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement secondLineMissingLineTax = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, secondLineMissingLineTaxOptions),
            "einvoice-xml-line-tax",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("ApplicableTradeTax", missingLineTax.Diagnostic);
        Assert.Contains("ApplicableTradeTax CategoryCode", missingLineTaxCategory.Diagnostic);
        Assert.Contains("ApplicableTradeTax RateApplicablePercent", missingLineTaxRate.Diagnostic);
        Assert.Contains("line 2 ApplicableTradeTax", secondLineMissingLineTax.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiLineTaxTypeCodeToBeVat() {
        var invalidLineTaxTypeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(lineTradeTaxTypeCodeValue: "GST"), "application/xml", PdfAssociatedFileRelationship.Data);
        var validLineTaxTypeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement invalidLineTaxType = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidLineTaxTypeOptions),
            "einvoice-xml-line-tax",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement validLineTaxType = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, validLineTaxTypeOptions),
            "einvoice-xml-line-tax",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("ApplicableTradeTax TypeCode", invalidLineTaxType.Diagnostic);
        Assert.Contains("VAT", invalidLineTaxType.Diagnostic);
        Assert.Contains("GST", invalidLineTaxType.Diagnostic);
        Assert.Contains("VAT line trade settlement tax type", validLineTaxType.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessForbidsCiiLineTaxRateForNotSubjectCategory() {
        var notSubjectWithRateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(lineTradeTaxCategoryCodeValue: "O", lineTradeTaxRateValue: "23"), "application/xml", PdfAssociatedFileRelationship.Data);
        var notSubjectWithoutRateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(lineTradeTaxCategoryCodeValue: "O", includeLineTradeTaxRate: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var standardWithoutRateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeLineTradeTaxRate: false), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement notSubjectWithRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, notSubjectWithRateOptions),
            "einvoice-xml-line-tax",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement notSubjectWithoutRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, notSubjectWithoutRateOptions),
            "einvoice-xml-line-tax",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement standardWithoutRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, standardWithoutRateOptions),
            "einvoice-xml-line-tax",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("category O", notSubjectWithRate.Diagnostic);
        Assert.Contains("Forbidden line tax rate categories: O", notSubjectWithRate.Diagnostic);
        Assert.Contains("VAT line trade settlement tax type", notSubjectWithoutRate.Diagnostic);
        Assert.Contains("Missing line tax rate categories: S", standardWithoutRate.Diagnostic);
    }


}
