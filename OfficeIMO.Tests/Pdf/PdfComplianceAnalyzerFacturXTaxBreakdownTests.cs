using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfComplianceAnalyzerTests {
    [Fact]
    public void FacturXReadinessRequiresCiiTaxBreakdownEssentials() {
        var missingTaxTypeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeTradeTaxTypeCode: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingTaxCategoryOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeTradeTaxCategoryCode: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingTaxRateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeTradeTaxRate: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingTaxAmountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeTradeTaxCalculatedAmount: false, includeTradeTaxBasisAmount: false), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingTaxType = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingTaxTypeOptions),
            "einvoice-xml-tax-breakdown",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingTaxCategory = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingTaxCategoryOptions),
            "einvoice-xml-tax-breakdown",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingTaxRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingTaxRateOptions),
            "einvoice-xml-tax-breakdown",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingTaxAmount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingTaxAmountOptions),
            "einvoice-xml-tax-breakdown",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("ApplicableTradeTax TypeCode", missingTaxType.Diagnostic);
        Assert.Contains("ApplicableTradeTax CategoryCode", missingTaxCategory.Diagnostic);
        Assert.Contains("ApplicableTradeTax RateApplicablePercent", missingTaxRate.Diagnostic);
        Assert.Contains("ApplicableTradeTax BasisAmount", missingTaxAmount.Diagnostic);
        Assert.Contains("ApplicableTradeTax CalculatedAmount", missingTaxAmount.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiTaxBreakdownTypeCodeToBeVat() {
        var invalidTaxTypeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxTypeCodeValue: "GST"), "application/xml", PdfAssociatedFileRelationship.Data);
        var validTaxTypeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement invalidTaxType = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidTaxTypeOptions),
            "einvoice-xml-tax-breakdown",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement validTaxType = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, validTaxTypeOptions),
            "einvoice-xml-tax-breakdown",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("ApplicableTradeTax TypeCode", invalidTaxType.Diagnostic);
        Assert.Contains("VAT", invalidTaxType.Diagnostic);
        Assert.Contains("GST", invalidTaxType.Diagnostic);
        Assert.Contains("VAT trade tax type", validTaxType.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresTypeCodeOnEachCiiTaxBreakdown() {
        var missingSecondTaxTypeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", AddHeaderTradeTaxWithoutTypeCode(CreateCiiXml()), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingSecondTaxType = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingSecondTaxTypeOptions),
            "einvoice-xml-tax-breakdown",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("ApplicableTradeTax TypeCode", missingSecondTaxType.Diagnostic);
        Assert.Contains("ApplicableTradeTax #2", missingSecondTaxType.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCategoryCodeOnEachCiiTaxBreakdown() {
        var missingSecondTaxCategoryOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", AddHeaderTradeTaxWithoutCategoryCode(CreateCiiXml()), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingSecondTaxCategory = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingSecondTaxCategoryOptions),
            "einvoice-xml-tax-breakdown",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("ApplicableTradeTax CategoryCode", missingSecondTaxCategory.Diagnostic);
        Assert.Contains("ApplicableTradeTax #2", missingSecondTaxCategory.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiTaxCategoryCodeListValue() {
        var invalidCategoryOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "Q"), "application/xml", PdfAssociatedFileRelationship.Data);
        var intraCommunityCategoryOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "K", lineTradeTaxCategoryCodeValue: "K"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement invalidCategory = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidCategoryOptions),
            "einvoice-xml-tax-category-code",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement intraCommunityCategory = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, intraCommunityCategoryOptions),
            "einvoice-xml-tax-category-code",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("ApplicableTradeTax CategoryCode", invalidCategory.Diagnostic);
        Assert.Contains("UNCL5305", invalidCategory.Diagnostic);
        Assert.Contains("Q", invalidCategory.Diagnostic);
        Assert.Contains("UNCL5305", intraCommunityCategory.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresNotSubjectTaxBreakdownExclusivity() {
        var lineNotSubjectWithoutHeaderBreakdownOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(
                lineTradeTaxCategoryCodeValue: "O",
                includeLineTradeTaxRate: false,
                headerTradeTaxExemptionReasonValue: "Not subject to VAT"), "application/xml", PdfAssociatedFileRelationship.Data);
        var allowanceNotSubjectWithoutHeaderBreakdownOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile(
                "factur-x.xml",
                AddHeaderAllowanceCharge(
                    CreateCiiXml(
                        includeAllowanceTotalAmount: true,
                        allowanceTotalAmount: "10.00",
                        taxBasisTotalAmount: "90.00",
                        taxTotalAmount: "20.70",
                        grandTotalAmount: "110.70",
                        headerTradeTaxBasisAmountValue: "90.00",
                        headerTradeTaxCalculatedAmountValue: "20.70"),
                    false,
                    "O",
                    "10.00"),
                "application/xml",
                PdfAssociatedFileRelationship.Data);
        var duplicateNotSubjectBreakdownOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", AddHeaderTradeTax(CreateCiiXml(
                headerTradeTaxCategoryCodeValue: "O",
                lineTradeTaxCategoryCodeValue: "O",
                includeTradeTaxRate: false,
                includeLineTradeTaxRate: false,
                taxTotalAmount: "0.00",
                grandTotalAmount: "100.00",
                headerTradeTaxCalculatedAmountValue: "0.00",
                headerTradeTaxExemptionReasonValue: "Not subject to VAT"), "O", false), "application/xml", PdfAssociatedFileRelationship.Data);
        var mixedNotSubjectBreakdownOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", AddHeaderTradeTax(CreateCiiXml(
                headerTradeTaxCategoryCodeValue: "O",
                lineTradeTaxCategoryCodeValue: "O",
                includeTradeTaxRate: false,
                includeLineTradeTaxRate: false,
                taxTotalAmount: "0.00",
                grandTotalAmount: "100.00",
                headerTradeTaxCalculatedAmountValue: "0.00",
                headerTradeTaxExemptionReasonValue: "Not subject to VAT"), "S", true), "application/xml", PdfAssociatedFileRelationship.Data);
        var validNotSubjectBreakdownOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(
                headerTradeTaxCategoryCodeValue: "O",
                lineTradeTaxCategoryCodeValue: "O",
                includeTradeTaxRate: false,
                includeLineTradeTaxRate: false,
                taxTotalAmount: "0.00",
                grandTotalAmount: "100.00",
                headerTradeTaxCalculatedAmountValue: "0.00",
                headerTradeTaxExemptionReasonValue: "Not subject to VAT"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement lineNotSubjectWithoutHeaderBreakdown = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, lineNotSubjectWithoutHeaderBreakdownOptions),
            "einvoice-xml-tax-category-code",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement allowanceNotSubjectWithoutHeaderBreakdown = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, allowanceNotSubjectWithoutHeaderBreakdownOptions),
            "einvoice-xml-tax-category-code",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement duplicateNotSubjectBreakdown = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, duplicateNotSubjectBreakdownOptions),
            "einvoice-xml-tax-category-code",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement mixedNotSubjectBreakdown = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, mixedNotSubjectBreakdownOptions),
            "einvoice-xml-tax-category-code",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement validNotSubjectBreakdown = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, validNotSubjectBreakdownOptions),
            "einvoice-xml-tax-category-code",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("exactly one", lineNotSubjectWithoutHeaderBreakdown.Diagnostic);
        Assert.Contains("Header category O breakdown count: 0", lineNotSubjectWithoutHeaderBreakdown.Diagnostic);
        Assert.Contains("document-level allowance", allowanceNotSubjectWithoutHeaderBreakdown.Diagnostic);
        Assert.Contains("Header category O breakdown count: 0", allowanceNotSubjectWithoutHeaderBreakdown.Diagnostic);
        Assert.Contains("Header category O breakdown count: 2", duplicateNotSubjectBreakdown.Diagnostic);
        Assert.Contains("Other header tax categories: S", mixedNotSubjectBreakdown.Diagnostic);
        Assert.Contains("category-O line, allowance, and charge breakdown exclusivity", validNotSubjectBreakdown.Diagnostic);
    }


}
