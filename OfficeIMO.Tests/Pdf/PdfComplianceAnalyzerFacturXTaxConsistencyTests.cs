using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfComplianceAnalyzerTests {
    [Fact]
    public void FacturXReadinessRequiresCiiTaxCategoryConsistency() {
        var mismatchedLineRateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(lineTradeTaxRateValue: "8"), "application/xml", PdfAssociatedFileRelationship.Data);
        var normalizedRateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(lineTradeTaxRateValue: "23.00"), "application/xml", PdfAssociatedFileRelationship.Data);
        var mismatchedAllowanceRateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", AddHeaderAllowanceCharge(CreateCiiXml(), charge: false, categoryCode: "S", actualAmount: "10.00", includeRate: true, rateValue: "0"), "application/xml", PdfAssociatedFileRelationship.Data);
        var normalizedChargeRateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", AddHeaderAllowanceCharge(CreateCiiXml(), charge: true, categoryCode: "S", actualAmount: "10.00", includeRate: true, rateValue: "23.00"), "application/xml", PdfAssociatedFileRelationship.Data);
        var notSubjectWithoutRateOptions = new PdfOptions()
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
        var allowanceNotSubjectWithoutRateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", AddHeaderAllowanceCharge(CreateCiiXml(
                headerTradeTaxCategoryCodeValue: "O",
                lineTradeTaxCategoryCodeValue: "O",
                includeTradeTaxRate: false,
                includeLineTradeTaxRate: false,
                taxTotalAmount: "0.00",
                grandTotalAmount: "95.00",
                headerTradeTaxBasisAmountValue: "95.00",
                headerTradeTaxCalculatedAmountValue: "0.00",
                headerTradeTaxExemptionReasonValue: "Not subject to VAT"), charge: false, categoryCode: "O", actualAmount: "5.00"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement mismatchedLineRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, mismatchedLineRateOptions),
            "einvoice-xml-tax-category-consistency",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement normalizedRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, normalizedRateOptions),
            "einvoice-xml-tax-category-consistency",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement mismatchedAllowanceRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, mismatchedAllowanceRateOptions),
            "einvoice-xml-tax-category-consistency",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement normalizedChargeRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, normalizedChargeRateOptions),
            "einvoice-xml-tax-category-consistency",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement notSubjectWithoutRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, notSubjectWithoutRateOptions),
            "einvoice-xml-tax-category-consistency",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement allowanceNotSubjectWithoutRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, allowanceNotSubjectWithoutRateOptions),
            "einvoice-xml-tax-category-consistency",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("Unmatched line tax category/rate", mismatchedLineRate.Diagnostic);
        Assert.Contains("S/8", mismatchedLineRate.Diagnostic);
        Assert.Contains("Unmatched allowance/charge tax category/rate", mismatchedAllowanceRate.Diagnostic);
        Assert.Contains("S/0", mismatchedAllowanceRate.Diagnostic);
        Assert.Contains("match the header tax breakdown", normalizedRate.Diagnostic);
        Assert.Contains("allowance/charge tax category/rate", normalizedChargeRate.Diagnostic);
        Assert.Contains("category-O rate absence", notSubjectWithoutRate.Diagnostic);
        Assert.Contains("category-O rate absence", allowanceNotSubjectWithoutRate.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiTaxTotalConsistency() {
        var mismatchedBasisOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxBasisAmountValue: "99.00"), "application/xml", PdfAssociatedFileRelationship.Data);
        var mismatchedTaxOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCalculatedAmountValue: "20.00"), "application/xml", PdfAssociatedFileRelationship.Data);
        var mismatchedNotSubjectBasisOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(
                headerTradeTaxCategoryCodeValue: "O",
                lineTradeTaxCategoryCodeValue: "O",
                includeTradeTaxRate: false,
                includeLineTradeTaxRate: false,
                lineTotalAmount: "100.00",
                taxBasisTotalAmount: "99.00",
                taxTotalAmount: "0.00",
                grandTotalAmount: "100.00",
                headerTradeTaxBasisAmountValue: "99.00",
                headerTradeTaxCalculatedAmountValue: "0.00",
                headerTradeTaxExemptionReasonValue: "Not subject to VAT"), "application/xml", PdfAssociatedFileRelationship.Data);
        var validNotSubjectBasisOptions = new PdfOptions()
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
        var mismatchedStandardAllowanceBasisOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile(
                "factur-x.xml",
                AddHeaderAllowanceCharge(
                    CreateCiiXml(
                        includeAllowanceTotalAmount: true,
                        allowanceTotalAmount: "10.00",
                        lineTotalAmount: "100.00",
                        taxBasisTotalAmount: "95.00",
                        taxTotalAmount: "21.85",
                        grandTotalAmount: "116.85",
                        headerTradeTaxBasisAmountValue: "95.00",
                        headerTradeTaxCalculatedAmountValue: "21.85"),
                    false,
                    "S",
                    "10.00",
                    true,
                    "23.00"),
                "application/xml",
                PdfAssociatedFileRelationship.Data);
        var validStandardAllowanceBasisOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile(
                "factur-x.xml",
                AddHeaderAllowanceCharge(
                    CreateCiiXml(
                        includeAllowanceTotalAmount: true,
                        allowanceTotalAmount: "10.00",
                        lineTotalAmount: "100.00",
                        taxBasisTotalAmount: "90.00",
                        taxTotalAmount: "20.70",
                        grandTotalAmount: "110.70",
                        headerTradeTaxBasisAmountValue: "90.00",
                        headerTradeTaxCalculatedAmountValue: "20.70"),
                    false,
                    "S",
                    "10.00",
                    true,
                    "23.00"),
                "application/xml",
                PdfAssociatedFileRelationship.Data);
        var validStandardChargeBasisOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile(
                "factur-x.xml",
                AddHeaderAllowanceCharge(
                    CreateCiiXml(
                        includeChargeTotalAmount: true,
                        chargeTotalAmount: "5.00",
                        lineTotalAmount: "100.00",
                        taxBasisTotalAmount: "105.00",
                        taxTotalAmount: "24.15",
                        grandTotalAmount: "129.15",
                        headerTradeTaxBasisAmountValue: "105.00",
                        headerTradeTaxCalculatedAmountValue: "24.15"),
                    true,
                    "S",
                    "5.00",
                    true,
                    "23"),
                "application/xml",
                PdfAssociatedFileRelationship.Data);
        var allowanceChargeAdjustedNotSubjectBasisOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile(
                "factur-x.xml",
                AddHeaderAllowanceCharge(
                    AddHeaderAllowanceCharge(
                        CreateCiiXml(
                            headerTradeTaxCategoryCodeValue: "O",
                            lineTradeTaxCategoryCodeValue: "O",
                            includeTradeTaxRate: false,
                            includeLineTradeTaxRate: false,
                            includeAllowanceTotalAmount: true,
                            includeChargeTotalAmount: true,
                            allowanceTotalAmount: "10.00",
                            chargeTotalAmount: "5.00",
                            lineTotalAmount: "100.00",
                            taxBasisTotalAmount: "95.00",
                            taxTotalAmount: "0.00",
                            grandTotalAmount: "95.00",
                            headerTradeTaxBasisAmountValue: "95.00",
                            headerTradeTaxCalculatedAmountValue: "0.00",
                            headerTradeTaxExemptionReasonValue: "Not subject to VAT"),
                        false,
                        "O",
                        "10.00"),
                    true,
                    "O",
                    "5.00"),
                "application/xml",
                PdfAssociatedFileRelationship.Data);
        var mismatchedChargeOnlyCategoryRateBasisOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile(
                "factur-x.xml",
                AddHeaderAllowanceCharge(
                    AddHeaderTradeTax(
                        CreateCiiXml(
                            includeChargeTotalAmount: true,
                            chargeTotalAmount: "5.00",
                            taxBasisTotalAmount: "104.00",
                            taxTotalAmount: "23.80",
                            grandTotalAmount: "128.80"),
                        "S",
                        "20",
                        "4.00",
                        "0.80"),
                    true,
                    "S",
                    "5.00",
                    true,
                    "20"),
                "application/xml",
                PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement mismatchedBasis = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, mismatchedBasisOptions),
            "einvoice-xml-tax-total-consistency",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement mismatchedTax = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, mismatchedTaxOptions),
            "einvoice-xml-tax-total-consistency",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement mismatchedNotSubjectBasis = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, mismatchedNotSubjectBasisOptions),
            "einvoice-xml-tax-total-consistency",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement validNotSubjectBasis = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, validNotSubjectBasisOptions),
            "einvoice-xml-tax-total-consistency",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement mismatchedStandardAllowanceBasis = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, mismatchedStandardAllowanceBasisOptions),
            "einvoice-xml-tax-total-consistency",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement validStandardAllowanceBasis = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, validStandardAllowanceBasisOptions),
            "einvoice-xml-tax-total-consistency",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement validStandardChargeBasis = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, validStandardChargeBasisOptions),
            "einvoice-xml-tax-total-consistency",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement allowanceChargeAdjustedNotSubjectBasis = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, allowanceChargeAdjustedNotSubjectBasisOptions),
            "einvoice-xml-tax-total-consistency",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement mismatchedChargeOnlyCategoryRateBasis = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, mismatchedChargeOnlyCategoryRateBasisOptions),
            "einvoice-xml-tax-total-consistency",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("BasisAmount sum must match TaxBasisTotalAmount", mismatchedBasis.Diagnostic);
        Assert.Contains("CalculatedAmount sum must match TaxTotalAmount", mismatchedTax.Diagnostic);
        Assert.Contains("Category O ApplicableTradeTax BasisAmount", mismatchedNotSubjectBasis.Diagnostic);
        Assert.Contains("category-O taxable basis", validNotSubjectBasis.Diagnostic);
        Assert.Contains("same category/rate line net amounts", mismatchedStandardAllowanceBasis.Diagnostic);
        Assert.Contains("S/23 expected 90.00", mismatchedStandardAllowanceBasis.Diagnostic);
        Assert.Contains("category/rate adjusted taxable basis", validStandardAllowanceBasis.Diagnostic);
        Assert.Contains("category/rate adjusted taxable basis", validStandardChargeBasis.Diagnostic);
        Assert.Contains("category-O taxable basis", allowanceChargeAdjustedNotSubjectBasis.Diagnostic);
        Assert.Contains("S/20 expected 5.00", mismatchedChargeOnlyCategoryRateBasis.Diagnostic);
        Assert.Contains("line net 0.00", mismatchedChargeOnlyCategoryRateBasis.Diagnostic);
        Assert.Contains("plus charges 5.00", mismatchedChargeOnlyCategoryRateBasis.Diagnostic);
    }


}
