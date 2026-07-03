using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfComplianceAnalyzerTests {
    [Fact]
    public void FacturXReadinessRequiresZeroRatedCiiTaxCategoriesToUseZeroRate() {
        var nonZeroIntracommunityRateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "K", lineTradeTaxCategoryCodeValue: "K"), "application/xml", PdfAssociatedFileRelationship.Data);
        var zeroIntracommunityRateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "K", lineTradeTaxCategoryCodeValue: "K", headerTradeTaxRateValue: "0.00", lineTradeTaxRateValue: "0"), "application/xml", PdfAssociatedFileRelationship.Data);
        var standardRateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "S", lineTradeTaxCategoryCodeValue: "S", headerTradeTaxRateValue: "23", lineTradeTaxRateValue: "23.00"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement nonZeroIntracommunityRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, nonZeroIntracommunityRateOptions),
            "einvoice-xml-tax-category-rate",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement zeroIntracommunityRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, zeroIntracommunityRateOptions),
            "einvoice-xml-tax-category-rate",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement standardRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, standardRateOptions),
            "einvoice-xml-tax-category-rate",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("Peppol requires to be zero", nonZeroIntracommunityRate.Diagnostic);
        Assert.Contains("K/23", nonZeroIntracommunityRate.Diagnostic);
        Assert.Contains("AE, E, G, K, and Z", zeroIntracommunityRate.Diagnostic);
        Assert.Contains("AE, E, G, K, and Z", standardRate.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessAllowsNotSubjectCiiTaxCategoryWithoutRate() {
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
        var notSubjectWithHeaderRateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(
                headerTradeTaxCategoryCodeValue: "O",
                lineTradeTaxCategoryCodeValue: "O",
                includeLineTradeTaxRate: false,
                taxTotalAmount: "0.00",
                grandTotalAmount: "100.00",
                headerTradeTaxCalculatedAmountValue: "0.00",
                headerTradeTaxExemptionReasonValue: "Not subject to VAT"), "application/xml", PdfAssociatedFileRelationship.Data);
        var notSubjectAllowanceWithRateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile(
                "factur-x.xml",
                AddHeaderAllowanceCharge(
                    CreateCiiXml(
                        headerTradeTaxCategoryCodeValue: "O",
                        lineTradeTaxCategoryCodeValue: "O",
                        includeTradeTaxRate: false,
                        includeLineTradeTaxRate: false,
                        includeAllowanceTotalAmount: true,
                        allowanceTotalAmount: "10.00",
                        taxBasisTotalAmount: "90.00",
                        taxTotalAmount: "0.00",
                        grandTotalAmount: "90.00",
                        headerTradeTaxBasisAmountValue: "90.00",
                        headerTradeTaxCalculatedAmountValue: "0.00",
                        headerTradeTaxExemptionReasonValue: "Not subject to VAT"),
                    false,
                    "O",
                    "10.00",
                    true),
                "application/xml",
                PdfAssociatedFileRelationship.Data);
        var notSubjectChargeWithRateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile(
                "factur-x.xml",
                AddHeaderAllowanceCharge(
                    CreateCiiXml(
                        headerTradeTaxCategoryCodeValue: "O",
                        lineTradeTaxCategoryCodeValue: "O",
                        includeTradeTaxRate: false,
                        includeLineTradeTaxRate: false,
                        includeChargeTotalAmount: true,
                        chargeTotalAmount: "5.00",
                        taxBasisTotalAmount: "105.00",
                        taxTotalAmount: "0.00",
                        grandTotalAmount: "105.00",
                        headerTradeTaxBasisAmountValue: "105.00",
                        headerTradeTaxCalculatedAmountValue: "0.00",
                        headerTradeTaxExemptionReasonValue: "Not subject to VAT"),
                    true,
                    "O",
                    "5.00",
                    true),
                "application/xml",
                PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement notSubjectWithoutRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, notSubjectWithoutRateOptions),
            "einvoice-xml-tax-category-rate",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement notSubjectWithHeaderRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, notSubjectWithHeaderRateOptions),
            "einvoice-xml-tax-category-rate",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement notSubjectAllowanceWithRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, notSubjectAllowanceWithRateOptions),
            "einvoice-xml-tax-category-rate",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement notSubjectChargeWithRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, notSubjectChargeWithRateOptions),
            "einvoice-xml-tax-category-rate",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("rate absence for category O", notSubjectWithoutRate.Diagnostic);
        Assert.Contains("Forbidden tax category rate categories: O", notSubjectWithHeaderRate.Diagnostic);
        Assert.Contains("O document-level allowance", notSubjectAllowanceWithRate.Diagnostic);
        Assert.Contains("O document-level charge", notSubjectChargeWithRate.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresZeroRatedCiiTaxCategoriesToUseZeroCalculatedAmount() {
        var nonZeroIntracommunityAmountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "K", lineTradeTaxCategoryCodeValue: "K", headerTradeTaxRateValue: "0", lineTradeTaxRateValue: "0", headerTradeTaxCalculatedAmountValue: "23.45"), "application/xml", PdfAssociatedFileRelationship.Data);
        var zeroIntracommunityAmountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "K", lineTradeTaxCategoryCodeValue: "K", headerTradeTaxRateValue: "0", lineTradeTaxRateValue: "0", headerTradeTaxCalculatedAmountValue: "0.00"), "application/xml", PdfAssociatedFileRelationship.Data);
        var nonZeroNotSubjectAmountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "O", lineTradeTaxCategoryCodeValue: "O", headerTradeTaxCalculatedAmountValue: "1.00", headerTradeTaxExemptionReasonValue: "Not subject to VAT"), "application/xml", PdfAssociatedFileRelationship.Data);
        var zeroNotSubjectAmountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "O", lineTradeTaxCategoryCodeValue: "O", headerTradeTaxCalculatedAmountValue: "0.00", headerTradeTaxExemptionReasonValue: "Not subject to VAT"), "application/xml", PdfAssociatedFileRelationship.Data);
        var standardAmountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "S", lineTradeTaxCategoryCodeValue: "S", headerTradeTaxRateValue: "23", lineTradeTaxRateValue: "23.00", headerTradeTaxCalculatedAmountValue: "23.00"), "application/xml", PdfAssociatedFileRelationship.Data);
        var mismatchedStandardAmountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "S", lineTradeTaxCategoryCodeValue: "S", headerTradeTaxRateValue: "23", lineTradeTaxRateValue: "23.00", headerTradeTaxBasisAmountValue: "100.00", headerTradeTaxCalculatedAmountValue: "20.00", taxTotalAmount: "20.00", grandTotalAmount: "120.00"), "application/xml", PdfAssociatedFileRelationship.Data);
        var tightMismatchedStandardAmountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "S", lineTradeTaxCategoryCodeValue: "S", headerTradeTaxRateValue: "20", lineTradeTaxRateValue: "20.00", headerTradeTaxBasisAmountValue: "100.00", headerTradeTaxCalculatedAmountValue: "20.99", taxTotalAmount: "20.99", grandTotalAmount: "120.99"), "application/xml", PdfAssociatedFileRelationship.Data);
        var calculatedStandardAmountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "S", lineTradeTaxCategoryCodeValue: "S", headerTradeTaxRateValue: "23", lineTradeTaxRateValue: "23.00", headerTradeTaxBasisAmountValue: "100.00", headerTradeTaxCalculatedAmountValue: "23.00", taxTotalAmount: "23.00", grandTotalAmount: "123.00"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement nonZeroIntracommunityAmount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, nonZeroIntracommunityAmountOptions),
            "einvoice-xml-tax-category-amount",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement zeroIntracommunityAmount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, zeroIntracommunityAmountOptions),
            "einvoice-xml-tax-category-amount",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement nonZeroNotSubjectAmount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, nonZeroNotSubjectAmountOptions),
            "einvoice-xml-tax-category-amount",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement zeroNotSubjectAmount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, zeroNotSubjectAmountOptions),
            "einvoice-xml-tax-category-amount",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement standardAmount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, standardAmountOptions),
            "einvoice-xml-tax-category-amount",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement mismatchedStandardAmount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, mismatchedStandardAmountOptions),
            "einvoice-xml-tax-category-amount",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement tightMismatchedStandardAmount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, tightMismatchedStandardAmountOptions),
            "einvoice-xml-tax-category-amount",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement calculatedStandardAmount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, calculatedStandardAmountOptions),
            "einvoice-xml-tax-category-amount",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("Peppol requires to be zero", nonZeroIntracommunityAmount.Diagnostic);
        Assert.Contains("K/23.45", nonZeroIntracommunityAmount.Diagnostic);
        Assert.Contains("Peppol requires to be zero", nonZeroNotSubjectAmount.Diagnostic);
        Assert.Contains("O/1", nonZeroNotSubjectAmount.Diagnostic);
        Assert.Contains("AE, E, G, K, O, and Z", zeroIntracommunityAmount.Diagnostic);
        Assert.Contains("AE, E, G, K, O, and Z", zeroNotSubjectAmount.Diagnostic);
        Assert.Contains("AE, E, G, K, O, and Z", standardAmount.Diagnostic);
        Assert.Contains("taxable basis multiplied by VAT rate", mismatchedStandardAmount.Diagnostic);
        Assert.Contains("S/23 expected 23.00", mismatchedStandardAmount.Diagnostic);
        Assert.Contains("S/20 expected 20.00", tightMismatchedStandardAmount.Diagnostic);
        Assert.Contains("taxable-basis times rate", calculatedStandardAmount.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresParseableCiiTaxCategoryRates() {
        var malformedRateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxRateValue: "not-a-number"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement malformedRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, malformedRateOptions),
            "einvoice-xml-tax-category-rate",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("RateApplicablePercent", malformedRate.Diagnostic);
        Assert.Contains("parseable decimal percentage", malformedRate.Diagnostic);
        Assert.Contains("not-a-number", malformedRate.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresRatesForEveryNonOCategory() {
        var missingRateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile(
                "factur-x.xml",
                CreateCiiXml(
                    includeTradeTaxRate: false,
                    lineTradeTaxCategoryCodeValue: "O",
                    includeLineTradeTaxRate: false),
                "application/xml",
                PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingRateOptions),
            "einvoice-xml-tax-category-rate",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("RateApplicablePercent", missingRate.Diagnostic);
        Assert.Contains("Missing tax category rate", missingRate.Diagnostic);
        Assert.Contains("non-O", missingRate.Diagnostic);
        Assert.Contains("S", missingRate.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresParseableCiiTaxCategoryAmounts() {
        var malformedAmountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCalculatedAmountValue: "not-a-number"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement malformedAmount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, malformedAmountOptions),
            "einvoice-xml-tax-category-amount",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("CalculatedAmount", malformedAmount.Diagnostic);
        Assert.Contains("parseable decimal amount", malformedAmount.Diagnostic);
        Assert.Contains("not-a-number", malformedAmount.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRejectsGroupedCiiDecimalValues() {
        var groupedRateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxRateValue: "1,234.56"), "application/xml", PdfAssociatedFileRelationship.Data);
        var groupedAmountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCalculatedAmountValue: "1,234.56"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement groupedRate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, groupedRateOptions),
            "einvoice-xml-tax-category-rate",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement groupedAmount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, groupedAmountOptions),
            "einvoice-xml-tax-category-amount",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("1,234.56", groupedRate.Diagnostic);
        Assert.Contains("parseable decimal percentage", groupedRate.Diagnostic);
        Assert.Contains("1,234.56", groupedAmount.Diagnostic);
        Assert.Contains("parseable decimal amount", groupedAmount.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiTaxExemptionReasonForRequiredVatCategories() {
        var missingReasonOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "K", lineTradeTaxCategoryCodeValue: "K", headerTradeTaxRateValue: "0", lineTradeTaxRateValue: "0", headerTradeTaxCalculatedAmountValue: "0.00"), "application/xml", PdfAssociatedFileRelationship.Data);
        var reasonTextOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "K", lineTradeTaxCategoryCodeValue: "K", headerTradeTaxRateValue: "0", lineTradeTaxRateValue: "0", headerTradeTaxCalculatedAmountValue: "0.00", headerTradeTaxExemptionReasonValue: "Intra-community supply"), "application/xml", PdfAssociatedFileRelationship.Data);
        var reasonCodeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "E", lineTradeTaxCategoryCodeValue: "E", headerTradeTaxRateValue: "0", lineTradeTaxRateValue: "0", headerTradeTaxCalculatedAmountValue: "0.00", headerTradeTaxExemptionReasonCodeValue: "VATEX-EU-E"), "application/xml", PdfAssociatedFileRelationship.Data);
        var invalidNotSubjectReasonCodeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "O", lineTradeTaxCategoryCodeValue: "O", includeTradeTaxRate: false, includeLineTradeTaxRate: false, taxTotalAmount: "0.00", grandTotalAmount: "100.00", headerTradeTaxCalculatedAmountValue: "0.00", headerTradeTaxExemptionReasonCodeValue: "VATEX-EU-E"), "application/xml", PdfAssociatedFileRelationship.Data);
        var canonicalNotSubjectReasonCodeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "O", lineTradeTaxCategoryCodeValue: "O", includeTradeTaxRate: false, includeLineTradeTaxRate: false, taxTotalAmount: "0.00", grandTotalAmount: "100.00", headerTradeTaxCalculatedAmountValue: "0.00", headerTradeTaxExemptionReasonCodeValue: "VATEX-EU-O"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingReason = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingReasonOptions),
            "einvoice-xml-tax-exemption-reason",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement reasonText = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, reasonTextOptions),
            "einvoice-xml-tax-exemption-reason",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement reasonCode = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, reasonCodeOptions),
            "einvoice-xml-tax-exemption-reason",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement invalidNotSubjectReasonCode = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidNotSubjectReasonCodeOptions),
            "einvoice-xml-tax-exemption-reason",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement canonicalNotSubjectReasonCode = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, canonicalNotSubjectReasonCodeOptions),
            "einvoice-xml-tax-exemption-reason",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("ExemptionReason", missingReason.Diagnostic);
        Assert.Contains("Missing categories: K", missingReason.Diagnostic);
        Assert.Contains("VATEX-EU-O", invalidNotSubjectReasonCode.Diagnostic);
        Assert.Contains("VATEX-EU-E", invalidNotSubjectReasonCode.Diagnostic);
        Assert.Contains("canonical VATEX-EU-O", reasonText.Diagnostic);
        Assert.Contains("canonical VATEX-EU-O", reasonCode.Diagnostic);
        Assert.Contains("canonical VATEX-EU-O", canonicalNotSubjectReasonCode.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRejectsCiiTaxExemptionReasonForForbiddenVatCategories() {
        var forbiddenReasonOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "Z", lineTradeTaxCategoryCodeValue: "Z", headerTradeTaxRateValue: "0", lineTradeTaxRateValue: "0", headerTradeTaxCalculatedAmountValue: "0.00", headerTradeTaxExemptionReasonValue: "Zero rated"), "application/xml", PdfAssociatedFileRelationship.Data);
        var allowedMissingReasonOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "Z", lineTradeTaxCategoryCodeValue: "Z", headerTradeTaxRateValue: "0", lineTradeTaxRateValue: "0", headerTradeTaxCalculatedAmountValue: "0.00"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement forbiddenReason = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, forbiddenReasonOptions),
            "einvoice-xml-tax-exemption-reason",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement allowedMissingReason = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, allowedMissingReasonOptions),
            "einvoice-xml-tax-exemption-reason",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("forbids exemption reasons", forbiddenReason.Diagnostic);
        Assert.Contains("Categories with reason markers: Z", forbiddenReason.Diagnostic);
        Assert.Contains("AE, E, G, K, O, S, Z, L, and M", allowedMissingReason.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresSellerVatIdentifierForExportTaxCategory() {
        var missingSellerVatOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "G", lineTradeTaxCategoryCodeValue: "G", headerTradeTaxRateValue: "0", lineTradeTaxRateValue: "0", headerTradeTaxCalculatedAmountValue: "0.00", headerTradeTaxExemptionReasonValue: "Export outside the EU", includeSellerTaxRegistrationSchemeId: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var validSellerVatOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "G", lineTradeTaxCategoryCodeValue: "G", headerTradeTaxRateValue: "0", lineTradeTaxRateValue: "0", headerTradeTaxCalculatedAmountValue: "0.00", headerTradeTaxExemptionReasonValue: "Export outside the EU"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingSellerVat = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingSellerVatOptions),
            "einvoice-xml-tax-party-identifiers",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement validSellerVat = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, validSellerVatOptions),
            "einvoice-xml-tax-party-identifiers",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("seller VAT/tax identifier", missingSellerVat.Diagnostic);
        Assert.Contains("G", missingSellerVat.Diagnostic);
        Assert.Contains("AE, E, G, K, and O", validSellerVat.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresBuyerVatIdentifierForIntracommunityTaxCategory() {
        var missingBuyerVatOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(headerTradeTaxCategoryCodeValue: "K", lineTradeTaxCategoryCodeValue: "K", headerTradeTaxRateValue: "0", lineTradeTaxRateValue: "0", headerTradeTaxCalculatedAmountValue: "0.00", headerTradeTaxExemptionReasonValue: "Intra-community supply", includeBuyerTaxRegistrationSchemeId: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var standardRateMissingBuyerVatOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeBuyerTaxRegistrationSchemeId: false), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingBuyerVat = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingBuyerVatOptions),
            "einvoice-xml-tax-party-identifiers",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement standardRateMissingBuyerVat = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, standardRateMissingBuyerVatOptions),
            "einvoice-xml-tax-party-identifiers",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("buyer VAT/tax identifier", missingBuyerVat.Diagnostic);
        Assert.Contains("K", missingBuyerVat.Diagnostic);
        Assert.Contains("AE, E, G, K, and O", standardRateMissingBuyerVat.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessForbidsVatIdentifiersForNotSubjectTaxCategory() {
        var notSubjectWithVatIdentifierOptions = new PdfOptions()
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
        var headerOnlyNotSubjectWithVatIdentifierOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(
                headerTradeTaxCategoryCodeValue: "O",
                lineTradeTaxCategoryCodeValue: "S",
                includeTradeTaxRate: false,
                taxTotalAmount: "0.00",
                grandTotalAmount: "100.00",
                headerTradeTaxCalculatedAmountValue: "0.00",
                headerTradeTaxExemptionReasonValue: "Not subject to VAT"), "application/xml", PdfAssociatedFileRelationship.Data);
        var notSubjectWithoutVatIdentifierOptions = new PdfOptions()
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
                headerTradeTaxExemptionReasonValue: "Not subject to VAT",
                includeSellerTaxRegistrationSchemeId: false,
                includeBuyerTaxRegistrationSchemeId: false), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement notSubjectWithVatIdentifier = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, notSubjectWithVatIdentifierOptions),
            "einvoice-xml-tax-party-identifiers",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement headerOnlyNotSubjectWithVatIdentifier = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, headerOnlyNotSubjectWithVatIdentifierOptions),
            "einvoice-xml-tax-party-identifiers",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement notSubjectWithoutVatIdentifier = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, notSubjectWithoutVatIdentifierOptions),
            "einvoice-xml-tax-party-identifiers",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("Peppol category O", notSubjectWithVatIdentifier.Diagnostic);
        Assert.Contains("seller VAT identifier for categories O", notSubjectWithVatIdentifier.Diagnostic);
        Assert.Contains("buyer VAT identifier for categories O", notSubjectWithVatIdentifier.Diagnostic);
        Assert.Contains("seller VAT identifier for categories O", headerOnlyNotSubjectWithVatIdentifier.Diagnostic);
        Assert.Contains("AE, E, G, K, and O", notSubjectWithoutVatIdentifier.Diagnostic);
    }


}
