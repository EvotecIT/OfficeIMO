using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfComplianceAnalyzerTests {
    [Fact]
    public void FacturXReadinessRequiresCiiPaymentTermsEssentials() {
        var missingPaymentTermsOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includePaymentTerms: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingPaymentTermsMarkerOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includePaymentTermsDescription: false, includePaymentTermsDueDate: false), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingPaymentTerms = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingPaymentTermsOptions),
            "einvoice-xml-payment-terms",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingPaymentTermsMarker = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingPaymentTermsMarkerOptions),
            "einvoice-xml-payment-terms",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("SpecifiedTradePaymentTerms", missingPaymentTerms.Diagnostic);
        Assert.Contains("DueDateDateTime or Description", missingPaymentTermsMarker.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiAmountConsistency() {
        var mismatchedLineOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(lineTotalAmount: "99.00"), "application/xml", PdfAssociatedFileRelationship.Data);
        var mismatchedGrandOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(grandTotalAmount: "124.00"), "application/xml", PdfAssociatedFileRelationship.Data);
        var unparseableOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(taxTotalAmount: "not-a-number"), "application/xml", PdfAssociatedFileRelationship.Data);
        var allowanceAdjustedOptions = new PdfOptions()
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
                    "S",
                    "10.00",
                    true,
                    "23"),
                "application/xml",
                PdfAssociatedFileRelationship.Data);
        var mismatchedAllowanceTotalOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile(
                "factur-x.xml",
                AddHeaderAllowanceCharge(
                    CreateCiiXml(
                        includeAllowanceTotalAmount: true,
                        allowanceTotalAmount: "9.00",
                        taxBasisTotalAmount: "91.00",
                        taxTotalAmount: "20.93",
                        grandTotalAmount: "111.93",
                        headerTradeTaxBasisAmountValue: "91.00",
                        headerTradeTaxCalculatedAmountValue: "20.93"),
                    false,
                    "S",
                    "10.00",
                    true,
                    "23"),
                "application/xml",
                PdfAssociatedFileRelationship.Data);
        var mismatchedChargeTotalOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile(
                "factur-x.xml",
                AddHeaderAllowanceCharge(
                    CreateCiiXml(
                        includeChargeTotalAmount: true,
                        chargeTotalAmount: "4.00",
                        taxBasisTotalAmount: "104.00",
                        taxTotalAmount: "23.92",
                        grandTotalAmount: "127.92",
                        headerTradeTaxBasisAmountValue: "104.00",
                        headerTradeTaxCalculatedAmountValue: "23.92"),
                    true,
                    "S",
                    "5.00",
                    true,
                    "23"),
                "application/xml",
                PdfAssociatedFileRelationship.Data);
        var paidAdjustedOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(
                includeDuePayableAmount: true,
                includePaidAmount: true,
                includeRoundingAmount: true,
                paidAmount: "23.00",
                roundingAmount: "0.05",
                duePayableAmount: "100.05"), "application/xml", PdfAssociatedFileRelationship.Data);
        var mismatchedDueOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(
                includeDuePayableAmount: true,
                includePaidAmount: true,
                paidAmount: "23.00",
                duePayableAmount: "99.00"), "application/xml", PdfAssociatedFileRelationship.Data);
        var zeroAllowanceWithoutComponentsOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeAllowanceTotalAmount: true, allowanceTotalAmount: "0.00"), "application/xml", PdfAssociatedFileRelationship.Data);
        var zeroChargeWithoutComponentsOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeChargeTotalAmount: true, chargeTotalAmount: "0.00"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement mismatchedLine = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, mismatchedLineOptions),
            "einvoice-xml-amount-consistency",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement mismatchedGrand = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, mismatchedGrandOptions),
            "einvoice-xml-amount-consistency",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement unparseable = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, unparseableOptions),
            "einvoice-xml-amount-consistency",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement allowanceAdjusted = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, allowanceAdjustedOptions),
            "einvoice-xml-amount-consistency",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement mismatchedAllowanceTotal = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, mismatchedAllowanceTotalOptions),
            "einvoice-xml-amount-consistency",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement mismatchedChargeTotal = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, mismatchedChargeTotalOptions),
            "einvoice-xml-amount-consistency",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement paidAdjusted = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, paidAdjustedOptions),
            "einvoice-xml-amount-consistency",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement mismatchedDue = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, mismatchedDueOptions),
            "einvoice-xml-amount-consistency",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement zeroAllowanceWithoutComponents = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, zeroAllowanceWithoutComponentsOptions),
            "einvoice-xml-amount-consistency",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement zeroChargeWithoutComponents = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, zeroChargeWithoutComponentsOptions),
            "einvoice-xml-amount-consistency",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("LineTotalAmount sum", mismatchedLine.Diagnostic);
        Assert.Contains("TaxBasisTotalAmount plus TaxTotalAmount", mismatchedGrand.Diagnostic);
        Assert.Contains("TaxTotalAmount", unparseable.Diagnostic);
        Assert.Contains("parseable decimal", unparseable.Diagnostic);
        Assert.Contains("allowance", allowanceAdjusted.Diagnostic);
        Assert.Contains("AllowanceTotalAmount", mismatchedAllowanceTotal.Diagnostic);
        Assert.Contains("document-level allowance", mismatchedAllowanceTotal.Diagnostic);
        Assert.Contains("ChargeTotalAmount", mismatchedChargeTotal.Diagnostic);
        Assert.Contains("document-level charge", mismatchedChargeTotal.Diagnostic);
        Assert.Contains("due payable", paidAdjusted.Diagnostic);
        Assert.Contains("DuePayableAmount", mismatchedDue.Diagnostic);
        Assert.Contains("PaidAmount", mismatchedDue.Diagnostic);
        Assert.Contains("document-level allowance", zeroAllowanceWithoutComponents.Diagnostic);
        Assert.Contains("document-level charge", zeroChargeWithoutComponents.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessExcludesHeaderLineTotalAmountFromLineSums() {
        var headerLineTotalOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", AddHeaderLineTotalAmount(CreateCiiXml()), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement headerLineTotal = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, headerLineTotalOptions),
            "einvoice-xml-amount-consistency",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("line", headerLineTotal.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiAllowanceChargeReasons() {
        var missingAllowanceReasonOptions = new PdfOptions()
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
                    "S",
                    "10.00",
                    true,
                    "23",
                    false),
                "application/xml",
                PdfAssociatedFileRelationship.Data);
        var missingChargeReasonOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile(
                "factur-x.xml",
                AddHeaderAllowanceCharge(
                    CreateCiiXml(
                        includeChargeTotalAmount: true,
                        chargeTotalAmount: "5.00",
                        taxBasisTotalAmount: "105.00",
                        taxTotalAmount: "24.15",
                        grandTotalAmount: "129.15",
                        headerTradeTaxBasisAmountValue: "105.00",
                        headerTradeTaxCalculatedAmountValue: "24.15"),
                    true,
                    "S",
                    "5.00",
                    true,
                    "23",
                    false),
                "application/xml",
                PdfAssociatedFileRelationship.Data);
        var reasonedAllowanceChargeOptions = new PdfOptions()
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
                    "S",
                    "10.00",
                    true,
                    "23"),
                "application/xml",
                PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingAllowanceReason = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingAllowanceReasonOptions),
            "einvoice-xml-allowance-charge-reason",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingChargeReason = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingChargeReasonOptions),
            "einvoice-xml-allowance-charge-reason",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement reasonedAllowanceCharge = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, reasonedAllowanceChargeOptions),
            "einvoice-xml-allowance-charge-reason",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("document-level allowance Reason or ReasonCode", missingAllowanceReason.Diagnostic);
        Assert.Contains("ActualAmount 10.00", missingAllowanceReason.Diagnostic);
        Assert.Contains("document-level charge Reason or ReasonCode", missingChargeReason.Diagnostic);
        Assert.Contains("ActualAmount 5.00", missingChargeReason.Diagnostic);
        Assert.Contains("Reason or ReasonCode", reasonedAllowanceCharge.Diagnostic);
    }


}
