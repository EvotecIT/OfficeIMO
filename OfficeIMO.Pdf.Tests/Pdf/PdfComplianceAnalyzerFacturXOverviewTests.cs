using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfComplianceAnalyzerTests {
    [Fact]
    public void FacturXReadinessRecognizesXmlDataAttachmentAndReportsEinvoiceGaps() {
        byte[] invoiceXml = CreateCiiXml();
        var options = new PdfOptions {
                FileVersion = PdfFileVersion.Pdf17,
                IncludeStandardFontToUnicodeMaps = true
            }
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .AddFacturXInvoiceXml(invoiceXml);

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, options);

        Assert.False(report.IsReady);
        AssertRequirement(report, "pdf-file-version", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "pdfa-identification", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-attachment", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-profile-context", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-document-header", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-document-type-code", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-date-format", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-trade-transaction", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-party-identification", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-country-code", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-electronic-address", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-party-tax-registration", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-party-tax-registration-scheme", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-line-item", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-unit-code", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-line-pricing", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-line-amount-consistency", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-line-tax", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-settlement-summary", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-currency-consistency", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-currency-code", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-tax-breakdown", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-tax-category-code", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-tax-category-rate", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-tax-category-amount", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-tax-exemption-reason", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-tax-party-identifiers", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-tax-category-consistency", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-tax-total-consistency", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-payment-instructions", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-payment-means-code", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-payment-account-format", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-payment-terms", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-amount-consistency", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-allowance-charge-reason", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-attachment-params", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xmp-extension", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "output-intent-policy", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "mustang-validation", PdfComplianceRequirementStatus.Unsupported);
    }

    [Fact]
    public void FacturXGroundworkHelperSatisfiesConfiguredEinvoiceReadinessWithoutEnablingProfile() {
        var options = new PdfOptions()
            .ConfigureFacturXGroundwork(CreateCiiXml());

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, options);

        Assert.Equal(PdfComplianceProfile.None, options.ComplianceProfile);
        Assert.False(report.IsReady);
        AssertRequirement(report, "pdf-file-version", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "xmp-metadata", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "pdfa-identification", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "output-intent", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "output-intent-policy", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "standard-font-to-unicode", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-attachment", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-profile-context", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-document-header", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-document-type-code", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-date-format", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-trade-transaction", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-party-identification", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-country-code", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-electronic-address", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-party-tax-registration", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-party-tax-registration-scheme", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-line-item", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-unit-code", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-line-pricing", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-line-amount-consistency", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-line-tax", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-settlement-summary", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-currency-consistency", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-currency-code", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-tax-breakdown", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-tax-category-code", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-tax-category-rate", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-tax-category-amount", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-tax-exemption-reason", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-tax-party-identifiers", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-tax-category-consistency", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-tax-total-consistency", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-payment-instructions", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-payment-means-code", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-payment-account-format", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-payment-terms", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-amount-consistency", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-allowance-charge-reason", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xml-attachment-params", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "einvoice-xmp-extension", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "mustang-validation", PdfComplianceRequirementStatus.Unsupported);
    }

    [Fact]
    public void FacturXReadinessReportsMissingOrMismatchedEinvoiceXmpMetadata() {
        var missingOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(), "application/xml", PdfAssociatedFileRelationship.Data);
        var mismatchedOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(new PdfElectronicInvoiceMetadata("ORDER", "invoice.xml", "1.0", "EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missing = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingOptions),
            "einvoice-xmp-extension",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement mismatched = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, mismatchedOptions),
            "einvoice-xmp-extension",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("Set PdfOptions.SetElectronicInvoiceMetadata", missing.Diagnostic);
        Assert.Contains("INVOICE", mismatched.Diagnostic);
        Assert.Contains("factur-x.xml", mismatched.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresKnownEinvoiceXmpConformanceLevel() {
        var invalidOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("CUSTOM"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(), "application/xml", PdfAssociatedFileRelationship.Data);
        var basicWlOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("BASIC_WL"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement invalid = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidOptions),
            "einvoice-xmp-extension",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement basicWl = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, basicWlOptions),
            "einvoice-xmp-extension",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("conformance level", invalid.Diagnostic);
        Assert.Contains("EN 16931", invalid.Diagnostic);
        Assert.Contains("canonical", basicWl.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresKnownCiiProfileContext() {
        var missingContextOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(null), "application/xml", PdfAssociatedFileRelationship.Data);
        var unknownContextOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml("urn:example:custom-profile"), "application/xml", PdfAssociatedFileRelationship.Data);
        var substringContextOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml("not-en16931-test"), "application/xml", PdfAssociatedFileRelationship.Data);
        var xRechnungOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("XRECHNUNG"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml("urn:cen.eu:en16931:2017#compliant#urn:xoev-de:kosit:standard:xrechnung_3.0"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missing = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingContextOptions),
            "einvoice-xml-profile-context",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement unknown = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, unknownContextOptions),
            "einvoice-xml-profile-context",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement substring = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, substringContextOptions),
            "einvoice-xml-profile-context",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement xRechnung = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, xRechnungOptions),
            "einvoice-xml-profile-context",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("GuidelineSpecifiedDocumentContextParameter", missing.Diagnostic);
        Assert.Contains("recognized", unknown.Diagnostic);
        Assert.Contains("custom-profile", unknown.Diagnostic);
        Assert.Contains("not-en16931-test", substring.Diagnostic);
        Assert.Contains("recognized EN 16931", xRechnung.Diagnostic);
    }


}
