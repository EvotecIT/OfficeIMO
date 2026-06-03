using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfComplianceAnalyzerTests {
    [Fact]
    public void PdfA3BReadinessReportsSatisfiedGroundworkAndUnsupportedProfileGates() {
        var options = new PdfOptions {
                FileVersion = PdfFileVersion.Pdf17,
                IncludeStandardFontToUnicodeMaps = true
            }
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent();

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.PdfA3B, options);

        Assert.Equal(PdfComplianceProfile.PdfA3B, report.Profile);
        Assert.Equal("PDF/A-3b", report.DisplayName);
        Assert.False(report.IsReady);
        AssertRequirement(report, "pdf-file-version", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "xmp-metadata", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "pdfa-identification", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "output-intent", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "output-intent-policy", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "verapdf-validation", PdfComplianceRequirementStatus.Unsupported);
        Assert.Empty(report.MissingRequirements);
        Assert.Contains(report.UnsupportedRequirements, requirement => requirement.Id == "verapdf-validation");
    }

    [Fact]
    public void PdfA3BReadinessSeparatesOutputIntentPresenceFromApprovedPolicy() {
        var genericOptions = new PdfOptions {
                FileVersion = PdfFileVersion.Pdf17
            }
            .SetPdfAIdentification(3, "B")
            .SetOutputIntent(CreateMinimalIccProfile(), "OfficeIMO RGB");
        var policyOptions = new PdfOptions {
                FileVersion = PdfFileVersion.Pdf17
            }
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent();

        PdfComplianceReadinessReport genericReport = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.PdfA3B, genericOptions);
        PdfComplianceReadinessReport policyReport = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.PdfA3B, policyOptions);

        PdfComplianceRequirement presence = AssertRequirement(genericReport, "output-intent", PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement missingPolicy = AssertRequirement(genericReport, "output-intent-policy", PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement satisfiedPolicy = AssertRequirement(policyReport, "output-intent-policy", PdfComplianceRequirementStatus.Satisfied);
        Assert.Contains("parseable", presence.Diagnostic);
        Assert.Contains("known profile policy", missingPolicy.Diagnostic);
        Assert.Contains("veraPDF", satisfiedPolicy.Diagnostic);
    }

    [Fact]
    public void PdfA3BReadinessRejectsMismatchedOutputIntentPolicy() {
        var cmykOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetOutputIntent(CreateMinimalIccProfile("CMYK"), policy: PdfOutputIntentPolicy.SrgbIec6196621);
        var identifierOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetOutputIntent(CreateMinimalIccProfile(), "OfficeIMO RGB", PdfOutputIntentPolicy.SrgbIec6196621);

        PdfComplianceRequirement cmyk = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.PdfA3B, cmykOptions),
            "output-intent-policy",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement identifier = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.PdfA3B, identifierOptions),
            "output-intent-policy",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("RGB ICC", cmyk.Diagnostic);
        Assert.Contains("OutputConditionIdentifier", identifier.Diagnostic);
    }

    [Fact]
    public void PdfA3BReadinessReportsMissingPdf17FileVersionGroundwork() {
        var options = new PdfOptions {
                IncludeStandardFontToUnicodeMaps = true
            }
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent();

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.PdfA3B, options);

        PdfComplianceRequirement requirement = AssertRequirement(report, "pdf-file-version", PdfComplianceRequirementStatus.Missing);
        Assert.Contains(nameof(PdfFileVersion.Pdf17), requirement.Diagnostic);
    }

    [Fact]
    public void PdfA3BReadinessReportsMissingGeneratedStandardFontEmbedding() {
        var options = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent();

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(
            PdfComplianceProfile.PdfA3B,
            options,
            new[] { PdfStandardFont.Helvetica, PdfStandardFont.HelveticaBold });

        Assert.False(report.IsReady);
        PdfComplianceRequirement requirement = AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("Helvetica", requirement.Diagnostic);
        Assert.Contains("Helvetica-Bold", requirement.Diagnostic);
    }

    [Fact]
    public void PdfA3BReadinessAcceptsEmbeddedMappingsForGeneratedStandardFonts() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath), "HelveticaAudit");

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(
            PdfComplianceProfile.PdfA3B,
            options,
            new[] { PdfStandardFont.Helvetica });

        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Satisfied);
    }

    [Fact]
    public void PdfA3BReadinessRejectsInvalidEmbeddedMappingsForGeneratedStandardFonts() {
        var options = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .EmbedStandardFont(PdfStandardFont.Helvetica, new byte[] { 1 }, "HelveticaAudit");

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(
            PdfComplianceProfile.PdfA3B,
            options,
            new[] { PdfStandardFont.Helvetica });

        PdfComplianceRequirement requirement = AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("invalid embedded TrueType", requirement.Diagnostic);
        Assert.Contains("Helvetica", requirement.Diagnostic);
    }

    [Fact]
    public void PdfA3BReadinessTreatsNoGeneratedStandardFontsAsSatisfied() {
        var options = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent();

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(
            PdfComplianceProfile.PdfA3B,
            options,
            Array.Empty<PdfStandardFont>());

        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Satisfied);
    }

    [Fact]
    public void PdfA3UReadinessReportsWrongIdentificationAndMissingUnicodeMaps() {
        var options = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetOutputIntent(CreateMinimalIccProfile(), "OfficeIMO RGB");

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.PdfA3U, options);

        Assert.False(report.IsReady);
        AssertRequirement(report, "pdfa-identification", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "standard-font-to-unicode", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "full-unicode-mapping", PdfComplianceRequirementStatus.Unsupported);
        Assert.Contains("PDF/A-3u", report.FindRequirement("pdfa-identification")!.Diagnostic);
    }

    [Fact]
    public void PdfA3AReadinessKeepsPdfUaSpecificChecksOutOfPdfAAccessibility() {
        var options = new PdfOptions {
                FileVersion = PdfFileVersion.Pdf17,
                Language = "en-US",
                IncludeStandardFontToUnicodeMaps = true
            }
            .SetPdfAIdentification(3, "A")
            .SetSrgbOutputIntent();

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.PdfA3A, options);

        AssertRequirement(report, "pdf-file-version", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "pdfa-identification", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "standard-font-to-unicode", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "document-language", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "tagged-structure", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "tagged-page-tab-order", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "tagged-parent-tree-next-key", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-document-structure-root", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-document-structure-language", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-text-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-list-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-list-structure-containers", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-table-cell-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-table-structure-containers", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-table-header-scope-attributes", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-table-span-attributes", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-table-caption-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-link-annotation-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-link-text-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-form-widget-structure-references", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "generated-form-field-accessible-names", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "generated-image-alternate-text", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "generated-image-structure-references", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "generated-drawing-alternate-text", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "generated-drawing-structure-references", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "decorative-drawing-artifacts", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "decorative-running-page-text-artifacts", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "decorative-flow-rule-artifacts", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "decorative-layout-artifacts", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "alternate-text", PdfComplianceRequirementStatus.Unsupported);
        Assert.Null(report.FindRequirement("pdfua-identification"));
        Assert.Null(report.FindRequirement("document-title"));
        Assert.Null(report.FindRequirement("display-document-title"));
    }

    [Fact]
    public void PdfAGroundworkHelperSatisfiesConfiguredPdfAAccessibilityReadinessWithoutEnablingProfile() {
        var options = new PdfOptions()
            .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA3A, "en-US");

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.PdfA3A, options);

        Assert.Equal(PdfComplianceProfile.None, options.ComplianceProfile);
        Assert.False(report.IsReady);
        AssertRequirement(report, "pdf-file-version", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "xmp-metadata", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "pdfa-identification", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "output-intent", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "output-intent-policy", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "standard-font-to-unicode", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "document-language", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "tagged-catalog-markers", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "tagged-page-tab-order", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-document-structure-root", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-document-structure-language", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "tagged-structure", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "verapdf-validation", PdfComplianceRequirementStatus.Unsupported);
        Assert.Null(report.FindRequirement("pdfua-identification"));
        Assert.Null(report.FindRequirement("display-document-title"));
    }

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

    [Fact]
    public void FacturXReadinessRequiresCiiDocumentHeaderEssentials() {
        var missingHeaderOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeDocumentHeader: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingTradeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeSupplyChainTradeTransaction: false), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingHeader = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingHeaderOptions),
            "einvoice-xml-document-header",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingTrade = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingTradeOptions),
            "einvoice-xml-document-header",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("ExchangedDocument ID", missingHeader.Diagnostic);
        Assert.Contains("ExchangedDocument TypeCode", missingHeader.Diagnostic);
        Assert.Contains("ExchangedDocument IssueDateTime", missingHeader.Diagnostic);
        Assert.Contains("SupplyChainTradeTransaction", missingTrade.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiDocumentTypeCodeListValue() {
        var invalidTypeCodeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(documentTypeCodeValue: "999"), "application/xml", PdfAssociatedFileRelationship.Data);
        var creditNoteTypeCodeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(documentTypeCodeValue: "381"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement invalidTypeCode = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidTypeCodeOptions),
            "einvoice-xml-document-type-code",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement creditNoteTypeCode = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, creditNoteTypeCodeOptions),
            "einvoice-xml-document-type-code",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("ExchangedDocument TypeCode", invalidTypeCode.Diagnostic);
        Assert.Contains("UNTDID 1001", invalidTypeCode.Diagnostic);
        Assert.Contains("999", invalidTypeCode.Diagnostic);
        Assert.Contains("UNTDID 1001", creditNoteTypeCode.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiDateFormats() {
        var invalidIssueDateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(issueDateTimeValue: "20261340"), "application/xml", PdfAssociatedFileRelationship.Data);
        var invalidDueDateOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(dueDateTimeValue: "20260230"), "application/xml", PdfAssociatedFileRelationship.Data);
        var validDateTimeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(issueDateTimeFormat: "203", issueDateTimeValue: "202606031430"), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingIssueFormatOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(issueDateTimeFormat: ""), "application/xml", PdfAssociatedFileRelationship.Data);
        var unknownDueFormatOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(dueDateTimeFormat: "999"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement invalidIssueDate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidIssueDateOptions),
            "einvoice-xml-date-format",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement invalidDueDate = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidDueDateOptions),
            "einvoice-xml-date-format",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement validDateTime = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, validDateTimeOptions),
            "einvoice-xml-date-format",
            PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement missingIssueFormat = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingIssueFormatOptions),
            "einvoice-xml-date-format",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement unknownDueFormat = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, unknownDueFormatOptions),
            "einvoice-xml-date-format",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("ExchangedDocument IssueDateTime", invalidIssueDate.Diagnostic);
        Assert.Contains("SpecifiedTradePaymentTerms DueDateDateTime", invalidDueDate.Diagnostic);
        Assert.Contains("DateTimeString", invalidDueDate.Diagnostic);
        Assert.Contains("ExchangedDocument IssueDateTime", missingIssueFormat.Diagnostic);
        Assert.Contains("SpecifiedTradePaymentTerms DueDateDateTime", unknownDueFormat.Diagnostic);
        Assert.Contains("issue", validDateTime.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiTradeTransactionEssentials() {
        var emptyTransactionOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeTradeTransactionEssentials: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingSellerOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeSellerTradeParty: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingAmountOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includePayableAmount: false), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement emptyTransaction = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, emptyTransactionOptions),
            "einvoice-xml-trade-transaction",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingSeller = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingSellerOptions),
            "einvoice-xml-trade-transaction",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingAmount = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingAmountOptions),
            "einvoice-xml-trade-transaction",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("ApplicableHeaderTradeAgreement", emptyTransaction.Diagnostic);
        Assert.Contains("SellerTradeParty", emptyTransaction.Diagnostic);
        Assert.Contains("BuyerTradeParty", emptyTransaction.Diagnostic);
        Assert.Contains("ApplicableHeaderTradeSettlement", emptyTransaction.Diagnostic);
        Assert.Contains("SpecifiedTradeSettlementHeaderMonetarySummation", emptyTransaction.Diagnostic);
        Assert.Contains("SellerTradeParty", missingSeller.Diagnostic);
        Assert.Contains("GrandTotalAmount or DuePayableAmount", missingAmount.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiPartyIdentificationEssentials() {
        var missingSellerCountryOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeSellerCountryId: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingBuyerNameOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeBuyerName: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingBuyerCountryOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeBuyerCountryId: false), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingSellerCountry = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingSellerCountryOptions),
            "einvoice-xml-party-identification",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingBuyerName = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingBuyerNameOptions),
            "einvoice-xml-party-identification",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingBuyerCountry = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingBuyerCountryOptions),
            "einvoice-xml-party-identification",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("SellerTradeParty PostalTradeAddress CountryID", missingSellerCountry.Diagnostic);
        Assert.Contains("BuyerTradeParty Name", missingBuyerName.Diagnostic);
        Assert.Contains("BuyerTradeParty PostalTradeAddress CountryID", missingBuyerCountry.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiCountryCodeListValue() {
        var invalidSellerCountryOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(sellerCountryIdValue: "POL"), "application/xml", PdfAssociatedFileRelationship.Data);
        var invalidBuyerCountryOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(buyerCountryIdValue: "D"), "application/xml", PdfAssociatedFileRelationship.Data);
        var swissCountryOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(sellerCountryIdValue: "CH", buyerCountryIdValue: "LI"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement invalidSellerCountry = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidSellerCountryOptions),
            "einvoice-xml-country-code",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement invalidBuyerCountry = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidBuyerCountryOptions),
            "einvoice-xml-country-code",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement swissCountry = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, swissCountryOptions),
            "einvoice-xml-country-code",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("ISO 3166-1 alpha-2", invalidSellerCountry.Diagnostic);
        Assert.Contains("POL", invalidSellerCountry.Diagnostic);
        Assert.Contains("D", invalidBuyerCountry.Diagnostic);
        Assert.Contains("ISO 3166-1 alpha-2", swissCountry.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiElectronicAddressSchemeListValue() {
        var missingSellerAddressOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeSellerElectronicAddress: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingBuyerSchemeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeBuyerElectronicAddressSchemeId: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var invalidSchemeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(sellerElectronicAddressSchemeIdValue: "9999"), "application/xml", PdfAssociatedFileRelationship.Data);
        var leitwegIdOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(buyerElectronicAddressSchemeIdValue: "0204", buyerElectronicAddressValue: "991-12345-XX"), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingSellerAddress = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingSellerAddressOptions),
            "einvoice-xml-electronic-address",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingBuyerScheme = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingBuyerSchemeOptions),
            "einvoice-xml-electronic-address",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement invalidScheme = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, invalidSchemeOptions),
            "einvoice-xml-electronic-address",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement leitwegId = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, leitwegIdOptions),
            "einvoice-xml-electronic-address",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("SellerTradeParty URIUniversalCommunication", missingSellerAddress.Diagnostic);
        Assert.Contains("BuyerTradeParty URIUniversalCommunication URIID schemeID", missingBuyerScheme.Diagnostic);
        Assert.Contains("Electronic Address Scheme", invalidScheme.Diagnostic);
        Assert.Contains("9999", invalidScheme.Diagnostic);
        Assert.Contains("Electronic Address Scheme", leitwegId.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiPartyTaxRegistrationEssentials() {
        var missingSellerTaxOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeSellerTaxRegistration: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingBuyerTaxOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeBuyerTaxRegistration: false), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingSellerTax = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingSellerTaxOptions),
            "einvoice-xml-party-tax-registration",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingBuyerTax = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingBuyerTaxOptions),
            "einvoice-xml-party-tax-registration",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("SellerTradeParty SpecifiedTaxRegistration ID", missingSellerTax.Diagnostic);
        Assert.Contains("category-specific", missingBuyerTax.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRequiresCiiPartyTaxRegistrationSchemeMetadata() {
        var missingSellerSchemeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeSellerTaxRegistrationSchemeId: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var missingBuyerSchemeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(includeBuyerTaxRegistrationSchemeId: false), "application/xml", PdfAssociatedFileRelationship.Data);
        var validSchemeOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .SetElectronicInvoiceMetadata(PdfElectronicInvoiceMetadata.FacturX("EN 16931"))
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceRequirement missingSellerScheme = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingSellerSchemeOptions),
            "einvoice-xml-party-tax-registration-scheme",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement missingBuyerScheme = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, missingBuyerSchemeOptions),
            "einvoice-xml-party-tax-registration-scheme",
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement validScheme = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, validSchemeOptions),
            "einvoice-xml-party-tax-registration-scheme",
            PdfComplianceRequirementStatus.Satisfied);

        Assert.Contains("SellerTradeParty SpecifiedTaxRegistration ID schemeID", missingSellerScheme.Diagnostic);
        Assert.Contains("BuyerTradeParty SpecifiedTaxRegistration ID schemeID", missingBuyerScheme.Diagnostic);
        Assert.Contains("schemeID metadata", validScheme.Diagnostic);
    }

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

        Assert.Contains("SpecifiedLineTradeAgreement", missingAgreement.Diagnostic);
        Assert.Contains("NetPriceProductTradePrice", missingAgreement.Diagnostic);
        Assert.Contains("ChargeAmount", missingPrice.Diagnostic);
        Assert.Contains("NetPriceProductTradePrice", grossOnlyPrice.Diagnostic);
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

        Assert.Contains("ApplicableTradeTax", missingLineTax.Diagnostic);
        Assert.Contains("ApplicableTradeTax CategoryCode", missingLineTaxCategory.Diagnostic);
        Assert.Contains("ApplicableTradeTax RateApplicablePercent", missingLineTaxRate.Diagnostic);
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
            PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement mismatchedAmountCurrency = AssertRequirement(
            PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, mismatchedAmountCurrencyOptions),
            "einvoice-xml-currency-consistency",
            PdfComplianceRequirementStatus.Missing);

        Assert.Contains("InvoiceCurrencyCode", missingInvoiceCurrency.Diagnostic);
        Assert.Contains("currencyID", missingAmountCurrency.Diagnostic);
        Assert.Contains("LineTotalAmount", missingAmountCurrency.Diagnostic);
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

        Assert.Contains("SpecifiedTradeSettlementPaymentMeans", missingPaymentMeans.Diagnostic);
        Assert.Contains("SpecifiedTradeSettlementPaymentMeans TypeCode", missingPaymentType.Diagnostic);
        Assert.Contains("PayeePartyCreditorFinancialAccount IBANID or ProprietaryID", missingAccount.Diagnostic);
        Assert.Contains("does not require creditor account identifiers", cashWithoutAccount.Diagnostic);
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

        Assert.Contains("SpecifiedTradeSettlementPaymentMeans TypeCode", invalidTypeCode.Diagnostic);
        Assert.Contains("UNCL4461", invalidTypeCode.Diagnostic);
        Assert.Contains("999", invalidTypeCode.Diagnostic);
        Assert.Contains("UNCL4461", creditTransfer.Diagnostic);
        Assert.Contains("UNCL4461", mutuallyDefined.Diagnostic);
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

    [Fact]
    public void FacturXReadinessRejectsNonCanonicalXmlAttachmentName() {
        var options = new PdfOptions {
                IncludeStandardFontToUnicodeMaps = true
            }
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .AddEmbeddedFile("invoice.xml", CreateCiiXml(), "application/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, options);

        PdfComplianceRequirement requirement = AssertRequirement(report, "einvoice-xml-attachment", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "einvoice-xml-attachment-params", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("factur-x.xml", requirement.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRejectsSupplementRelationshipForInvoiceXml() {
        var options = new PdfOptions {
                IncludeStandardFontToUnicodeMaps = true
            }
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .AddEmbeddedFile("factur-x.xml", CreateCiiXml(), "application/xml", PdfAssociatedFileRelationship.Supplement);

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, options);

        PdfComplianceRequirement requirement = AssertRequirement(report, "einvoice-xml-attachment", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("AFRelationship", requirement.Diagnostic);
        Assert.Contains("Alternative, Data, or Source", requirement.Diagnostic);
    }

    [Fact]
    public void FacturXReadinessRejectsMalformedOrWrongRootXmlAttachment() {
        var malformedOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .AddEmbeddedFile("factur-x.xml", Encoding.UTF8.GetBytes("<rsm:CrossIndustryInvoice />"), "application/xml", PdfAssociatedFileRelationship.Data);
        var wrongRootOptions = new PdfOptions()
            .SetPdfAIdentification(3, "B")
            .SetSrgbOutputIntent()
            .AddEmbeddedFile("factur-x.xml", Encoding.UTF8.GetBytes("<Invoice />"), "text/xml", PdfAssociatedFileRelationship.Data);

        PdfComplianceReadinessReport malformedReport = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, malformedOptions);
        PdfComplianceReadinessReport wrongRootReport = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.FacturX, wrongRootOptions);

        Assert.Contains("parseable XML", AssertRequirement(malformedReport, "einvoice-xml-attachment", PdfComplianceRequirementStatus.Missing).Diagnostic);
        Assert.Contains("CrossIndustryInvoice", AssertRequirement(wrongRootReport, "einvoice-xml-attachment", PdfComplianceRequirementStatus.Missing).Diagnostic);
    }

    [Fact]
    public void PdfUaReadinessReportsLanguageAndAccessibilityGaps() {
        var options = new PdfOptions {
            FileVersion = PdfFileVersion.Pdf17,
            Language = "en-US",
            IncludeStandardFontToUnicodeMaps = true
        }.SetPdfUaIdentification();

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.PdfUa1, options);

        Assert.False(report.IsReady);
        Assert.Equal("PDF/UA-1", report.DisplayName);
        AssertRequirement(report, "pdf-file-version", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "standard-font-to-unicode", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "pdfua-identification", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "document-title", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "display-document-title", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "document-language", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "tagged-catalog-markers", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "tagged-structure", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "tagged-page-tab-order", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "tagged-parent-tree-next-key", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-document-structure-root", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-document-structure-language", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-text-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-list-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-list-structure-containers", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-table-cell-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-table-structure-containers", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-table-header-scope-attributes", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-table-span-attributes", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-table-caption-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-link-annotation-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-link-text-structure-references", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "generated-form-widget-structure-references", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "generated-form-field-accessible-names", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "generated-image-structure-references", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "generated-drawing-alternate-text", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "generated-drawing-structure-references", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "decorative-drawing-artifacts", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "decorative-running-page-text-artifacts", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "decorative-flow-rule-artifacts", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "decorative-layout-artifacts", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "alternate-text", PdfComplianceRequirementStatus.Unsupported);
    }

    [Theory]
    [InlineData("English")]
    [InlineData("en_US")]
    public void PdfUaReadinessRejectsInvalidLanguageTags(string language) {
        var options = new PdfOptions {
            FileVersion = PdfFileVersion.Pdf17,
            Language = language,
            IncludeStandardFontToUnicodeMaps = true
        }
            .SetPdfUaIdentification()
            .EnableTaggedPdfCatalogMarkers();

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.PdfUa1, options);

        PdfComplianceRequirement documentLanguage = AssertRequirement(report, "document-language", PdfComplianceRequirementStatus.Missing);
        PdfComplianceRequirement structureLanguage = AssertRequirement(report, "generated-document-structure-language", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("valid language tag", documentLanguage.Diagnostic);
        Assert.Contains("valid language tag", structureLanguage.Diagnostic);
    }

    [Fact]
    public void PdfUaReadinessRecognizesTaggedCatalogMarkersWithoutClaimingFullStructure() {
        var options = new PdfOptions {
            FileVersion = PdfFileVersion.Pdf17,
            Language = "en-US",
            IncludeStandardFontToUnicodeMaps = true
        }
            .SetPdfUaIdentification()
            .EnableTaggedPdfCatalogMarkers();

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.PdfUa1, options);

        Assert.False(report.IsReady);
        PdfComplianceRequirement markers = AssertRequirement(report, "tagged-catalog-markers", PdfComplianceRequirementStatus.Satisfied);
        PdfComplianceRequirement structure = AssertRequirement(report, "tagged-structure", PdfComplianceRequirementStatus.Unsupported);
        Assert.Contains("/MarkInfo", markers.Diagnostic);
        Assert.Contains("complete marked-content reference coverage", structure.Diagnostic);
        AssertRequirement(report, "tagged-page-tab-order", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "tagged-parent-tree-next-key", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-document-structure-root", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-document-structure-language", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-text-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-list-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-list-structure-containers", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-table-cell-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-table-structure-containers", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-table-header-scope-attributes", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-table-span-attributes", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-table-caption-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-link-annotation-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-link-text-structure-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-form-widget-structure-references", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "generated-form-field-accessible-names", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "generated-image-structure-references", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "generated-drawing-alternate-text", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "generated-drawing-structure-references", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "decorative-drawing-artifacts", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "decorative-running-page-text-artifacts", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "decorative-flow-rule-artifacts", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "decorative-layout-artifacts", PdfComplianceRequirementStatus.Satisfied);
    }

    [Fact]
    public void PdfUaGroundworkHelperSatisfiesConfigurableAccessibilityReadinessWithoutEnablingProfile() {
        var options = new PdfOptions()
            .ConfigurePdfUaGroundwork("en-US");

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.PdfUa1, options);

        Assert.Equal(PdfComplianceProfile.None, options.ComplianceProfile);
        Assert.False(report.IsReady);
        AssertRequirement(report, "pdf-file-version", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "standard-font-to-unicode", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "pdfua-identification", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "display-document-title", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "document-language", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "tagged-catalog-markers", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "tagged-page-tab-order", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "tagged-parent-tree-next-key", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-document-structure-root", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "generated-document-structure-language", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "document-title", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "tagged-structure", PdfComplianceRequirementStatus.Unsupported);
        AssertRequirement(report, "full-unicode-mapping", PdfComplianceRequirementStatus.Unsupported);
    }

    [Fact]
    public void PdfUaReadinessReportsViewerDisplayTitlePreference() {
        var options = new PdfOptions {
            Language = "en-US",
            IncludeStandardFontToUnicodeMaps = true,
            ViewerPreferences = new PdfViewerPreferencesOptions {
                DisplayDocTitle = true
            }
        }.SetPdfUaIdentification();

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.PdfUa1, options);

        AssertRequirement(report, "display-document-title", PdfComplianceRequirementStatus.Satisfied);
    }

    [Fact]
    public void PdfUaReadinessReportsMissingIdentification() {
        var options = new PdfOptions {
            Language = "en-US",
            IncludeStandardFontToUnicodeMaps = true
        };

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.PdfUa1, options);

        PdfComplianceRequirement requirement = AssertRequirement(report, "pdfua-identification", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("Set PdfOptions.SetPdfUaIdentification", requirement.Diagnostic);
    }

    [Fact]
    public void OptionsOverloadUsesRequestedProfileAndNoneHasNoRequirements() {
        var options = new PdfOptions {
            ComplianceProfile = PdfComplianceProfile.PdfA2B
        };

        PdfComplianceReadinessReport requestedReport = PdfComplianceAnalyzer.Assess(options);
        PdfComplianceReadinessReport noneReport = PdfComplianceAnalyzer.Assess(PdfComplianceProfile.None, new PdfOptions());

        Assert.Equal(PdfComplianceProfile.PdfA2B, requestedReport.Profile);
        AssertRequirement(requestedReport, "pdfa-identification", PdfComplianceRequirementStatus.Missing);
        Assert.Equal(PdfComplianceProfile.None, noneReport.Profile);
        Assert.True(noneReport.IsReady);
        Assert.Empty(noneReport.Requirements);
        Assert.Null(noneReport.FindRequirement("pdfa-identification"));
    }

    private static PdfComplianceRequirement AssertRequirement(PdfComplianceReadinessReport report, string id, PdfComplianceRequirementStatus status) {
        PdfComplianceRequirement requirement = Assert.Single(report.Requirements, requirement => requirement.Id == id);
        Assert.True(requirement.Status == status, "Requirement " + id + " expected " + status + " but was " + requirement.Status + ": " + requirement.Diagnostic);
        Assert.False(string.IsNullOrWhiteSpace(requirement.DisplayName));
        Assert.False(string.IsNullOrWhiteSpace(requirement.Diagnostic));
        return requirement;
    }

    private static byte[] AddHeaderTradeTax(byte[] ciiXml, string categoryCode, bool includeRate) {
        string xml = Encoding.UTF8.GetString(ciiXml);
        string tradeTax = CreateApplicableTradeTax(
            true,
            true,
            true,
            includeRate,
            true,
            true,
            categoryCode,
            "VAT",
            includeRate ? "23" : "0",
            "0.00",
            "0.00",
            string.Equals(categoryCode, "O", StringComparison.Ordinal) ? "Not subject to VAT" : null,
            null,
            "EUR");
        return Encoding.UTF8.GetBytes(xml.Replace("</ram:ApplicableHeaderTradeSettlement>", tradeTax + "</ram:ApplicableHeaderTradeSettlement>"));
    }

    private static byte[] AddHeaderTradeTax(byte[] ciiXml, string categoryCode, string rateValue, string basisAmount, string calculatedAmount) {
        string xml = Encoding.UTF8.GetString(ciiXml);
        string tradeTax = CreateApplicableTradeTax(
            true,
            true,
            true,
            true,
            true,
            true,
            categoryCode,
            "VAT",
            rateValue,
            basisAmount,
            calculatedAmount,
            null,
            null,
            "EUR");
        return Encoding.UTF8.GetBytes(xml.Replace("</ram:ApplicableHeaderTradeSettlement>", tradeTax + "</ram:ApplicableHeaderTradeSettlement>"));
    }

    private static byte[] AddHeaderAllowanceCharge(byte[] ciiXml, bool charge, string categoryCode, string actualAmount, bool includeRate = false, string rateValue = "0", bool includeReason = true) {
        string xml = Encoding.UTF8.GetString(ciiXml);
        string rate = includeRate
            ? "<ram:RateApplicablePercent>" + rateValue + "</ram:RateApplicablePercent>"
            : string.Empty;
        string reason = includeReason
            ? "<ram:Reason>" + (charge ? "Service charge" : "Document allowance") + "</ram:Reason>"
            : string.Empty;
        string allowanceCharge =
            "<ram:SpecifiedTradeAllowanceCharge>" +
            "<ram:ChargeIndicator><udt:Indicator>" + (charge ? "true" : "false") + "</udt:Indicator></ram:ChargeIndicator>" +
            "<ram:ActualAmount currencyID=\"EUR\">" + actualAmount + "</ram:ActualAmount>" +
            reason +
            "<ram:CategoryTradeTax>" +
            "<ram:TypeCode>VAT</ram:TypeCode>" +
            "<ram:CategoryCode>" + categoryCode + "</ram:CategoryCode>" +
            rate +
            "</ram:CategoryTradeTax>" +
            "</ram:SpecifiedTradeAllowanceCharge>";
        return Encoding.UTF8.GetBytes(xml.Replace("<ram:SpecifiedTradeSettlementHeaderMonetarySummation>", allowanceCharge + "<ram:SpecifiedTradeSettlementHeaderMonetarySummation>"));
    }

    private static byte[] CreateGrossLinePriceCiiXml() {
        string xml = Encoding.UTF8.GetString(CreateCiiXml());
        xml = xml.Replace("NetPriceProductTradePrice", "GrossPriceProductTradePrice");
        return Encoding.UTF8.GetBytes(xml);
    }

    private static byte[] CreateCiiXml(
        string? profileContextId = "urn:factur-x.eu:1p0:en16931",
        bool includeDocumentHeader = true,
        bool includeSupplyChainTradeTransaction = true,
        bool includeTradeTransactionEssentials = true,
        bool includeSellerTradeParty = true,
        bool includeBuyerTradeParty = true,
        bool includeSellerName = true,
        bool includeSellerCountryId = true,
        bool includeSellerTaxRegistration = true,
        bool includeSellerElectronicAddress = true,
        bool includeSellerElectronicAddressSchemeId = true,
        bool includeBuyerName = true,
        bool includeBuyerCountryId = true,
        bool includeBuyerTaxRegistration = true,
        bool includeSellerTaxRegistrationSchemeId = true,
        bool includeBuyerTaxRegistrationSchemeId = true,
        bool includeBuyerElectronicAddress = true,
        bool includeBuyerElectronicAddressSchemeId = true,
        bool includePayableAmount = true,
        bool includeLineItem = true,
        bool includeLineItemProductName = true,
        bool includeLineTradeAgreement = true,
        bool includeLinePriceChargeAmount = true,
        bool includeLineTradeTax = true,
        bool includeLineTradeTaxTypeCode = true,
        bool includeLineTradeTaxCategoryCode = true,
        bool includeLineTradeTaxRate = true,
        bool includeLineTotalAmount = true,
        bool includeLineBilledQuantityUnitCode = true,
        bool includeInvoiceCurrencyCode = true,
        bool includeApplicableTradeTax = true,
        bool includeTradeTaxTypeCode = true,
        bool includeTradeTaxCategoryCode = true,
        bool includeTradeTaxRate = true,
        bool includeTradeTaxBasisAmount = true,
        bool includeTradeTaxCalculatedAmount = true,
        bool includeTaxTotals = true,
        bool includeAllowanceTotalAmount = false,
        bool includeChargeTotalAmount = false,
        bool includeDuePayableAmount = false,
        bool includePaidAmount = false,
        bool includeRoundingAmount = false,
        bool includePaymentMeans = true,
        bool includePaymentMeansTypeCode = true,
        bool includeCreditorAccount = true,
        bool includeCreditorAccountId = true,
        bool useCreditorProprietaryAccountId = false,
        string paymentMeansTypeCodeValue = "58",
        bool includePaymentTerms = true,
        bool includePaymentTermsDescription = true,
        bool includePaymentTermsDueDate = true,
        string documentTypeCodeValue = "380",
        string issueDateTimeFormat = "102",
        string issueDateTimeValue = "20260603",
        string dueDateTimeFormat = "102",
        string dueDateTimeValue = "20260703",
        string lineTotalAmount = "100.00",
        string linePriceChargeAmountValue = "100.00",
        string? linePriceBasisQuantityValue = null,
        string lineBilledQuantityValue = "1",
        string lineBilledQuantityUnitCodeValue = "C62",
        string taxBasisTotalAmount = "100.00",
        string taxTotalAmount = "23.00",
        string grandTotalAmount = "123.00",
        string allowanceTotalAmount = "0.00",
        string chargeTotalAmount = "0.00",
        string duePayableAmount = "123.00",
        string paidAmount = "0.00",
        string roundingAmount = "0.00",
        string headerTradeTaxCategoryCodeValue = "S",
        string headerTradeTaxTypeCodeValue = "VAT",
        string headerTradeTaxRateValue = "23",
        string headerTradeTaxBasisAmountValue = "100.00",
        string headerTradeTaxCalculatedAmountValue = "23.00",
        string? headerTradeTaxExemptionReasonValue = null,
        string? headerTradeTaxExemptionReasonCodeValue = null,
        string lineTradeTaxTypeCodeValue = "VAT",
        string lineTradeTaxCategoryCodeValue = "S",
        string lineTradeTaxRateValue = "23",
        string invoiceCurrencyCodeValue = "EUR",
        string? amountCurrencyId = "EUR",
        string creditorAccountIban = "PL61109010140000071219812874",
        string creditorProprietaryAccountId = "ACCOUNT-001",
        string sellerCountryIdValue = "PL",
        string buyerCountryIdValue = "DE",
        string sellerElectronicAddressValue = "PL1234567890",
        string sellerElectronicAddressSchemeIdValue = "9945",
        string buyerElectronicAddressValue = "DE123456789",
        string buyerElectronicAddressSchemeIdValue = "9930") {
        string context = profileContextId == null
            ? "<rsm:ExchangedDocumentContext />"
            : "<rsm:ExchangedDocumentContext>" +
              "<ram:GuidelineSpecifiedDocumentContextParameter>" +
              "<ram:ID>" + profileContextId + "</ram:ID>" +
              "</ram:GuidelineSpecifiedDocumentContextParameter>" +
              "</rsm:ExchangedDocumentContext>";
        string document = includeDocumentHeader
            ? "<rsm:ExchangedDocument>" +
              "<ram:ID>INV-2026-0001</ram:ID>" +
              "<ram:TypeCode>" + documentTypeCodeValue + "</ram:TypeCode>" +
              "<ram:IssueDateTime><udt:DateTimeString format=\"" + issueDateTimeFormat + "\">" + issueDateTimeValue + "</udt:DateTimeString></ram:IssueDateTime>" +
              "</rsm:ExchangedDocument>"
            : "<rsm:ExchangedDocument />";
        string transaction = CreateSupplyChainTradeTransaction(
            includeSupplyChainTradeTransaction,
            includeTradeTransactionEssentials,
            includeSellerTradeParty,
            includeBuyerTradeParty,
            includeSellerName,
            includeSellerCountryId,
            includeSellerTaxRegistration,
            includeSellerTaxRegistrationSchemeId,
            includeSellerElectronicAddress,
            includeSellerElectronicAddressSchemeId,
            includeBuyerName,
            includeBuyerCountryId,
            includeBuyerTaxRegistration,
            includeBuyerTaxRegistrationSchemeId,
            includeBuyerElectronicAddress,
            includeBuyerElectronicAddressSchemeId,
            includePayableAmount,
            includeLineItem,
            includeLineItemProductName,
            includeLineTradeAgreement,
            includeLinePriceChargeAmount,
            includeLineTradeTax,
            includeLineTradeTaxTypeCode,
            includeLineTradeTaxCategoryCode,
            includeLineTradeTaxRate,
            includeLineTotalAmount,
            includeLineBilledQuantityUnitCode,
            includeInvoiceCurrencyCode,
            includeApplicableTradeTax,
            includeTradeTaxTypeCode,
            includeTradeTaxCategoryCode,
            includeTradeTaxRate,
            includeTradeTaxBasisAmount,
            includeTradeTaxCalculatedAmount,
            includeTaxTotals,
            includeAllowanceTotalAmount,
            includeChargeTotalAmount,
            includeDuePayableAmount,
            includePaidAmount,
            includeRoundingAmount,
            includePaymentMeans,
            includePaymentMeansTypeCode,
            paymentMeansTypeCodeValue,
            includeCreditorAccount,
            includeCreditorAccountId,
            useCreditorProprietaryAccountId,
            includePaymentTerms,
            includePaymentTermsDescription,
            includePaymentTermsDueDate,
            dueDateTimeFormat,
            dueDateTimeValue,
            lineTotalAmount,
            linePriceChargeAmountValue,
            linePriceBasisQuantityValue,
            lineBilledQuantityValue,
            lineBilledQuantityUnitCodeValue,
            taxBasisTotalAmount,
            taxTotalAmount,
            grandTotalAmount,
            allowanceTotalAmount,
            chargeTotalAmount,
            duePayableAmount,
            paidAmount,
            roundingAmount,
            headerTradeTaxCategoryCodeValue,
            headerTradeTaxTypeCodeValue,
            headerTradeTaxRateValue,
            headerTradeTaxBasisAmountValue,
            headerTradeTaxCalculatedAmountValue,
            headerTradeTaxExemptionReasonValue,
            headerTradeTaxExemptionReasonCodeValue,
            lineTradeTaxTypeCodeValue,
            lineTradeTaxCategoryCodeValue,
            lineTradeTaxRateValue,
            invoiceCurrencyCodeValue,
            amountCurrencyId,
            creditorAccountIban,
            creditorProprietaryAccountId,
            sellerCountryIdValue,
            buyerCountryIdValue,
            sellerElectronicAddressValue,
            sellerElectronicAddressSchemeIdValue,
            buyerElectronicAddressValue,
            buyerElectronicAddressSchemeIdValue);
        return Encoding.UTF8.GetBytes(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<rsm:CrossIndustryInvoice xmlns:rsm=\"urn:un:unece:uncefact:data:standard:CrossIndustryInvoice:100\" xmlns:ram=\"urn:un:unece:uncefact:data:standard:ReusableAggregateBusinessInformationEntity:100\" xmlns:udt=\"urn:un:unece:uncefact:data:standard:UnqualifiedDataType:100\">" +
            context +
            document +
            transaction +
            "</rsm:CrossIndustryInvoice>");
    }

    private static string CreateSupplyChainTradeTransaction(
        bool includeSupplyChainTradeTransaction,
        bool includeTradeTransactionEssentials,
        bool includeSellerTradeParty,
        bool includeBuyerTradeParty,
        bool includeSellerName,
        bool includeSellerCountryId,
        bool includeSellerTaxRegistration,
        bool includeSellerTaxRegistrationSchemeId,
        bool includeSellerElectronicAddress,
        bool includeSellerElectronicAddressSchemeId,
        bool includeBuyerName,
        bool includeBuyerCountryId,
        bool includeBuyerTaxRegistration,
        bool includeBuyerTaxRegistrationSchemeId,
        bool includeBuyerElectronicAddress,
        bool includeBuyerElectronicAddressSchemeId,
        bool includePayableAmount,
        bool includeLineItem,
        bool includeLineItemProductName,
        bool includeLineTradeAgreement,
        bool includeLinePriceChargeAmount,
        bool includeLineTradeTax,
        bool includeLineTradeTaxTypeCode,
        bool includeLineTradeTaxCategoryCode,
        bool includeLineTradeTaxRate,
        bool includeLineTotalAmount,
        bool includeLineBilledQuantityUnitCode,
        bool includeInvoiceCurrencyCode,
        bool includeApplicableTradeTax,
        bool includeTradeTaxTypeCode,
        bool includeTradeTaxCategoryCode,
        bool includeTradeTaxRate,
        bool includeTradeTaxBasisAmount,
        bool includeTradeTaxCalculatedAmount,
        bool includeTaxTotals,
        bool includeAllowanceTotalAmount,
        bool includeChargeTotalAmount,
        bool includeDuePayableAmount,
        bool includePaidAmount,
        bool includeRoundingAmount,
        bool includePaymentMeans,
        bool includePaymentMeansTypeCode,
        string paymentMeansTypeCodeValue,
        bool includeCreditorAccount,
        bool includeCreditorAccountId,
        bool useCreditorProprietaryAccountId,
        bool includePaymentTerms,
        bool includePaymentTermsDescription,
        bool includePaymentTermsDueDate,
        string dueDateTimeFormat,
        string dueDateTimeValue,
        string lineTotalAmountValue,
        string linePriceChargeAmountValue,
        string? linePriceBasisQuantityValue,
        string lineBilledQuantityValue,
        string lineBilledQuantityUnitCodeValue,
        string taxBasisTotalAmountValue,
        string taxTotalAmountValue,
        string grandTotalAmountValue,
        string allowanceTotalAmountValue,
        string chargeTotalAmountValue,
        string duePayableAmountValue,
        string paidAmountValue,
        string roundingAmountValue,
        string headerTradeTaxCategoryCodeValue,
        string headerTradeTaxTypeCodeValue,
        string headerTradeTaxRateValue,
        string headerTradeTaxBasisAmountValue,
        string headerTradeTaxCalculatedAmountValue,
        string? headerTradeTaxExemptionReasonValue,
        string? headerTradeTaxExemptionReasonCodeValue,
        string lineTradeTaxTypeCodeValue,
        string lineTradeTaxCategoryCodeValue,
        string lineTradeTaxRateValue,
        string invoiceCurrencyCodeValue,
        string? amountCurrencyId,
        string creditorAccountIban,
        string creditorProprietaryAccountId,
        string sellerCountryIdValue,
        string buyerCountryIdValue,
        string sellerElectronicAddressValue,
        string sellerElectronicAddressSchemeIdValue,
        string buyerElectronicAddressValue,
        string buyerElectronicAddressSchemeIdValue) {
        if (!includeSupplyChainTradeTransaction) {
            return string.Empty;
        }

        if (!includeTradeTransactionEssentials) {
            return "<rsm:SupplyChainTradeTransaction />";
        }

        string seller = includeSellerTradeParty
            ? CreateTradeParty("SellerTradeParty", "OfficeIMO Seller", sellerCountryIdValue, "PL1234567890", includeSellerName, includeSellerCountryId, includeSellerTaxRegistration, includeSellerTaxRegistrationSchemeId, includeSellerElectronicAddress, includeSellerElectronicAddressSchemeId, sellerElectronicAddressValue, sellerElectronicAddressSchemeIdValue)
            : string.Empty;
        string buyer = includeBuyerTradeParty
            ? CreateTradeParty("BuyerTradeParty", "OfficeIMO Buyer", buyerCountryIdValue, "DE123456789", includeBuyerName, includeBuyerCountryId, includeBuyerTaxRegistration, includeBuyerTaxRegistrationSchemeId, includeBuyerElectronicAddress, includeBuyerElectronicAddressSchemeId, buyerElectronicAddressValue, buyerElectronicAddressSchemeIdValue)
            : string.Empty;
        string amount = includePayableAmount
            ? "<ram:GrandTotalAmount" + CurrencyAttribute(amountCurrencyId) + ">" + grandTotalAmountValue + "</ram:GrandTotalAmount>"
            : string.Empty;
        string duePayable = includeDuePayableAmount
            ? "<ram:DuePayableAmount" + CurrencyAttribute(amountCurrencyId) + ">" + duePayableAmountValue + "</ram:DuePayableAmount>"
            : string.Empty;
        string paid = includePaidAmount
            ? "<ram:PaidAmount" + CurrencyAttribute(amountCurrencyId) + ">" + paidAmountValue + "</ram:PaidAmount>"
            : string.Empty;
        string rounding = includeRoundingAmount
            ? "<ram:RoundingAmount" + CurrencyAttribute(amountCurrencyId) + ">" + roundingAmountValue + "</ram:RoundingAmount>"
            : string.Empty;
        string lineItem = CreateIncludedSupplyChainTradeLineItem(includeLineItem, includeLineItemProductName, includeLineTradeAgreement, includeLinePriceChargeAmount, includeLineTradeTax, includeLineTradeTaxTypeCode, includeLineTradeTaxCategoryCode, includeLineTradeTaxRate, includeLineTotalAmount, includeLineBilledQuantityUnitCode, lineTotalAmountValue, linePriceChargeAmountValue, linePriceBasisQuantityValue, lineBilledQuantityValue, lineBilledQuantityUnitCodeValue, lineTradeTaxTypeCodeValue, lineTradeTaxCategoryCodeValue, lineTradeTaxRateValue, amountCurrencyId);
        string currencyCode = includeInvoiceCurrencyCode
            ? "<ram:InvoiceCurrencyCode>" + invoiceCurrencyCodeValue + "</ram:InvoiceCurrencyCode>"
            : string.Empty;
        string tradeTax = CreateApplicableTradeTax(
            includeApplicableTradeTax,
            includeTradeTaxTypeCode,
            includeTradeTaxCategoryCode,
            includeTradeTaxRate,
            includeTradeTaxBasisAmount,
            includeTradeTaxCalculatedAmount,
            headerTradeTaxCategoryCodeValue,
            headerTradeTaxTypeCodeValue,
            headerTradeTaxRateValue,
            headerTradeTaxBasisAmountValue,
            headerTradeTaxCalculatedAmountValue,
            headerTradeTaxExemptionReasonValue,
            headerTradeTaxExemptionReasonCodeValue,
            amountCurrencyId);
        string taxTotals = includeTaxTotals
            ? "<ram:TaxBasisTotalAmount" + CurrencyAttribute(amountCurrencyId) + ">" + taxBasisTotalAmountValue + "</ram:TaxBasisTotalAmount>" +
              "<ram:TaxTotalAmount" + CurrencyAttribute(amountCurrencyId) + ">" + taxTotalAmountValue + "</ram:TaxTotalAmount>"
            : string.Empty;
        string allowanceTotal = includeAllowanceTotalAmount
            ? "<ram:AllowanceTotalAmount" + CurrencyAttribute(amountCurrencyId) + ">" + allowanceTotalAmountValue + "</ram:AllowanceTotalAmount>"
            : string.Empty;
        string chargeTotal = includeChargeTotalAmount
            ? "<ram:ChargeTotalAmount" + CurrencyAttribute(amountCurrencyId) + ">" + chargeTotalAmountValue + "</ram:ChargeTotalAmount>"
            : string.Empty;
        string paymentMeans = CreatePaymentMeans(includePaymentMeans, includePaymentMeansTypeCode, paymentMeansTypeCodeValue, includeCreditorAccount, includeCreditorAccountId, useCreditorProprietaryAccountId, creditorAccountIban, creditorProprietaryAccountId);
        string paymentTerms = CreatePaymentTerms(includePaymentTerms, includePaymentTermsDescription, includePaymentTermsDueDate, dueDateTimeFormat, dueDateTimeValue);
        return "<rsm:SupplyChainTradeTransaction>" +
               lineItem +
               "<ram:ApplicableHeaderTradeAgreement>" +
               seller +
               buyer +
               "</ram:ApplicableHeaderTradeAgreement>" +
               "<ram:ApplicableHeaderTradeSettlement>" +
                currencyCode +
                tradeTax +
                paymentMeans +
                paymentTerms +
                "<ram:SpecifiedTradeSettlementHeaderMonetarySummation>" +
               taxTotals +
               allowanceTotal +
               chargeTotal +
               amount +
               paid +
               rounding +
               duePayable +
               "</ram:SpecifiedTradeSettlementHeaderMonetarySummation>" +
               "</ram:ApplicableHeaderTradeSettlement>" +
               "</rsm:SupplyChainTradeTransaction>";
    }

    private static string CreatePaymentMeans(
        bool includePaymentMeans,
        bool includePaymentMeansTypeCode,
        string paymentMeansTypeCodeValue,
        bool includeCreditorAccount,
        bool includeCreditorAccountId,
        bool useCreditorProprietaryAccountId,
        string creditorAccountIban,
        string creditorProprietaryAccountId) {
        if (!includePaymentMeans) {
            return string.Empty;
        }

        string typeCode = includePaymentMeansTypeCode
            ? "<ram:TypeCode>" + paymentMeansTypeCodeValue + "</ram:TypeCode>"
            : string.Empty;
        string accountId = includeCreditorAccountId
            ? useCreditorProprietaryAccountId
                ? "<ram:ProprietaryID>" + creditorProprietaryAccountId + "</ram:ProprietaryID>"
                : "<ram:IBANID>" + creditorAccountIban + "</ram:IBANID>"
            : string.Empty;
        string account = includeCreditorAccount
            ? "<ram:PayeePartyCreditorFinancialAccount>" +
              accountId +
              "</ram:PayeePartyCreditorFinancialAccount>"
            : string.Empty;
        return "<ram:SpecifiedTradeSettlementPaymentMeans>" +
               typeCode +
               account +
               "</ram:SpecifiedTradeSettlementPaymentMeans>";
    }

    private static string CreatePaymentTerms(bool includePaymentTerms, bool includePaymentTermsDescription, bool includePaymentTermsDueDate, string dueDateTimeFormat, string dueDateTimeValue) {
        if (!includePaymentTerms) {
            return string.Empty;
        }

        string description = includePaymentTermsDescription
            ? "<ram:Description>Due within 30 days</ram:Description>"
            : string.Empty;
        string dueDate = includePaymentTermsDueDate
            ? "<ram:DueDateDateTime><udt:DateTimeString format=\"" + dueDateTimeFormat + "\">" + dueDateTimeValue + "</udt:DateTimeString></ram:DueDateDateTime>"
            : string.Empty;

        return "<ram:SpecifiedTradePaymentTerms>" +
               description +
               dueDate +
               "</ram:SpecifiedTradePaymentTerms>";
    }

    private static byte[] CreateTwoLineCiiXmlWithSecondLineMissingProductName() {
        string firstLine = CreateIncludedSupplyChainTradeLineItem(
            includeLineItem: true,
            includeLineItemProductName: true,
            includeLineTradeAgreement: true,
            includeLinePriceChargeAmount: true,
            includeLineTradeTax: true,
            includeLineTradeTaxTypeCode: true,
            includeLineTradeTaxCategoryCode: true,
            includeLineTradeTaxRate: true,
            includeLineTotalAmount: true,
            includeLineBilledQuantityUnitCode: true,
            lineTotalAmountValue: "100.00",
            linePriceChargeAmountValue: "100.00",
            linePriceBasisQuantityValue: null,
            lineBilledQuantityValue: "1",
            lineBilledQuantityUnitCodeValue: "C62",
            lineTradeTaxTypeCodeValue: "VAT",
            lineTradeTaxCategoryCodeValue: "S",
            lineTradeTaxRateValue: "23",
            amountCurrencyId: "EUR");
        string secondLine = CreateIncludedSupplyChainTradeLineItem(
            includeLineItem: true,
            includeLineItemProductName: false,
            includeLineTradeAgreement: true,
            includeLinePriceChargeAmount: true,
            includeLineTradeTax: true,
            includeLineTradeTaxTypeCode: true,
            includeLineTradeTaxCategoryCode: true,
            includeLineTradeTaxRate: true,
            includeLineTotalAmount: true,
            includeLineBilledQuantityUnitCode: true,
            lineTotalAmountValue: "100.00",
            linePriceChargeAmountValue: "100.00",
            linePriceBasisQuantityValue: null,
            lineBilledQuantityValue: "1",
            lineBilledQuantityUnitCodeValue: "C62",
            lineTradeTaxTypeCodeValue: "VAT",
            lineTradeTaxCategoryCodeValue: "S",
            lineTradeTaxRateValue: "23",
            amountCurrencyId: "EUR")
            .Replace("<ram:LineID>1</ram:LineID>", "<ram:LineID>2</ram:LineID>");
        string xml = Encoding.UTF8.GetString(CreateCiiXml())
            .Replace(firstLine, firstLine + secondLine);
        return Encoding.UTF8.GetBytes(xml);
    }

    private static byte[] CreateCiiXmlWithLineAllowance() {
        string xml = Encoding.UTF8.GetString(CreateCiiXml(
            lineTotalAmount: "95.00",
            linePriceChargeAmountValue: "10.00",
            lineBilledQuantityValue: "10",
            taxBasisTotalAmount: "95.00",
            taxTotalAmount: "21.85",
            grandTotalAmount: "116.85",
            headerTradeTaxBasisAmountValue: "95.00",
            headerTradeTaxCalculatedAmountValue: "21.85"));
        string allowance =
            "<ram:SpecifiedTradeAllowanceCharge>" +
            "<ram:ChargeIndicator><udt:Indicator>false</udt:Indicator></ram:ChargeIndicator>" +
            "<ram:ActualAmount currencyID=\"EUR\">5.00</ram:ActualAmount>" +
            "</ram:SpecifiedTradeAllowanceCharge>";
        return Encoding.UTF8.GetBytes(xml.Replace("<ram:SpecifiedTradeSettlementLineMonetarySummation>", allowance + "<ram:SpecifiedTradeSettlementLineMonetarySummation>"));
    }

    private static string CreateApplicableTradeTax(
        bool includeApplicableTradeTax,
        bool includeTradeTaxTypeCode,
        bool includeTradeTaxCategoryCode,
        bool includeTradeTaxRate,
        bool includeTradeTaxBasisAmount,
        bool includeTradeTaxCalculatedAmount,
        string headerTradeTaxCategoryCodeValue,
        string headerTradeTaxTypeCodeValue,
        string headerTradeTaxRateValue,
        string headerTradeTaxBasisAmountValue,
        string headerTradeTaxCalculatedAmountValue,
        string? headerTradeTaxExemptionReasonValue,
        string? headerTradeTaxExemptionReasonCodeValue,
        string? amountCurrencyId) {
        if (!includeApplicableTradeTax) {
            return string.Empty;
        }

        string calculatedAmount = includeTradeTaxCalculatedAmount
            ? "<ram:CalculatedAmount" + CurrencyAttribute(amountCurrencyId) + ">" + headerTradeTaxCalculatedAmountValue + "</ram:CalculatedAmount>"
            : string.Empty;
        string typeCode = includeTradeTaxTypeCode
            ? "<ram:TypeCode>" + headerTradeTaxTypeCodeValue + "</ram:TypeCode>"
            : string.Empty;
        string basisAmount = includeTradeTaxBasisAmount
            ? "<ram:BasisAmount" + CurrencyAttribute(amountCurrencyId) + ">" + headerTradeTaxBasisAmountValue + "</ram:BasisAmount>"
            : string.Empty;
        string categoryCode = includeTradeTaxCategoryCode
            ? "<ram:CategoryCode>" + headerTradeTaxCategoryCodeValue + "</ram:CategoryCode>"
            : string.Empty;
        string rate = includeTradeTaxRate
            ? "<ram:RateApplicablePercent>" + headerTradeTaxRateValue + "</ram:RateApplicablePercent>"
            : string.Empty;
        string exemptionReason = headerTradeTaxExemptionReasonValue == null
            ? string.Empty
            : "<ram:ExemptionReason>" + headerTradeTaxExemptionReasonValue + "</ram:ExemptionReason>";
        string exemptionReasonCode = headerTradeTaxExemptionReasonCodeValue == null
            ? string.Empty
            : "<ram:ExemptionReasonCode>" + headerTradeTaxExemptionReasonCodeValue + "</ram:ExemptionReasonCode>";

        return "<ram:ApplicableTradeTax>" +
               calculatedAmount +
               typeCode +
               basisAmount +
               categoryCode +
               rate +
               exemptionReason +
               exemptionReasonCode +
               "</ram:ApplicableTradeTax>";
    }

    private static string CreateTradeParty(
        string elementName,
        string nameValue,
        string countryId,
        string taxId,
        bool includeName,
        bool includeCountryId,
        bool includeTaxRegistration,
        bool includeTaxRegistrationSchemeId,
        bool includeElectronicAddress,
        bool includeElectronicAddressSchemeId,
        string electronicAddressValue,
        string electronicAddressSchemeIdValue) {
        string name = includeName
            ? "<ram:Name>" + nameValue + "</ram:Name>"
            : string.Empty;
        string taxRegistrationSchemeId = includeTaxRegistrationSchemeId
            ? " schemeID=\"VA\""
            : string.Empty;
        string taxRegistration = includeTaxRegistration
            ? "<ram:SpecifiedTaxRegistration><ram:ID" + taxRegistrationSchemeId + ">" + taxId + "</ram:ID></ram:SpecifiedTaxRegistration>"
            : string.Empty;
        string electronicAddressSchemeId = includeElectronicAddressSchemeId
            ? " schemeID=\"" + electronicAddressSchemeIdValue + "\""
            : string.Empty;
        string electronicAddress = includeElectronicAddress
            ? "<ram:URIUniversalCommunication><ram:URIID" + electronicAddressSchemeId + ">" + electronicAddressValue + "</ram:URIID></ram:URIUniversalCommunication>"
            : string.Empty;
        string country = includeCountryId
            ? "<ram:CountryID>" + countryId + "</ram:CountryID>"
            : string.Empty;
        string address = "<ram:PostalTradeAddress>" +
                         "<ram:PostcodeCode>00-001</ram:PostcodeCode>" +
                         "<ram:LineOne>Compliance Street 1</ram:LineOne>" +
                         "<ram:CityName>Warsaw</ram:CityName>" +
                         country +
                         "</ram:PostalTradeAddress>";
        return "<ram:" + elementName + ">" +
               name +
               taxRegistration +
               electronicAddress +
               address +
               "</ram:" + elementName + ">";
    }

    private static string CreateIncludedSupplyChainTradeLineItem(
        bool includeLineItem,
        bool includeLineItemProductName,
        bool includeLineTradeAgreement,
        bool includeLinePriceChargeAmount,
        bool includeLineTradeTax,
        bool includeLineTradeTaxTypeCode,
        bool includeLineTradeTaxCategoryCode,
        bool includeLineTradeTaxRate,
        bool includeLineTotalAmount,
        bool includeLineBilledQuantityUnitCode,
        string lineTotalAmountValue,
        string linePriceChargeAmountValue,
        string? linePriceBasisQuantityValue,
        string lineBilledQuantityValue,
        string lineBilledQuantityUnitCodeValue,
        string lineTradeTaxTypeCodeValue,
        string lineTradeTaxCategoryCodeValue,
        string lineTradeTaxRateValue,
        string? amountCurrencyId) {
        if (!includeLineItem) {
            return string.Empty;
        }

        string productName = includeLineItemProductName
            ? "<ram:Name>OfficeIMO PDF compliance work</ram:Name>"
            : string.Empty;
        string linePriceChargeAmount = includeLinePriceChargeAmount
            ? "<ram:ChargeAmount" + CurrencyAttribute(amountCurrencyId) + ">" + linePriceChargeAmountValue + "</ram:ChargeAmount>"
            : string.Empty;
        string linePriceBasisQuantity = linePriceBasisQuantityValue == null
            ? string.Empty
            : "<ram:BasisQuantity>" + linePriceBasisQuantityValue + "</ram:BasisQuantity>";
        string lineTradeAgreement = includeLineTradeAgreement
            ? "<ram:SpecifiedLineTradeAgreement>" +
              "<ram:NetPriceProductTradePrice>" +
              linePriceChargeAmount +
              linePriceBasisQuantity +
              "</ram:NetPriceProductTradePrice>" +
              "</ram:SpecifiedLineTradeAgreement>"
            : string.Empty;
        string lineTotalAmount = includeLineTotalAmount
            ? "<ram:LineTotalAmount" + CurrencyAttribute(amountCurrencyId) + ">" + lineTotalAmountValue + "</ram:LineTotalAmount>"
            : string.Empty;
        string lineTradeTax = CreateLineApplicableTradeTax(includeLineTradeTax, includeLineTradeTaxTypeCode, includeLineTradeTaxCategoryCode, includeLineTradeTaxRate, lineTradeTaxTypeCodeValue, lineTradeTaxCategoryCodeValue, lineTradeTaxRateValue);
        string billedQuantityUnitCode = includeLineBilledQuantityUnitCode
            ? " unitCode=\"" + lineBilledQuantityUnitCodeValue + "\""
            : string.Empty;
        return "<ram:IncludedSupplyChainTradeLineItem>" +
               "<ram:AssociatedDocumentLineDocument><ram:LineID>1</ram:LineID></ram:AssociatedDocumentLineDocument>" +
               "<ram:SpecifiedTradeProduct>" + productName + "</ram:SpecifiedTradeProduct>" +
               lineTradeAgreement +
               "<ram:SpecifiedLineTradeDelivery><ram:BilledQuantity" + billedQuantityUnitCode + ">" + lineBilledQuantityValue + "</ram:BilledQuantity></ram:SpecifiedLineTradeDelivery>" +
               "<ram:SpecifiedLineTradeSettlement>" +
               lineTradeTax +
               "<ram:SpecifiedTradeSettlementLineMonetarySummation>" +
               lineTotalAmount +
               "</ram:SpecifiedTradeSettlementLineMonetarySummation>" +
               "</ram:SpecifiedLineTradeSettlement>" +
               "</ram:IncludedSupplyChainTradeLineItem>";
    }

    private static string CurrencyAttribute(string? amountCurrencyId) {
        return amountCurrencyId == null
            ? string.Empty
            : " currencyID=\"" + amountCurrencyId + "\"";
    }

    private static string CreateLineApplicableTradeTax(
        bool includeLineTradeTax,
        bool includeLineTradeTaxTypeCode,
        bool includeLineTradeTaxCategoryCode,
        bool includeLineTradeTaxRate,
        string lineTradeTaxTypeCodeValue,
        string lineTradeTaxCategoryCodeValue,
        string lineTradeTaxRateValue) {
        if (!includeLineTradeTax) {
            return string.Empty;
        }

        string typeCode = includeLineTradeTaxTypeCode
            ? "<ram:TypeCode>" + lineTradeTaxTypeCodeValue + "</ram:TypeCode>"
            : string.Empty;
        string categoryCode = includeLineTradeTaxCategoryCode
            ? "<ram:CategoryCode>" + lineTradeTaxCategoryCodeValue + "</ram:CategoryCode>"
            : string.Empty;
        string rate = includeLineTradeTaxRate
            ? "<ram:RateApplicablePercent>" + lineTradeTaxRateValue + "</ram:RateApplicablePercent>"
            : string.Empty;

        return "<ram:ApplicableTradeTax>" +
               typeCode +
               categoryCode +
               rate +
               "</ram:ApplicableTradeTax>";
    }

    private static byte[] CreateMinimalIccProfile(string colorSpace = "RGB ") {
        byte[] profile = new byte[132];
        profile[3] = 132;
        Encoding.ASCII.GetBytes(colorSpace, 0, 4, profile, 16);
        profile[36] = (byte)'a';
        profile[37] = (byte)'c';
        profile[38] = (byte)'s';
        profile[39] = (byte)'p';
        return profile;
    }
}
