using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfComplianceAnalyzerTests {
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
    public void PdfA3UReadinessAcceptsEmbeddedType0FontUnicodeCoverage() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        PdfDocument document = PdfDocument.Create(new PdfOptions()
                .SetPdfAIdentification(3, "U")
                .SetSrgbOutputIntent())
            .UseFontFamily("Unicode readiness font", fontPath)
            .Paragraph(paragraph => paragraph.Text("Embedded Type0 ToUnicode coverage."));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfA3U);

        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "standard-font-to-unicode", PdfComplianceRequirementStatus.Satisfied);
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


}
