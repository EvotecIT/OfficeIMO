using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfComplianceAnalyzerTests {
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


}
