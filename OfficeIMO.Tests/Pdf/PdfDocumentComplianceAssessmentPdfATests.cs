using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentComplianceAssessmentTests {

    [Fact]
    public void AssessComplianceUsesGeneratedStandardFontUsageToReportMissingEmbedding() {
        PdfDocument document = CreatePdfA3GroundworkDocument()
            .Paragraph(paragraph => paragraph.Text("Generated body text requires a standard-font resource."));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfA3B);

        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Missing);
    }

    [Fact]
    public void AssessComplianceUsesGeneratedStandardFontUsageToAcceptEmbeddedMapping() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        PdfDocument document = CreatePdfA3GroundworkDocument()
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath), "HelveticaAudit")
            .Paragraph(paragraph => paragraph.Text("Generated body text has an embedded standard-font mapping."));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfA3B);

        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "verapdf-validation", PdfComplianceRequirementStatus.Unsupported);
        Assert.False(report.IsReady);
    }

    [Fact]
    public void AssessComplianceUsesPageScopedEmbeddedFontFamilyCoverage() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        PdfDocument document = CreatePdfA3GroundworkDocument()
            .Page(page => page
                .UseFontFamily("Scoped Audit Font", fontPath)
                .Content(content => content.Item(item => item.Paragraph(paragraph => paragraph.Text("Page-scoped embedded font coverage.")))));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfA3B);

        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Satisfied);
    }

    [Fact]
    public void AssessComplianceUsesConfiguredProfileWhenNoProfileIsSupplied() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = CreatePdfA3GroundworkOptions();
        options.ComplianceProfile = PdfComplianceProfile.PdfA3B;

        PdfDocument document = PdfDocument.Create(options)
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath), "HelveticaAudit")
            .Paragraph(paragraph => paragraph.Text("Configured profile readiness."));

        PdfComplianceReadinessReport report = document.AssessCompliance();

        Assert.Equal(PdfComplianceProfile.PdfA3B, report.Profile);
        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Satisfied);
    }

    [Fact]
    public void AssessComplianceReportsInvalidEmbeddedMappingBeforeGeneration() {
        PdfDocument document = CreatePdfA3GroundworkDocument()
            .EmbedStandardFont(PdfStandardFont.Helvetica, new byte[] { 1 }, "HelveticaAudit")
            .Paragraph(paragraph => paragraph.Text("Invalid embedded font mapping."));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfA3B);

        PdfComplianceRequirement requirement = AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("invalid embedded TrueType", requirement.Diagnostic);
    }

    [Fact]
    public void InvalidEmbeddedMappingStillFailsGenerationAfterComplianceAssessment() {
        PdfDocument document = CreatePdfA3GroundworkDocument()
            .EmbedStandardFont(PdfStandardFont.Helvetica, new byte[] { 1 }, "HelveticaAudit")
            .Paragraph(paragraph => paragraph.Text("Invalid embedded font mapping."));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfA3B);
        NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.ToBytes());

        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("TrueType font", exception.Message);
    }

}
