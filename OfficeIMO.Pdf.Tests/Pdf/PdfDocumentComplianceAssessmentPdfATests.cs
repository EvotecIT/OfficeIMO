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
            .Paragraph(paragraph => paragraph.AddText("Generated body text requires a standard-font resource."));

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
            .Paragraph(paragraph => paragraph.AddText("Generated body text has an embedded standard-font mapping."));

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
                .Content(content => content.Item(item => item.AddParagraph(paragraph => paragraph.AddText("Page-scoped embedded font coverage.")))));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfA3B);

        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Satisfied);
    }

    [Fact]
    public void AssessComplianceIncludesNamedFontResourcesInEmbeddedCoverage() {
        string? fontPath = PdfComplianceTestFonts.FindBundledOpenTypeCffFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);
        var options = CreatePdfA3GroundworkOptions()
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Named Compliance", fontData));
        PdfDocument document = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.AddText("Named font coverage is included in generated evidence.").SetFontFamily("Named Compliance"));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfA3B);

        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Satisfied);
    }

    [Fact]
    public void AssessComplianceIgnoresUnusedExplicitStandardDefaultFont() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = CreatePdfA3GroundworkOptions();
        options.DefaultFont = PdfStandardFont.TimesRoman;
        options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
            "Named Compliance Only",
            File.ReadAllBytes(fontPath)));
        PdfDocument document = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.AddText("Only the embedded named font paints text.").SetFontFamily("Named Compliance Only"));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfA3B);

        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Satisfied);
    }

    [Fact]
    public void AssessComplianceIncludesNamedPageTextAndListMarkerResourcesInEmbeddedCoverage() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        byte[] fontData = File.ReadAllBytes(fontPath);
        var options = CreatePdfA3GroundworkOptions()
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "HelveticaAudit")
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Named Surface Compliance", fontData));
        options.PageSetup.Header.Show = true;
        options.PageSetup.Header.Format = "Named header";
        options.PageSetup.Header.FontFamily = "Named Surface Compliance";
        options.PageSetup.Footer.Show = true;
        options.PageSetup.Footer.Format = "Named footer";
        options.PageSetup.Footer.FontFamily = "Named Surface Compliance";
        var listStyle = new PdfListStyle {
            MarkerFontFamily = "Named Surface Compliance"
        };
        PdfDocument document = PdfDocument.Create(options)
            .RichNumbered(
                new[] { new PdfListItem("Embedded body font coverage.", marker: "1.") },
                style: listStyle);

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfA3B);

        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Satisfied);
    }

    [Fact]
    public void AssessComplianceRejectsInvalidNamedFontResources() {
        string? fontPath = PdfComplianceTestFonts.FindLocalTrueTypeFont();
        if (fontPath == null) {
            return;
        }

        var options = CreatePdfA3GroundworkOptions()
            .EmbedStandardFont(PdfStandardFont.Helvetica, File.ReadAllBytes(fontPath), "HelveticaAudit")
            .RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Broken Named Compliance", new byte[] { 1 }));
        PdfDocument document = PdfDocument.Create(options)
            .Paragraph(paragraph => paragraph.AddText("Invalid named font coverage.").SetFontFamily("Broken Named Compliance"));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfA3B);

        PdfComplianceRequirement requirement = AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("broken named compliance", requirement.Diagnostic, StringComparison.OrdinalIgnoreCase);
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
            .Paragraph(paragraph => paragraph.AddText("Configured profile readiness."));

        PdfComplianceReadinessReport report = document.AssessCompliance();

        Assert.Equal(PdfComplianceProfile.PdfA3B, report.Profile);
        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Satisfied);
    }

    [Fact]
    public void AssessComplianceReportsInvalidEmbeddedMappingBeforeGeneration() {
        PdfDocument document = CreatePdfA3GroundworkDocument()
            .EmbedStandardFont(PdfStandardFont.Helvetica, new byte[] { 1 }, "HelveticaAudit")
            .Paragraph(paragraph => paragraph.AddText("Invalid embedded font mapping."));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfA3B);

        PdfComplianceRequirement requirement = AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("invalid embedded TrueType", requirement.Diagnostic);
    }

    [Fact]
    public void InvalidEmbeddedMappingStillFailsGenerationAfterComplianceAssessment() {
        PdfDocument document = CreatePdfA3GroundworkDocument()
            .EmbedStandardFont(PdfStandardFont.Helvetica, new byte[] { 1 }, "HelveticaAudit")
            .Paragraph(paragraph => paragraph.AddText("Invalid embedded font mapping."));

        PdfComplianceReadinessReport report = document.AssessCompliance(PdfComplianceProfile.PdfA3B);
        NotSupportedException exception = Assert.Throws<NotSupportedException>(() => document.ToBytes());

        AssertRequirement(report, "embedded-font-coverage", PdfComplianceRequirementStatus.Missing);
        Assert.Contains("TrueType font", exception.Message);
    }

}
