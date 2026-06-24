using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfComplianceReadbackTests {
    [Fact]
    public void AssessReadback_ReportsPdfA3GroundworkFromSavedPdf() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                FileVersion = PdfFileVersion.Pdf17,
                IncludeStandardFontToUnicodeMaps = true
            })
            .PdfAIdentification(3, "B")
            .Meta(title: "Readback PDF/A", author: "OfficeIMO")
            .Language("en-US")
            .SrgbOutputIntent()
            .Paragraph(paragraph => paragraph.Text("Readback PDF/A groundwork"))
            .ToBytes();

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfA3B, pdf);

        AssertRequirement(report, "readback-pdf-file-version", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "readback-xmp-metadata", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "readback-pdfa-identification", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "readback-output-intent", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "readback-output-intent-policy", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "verapdf-validation", PdfComplianceRequirementStatus.Unsupported);
    }

    [Fact]
    public void AssessReadback_UsesCatalogVersionAsEffectiveVersion() {
        byte[] generated = PdfDocument.Create(new PdfOptions {
                FileVersion = PdfFileVersion.Pdf17,
                IncludeStandardFontToUnicodeMaps = true
            })
            .PdfAIdentification(3, "B")
            .Meta(title: "Readback PDF/A", author: "OfficeIMO")
            .Language("en-US")
            .SrgbOutputIntent()
            .Paragraph(paragraph => paragraph.Text("Readback effective version groundwork"))
            .ToBytes();
        string rewritten = PdfEncoding.Latin1GetString(generated)
            .Replace("%PDF-1.7", "%PDF-1.4")
            .Replace("/Type /Catalog", "/Type /Catalog /Version /1.7");
        byte[] pdf = PdfEncoding.Latin1GetBytes(rewritten);

        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfA3B, info);

        Assert.Equal("1.4", info.HeaderVersion);
        Assert.Equal("1.7", info.CatalogVersion);
        Assert.Equal("1.7", info.EffectiveVersion);
        AssertRequirement(report, "readback-pdf-file-version", PdfComplianceRequirementStatus.Satisfied);
    }

    [Fact]
    public void AssessReadback_RejectsPdf20EffectiveVersionForPdfA3() {
        byte[] generated = PdfDocument.Create(new PdfOptions {
                FileVersion = PdfFileVersion.Pdf17,
                IncludeStandardFontToUnicodeMaps = true
            })
            .PdfAIdentification(3, "B")
            .Meta(title: "Readback PDF/A", author: "OfficeIMO")
            .Language("en-US")
            .SrgbOutputIntent()
            .Paragraph(paragraph => paragraph.Text("Readback effective version groundwork"))
            .ToBytes();
        string rewritten = PdfEncoding.Latin1GetString(generated)
            .Replace("/Type /Catalog", "/Type /Catalog /Version /2.0");
        byte[] pdf = PdfEncoding.Latin1GetBytes(rewritten);

        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfA3B, info);

        Assert.Equal("2.0", info.EffectiveVersion);
        AssertRequirement(report, "readback-pdf-file-version", PdfComplianceRequirementStatus.Missing);
    }

    [Fact]
    public void AssessReadback_ReportsPdfA4GroundworkFromSavedPdf() {
        byte[] pdf = PdfDocument.Create()
            .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA4F, "en-US")
            .Meta(title: "Readback PDF/A-4", author: "OfficeIMO")
            .Paragraph(paragraph => paragraph.Text("Readback PDF/A-4 groundwork"))
            .ToBytes();

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfA4F, pdf);

        AssertRequirement(report, "readback-pdf-file-version", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "readback-xmp-metadata", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "readback-pdfa-identification", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "readback-output-intent", PdfComplianceRequirementStatus.Satisfied);
    }

    [Fact]
    public void AssessReadback_UsesCatalogVersionOverrideForPdfA4() {
        byte[] pdf = PdfDocument.Create()
            .ConfigurePdfAGroundwork(PdfComplianceProfile.PdfA4F, "en-US")
            .Meta(title: "Catalog version PDF/A-4", author: "OfficeIMO")
            .Paragraph(paragraph => paragraph.Text("Catalog version override"))
            .ToBytes();
        string raw = Encoding.ASCII.GetString(pdf)
            .Replace("%PDF-2.0", "%PDF-1.7")
            .Replace("/Type /Catalog", "/Type /Catalog /Version /2.0");
        byte[] header17Catalog20 = Encoding.ASCII.GetBytes(raw);

        PdfDocumentInfo info = PdfInspector.Inspect(header17Catalog20);
        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfA4F, info);

        Assert.Equal("1.7", info.HeaderVersion);
        Assert.Equal("2.0", info.CatalogVersion);
        Assert.Equal("2.0", info.EffectiveVersion);
        AssertRequirement(report, "readback-pdf-file-version", PdfComplianceRequirementStatus.Satisfied);
    }

    [Fact]
    public void AssessReadback_ReportsPdfUaGroundworkFromSavedPdf() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                FileVersion = PdfFileVersion.Pdf17,
                IncludeStandardFontToUnicodeMaps = true
            })
            .ConfigurePdfUaGroundwork("en-US")
            .Meta(title: "Readback PDF/UA", author: "OfficeIMO")
            .H1("Readback PDF/UA")
            .Paragraph(paragraph => paragraph.Text("Readback PDF/UA groundwork"))
            .ToBytes();

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfUa1, pdf);

        AssertRequirement(report, "readback-pdf-file-version", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "readback-pdfua-identification", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "readback-document-title", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "readback-display-document-title", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "readback-document-language", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "readback-marked-catalog", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "readback-structure-root", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "readback-parent-tree-next-key", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "readback-document-structure-element", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "pdfua-validation", PdfComplianceRequirementStatus.Unsupported);
    }

    [Fact]
    public void AssessReadback_ReportsPdfUa2GroundworkFromSavedPdf() {
        byte[] pdf = PdfDocument.Create()
            .ConfigurePdfUaGroundwork(PdfComplianceProfile.PdfUa2, "en-US")
            .Meta(title: "Readback PDF/UA-2", author: "OfficeIMO")
            .H1("Readback PDF/UA-2")
            .Paragraph(paragraph => paragraph.Text("Readback PDF/UA-2 groundwork"))
            .ToBytes();

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfUa2, pdf);

        AssertRequirement(report, "readback-pdf-file-version", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "readback-pdfua-identification", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "readback-structure-element-count", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "readback-marked-content-references", PdfComplianceRequirementStatus.Satisfied);
        AssertRequirement(report, "pdfua-validation", PdfComplianceRequirementStatus.Unsupported);
    }

    [Fact]
    public void AssessReadback_ReportsMissingPdfUaGroundwork() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Plain PDF"))
            .ToBytes();

        PdfComplianceReadinessReport report = PdfComplianceAnalyzer.AssessReadback(PdfComplianceProfile.PdfUa1, pdf);

        AssertRequirement(report, "readback-pdfua-identification", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "readback-document-title", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "readback-document-language", PdfComplianceRequirementStatus.Missing);
        AssertRequirement(report, "readback-marked-catalog", PdfComplianceRequirementStatus.Missing);
    }

    private static PdfComplianceRequirement AssertRequirement(PdfComplianceReadinessReport report, string id, PdfComplianceRequirementStatus status) {
        PdfComplianceRequirement? requirement = report.FindRequirement(id);
        Assert.NotNull(requirement);
        Assert.Equal(status, requirement.Status);
        return requirement;
    }
}
