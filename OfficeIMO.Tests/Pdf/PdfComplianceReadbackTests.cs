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
