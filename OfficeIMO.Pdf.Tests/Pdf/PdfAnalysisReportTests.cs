using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfAnalysisReportTests {
    [Fact]
    public void AnalyzeReturnsOneCoherentHealthAndCapabilityReport() {
        byte[] bytes = PdfDocument.Create()
            .Meta(title: "Analysis source")
            .Paragraph(paragraph => paragraph.Text("Readable analysis content"))
            .ToBytes();

        PdfAnalysisReport report = PdfDocument.Open(bytes).Analyze();

        Assert.True(report.IsHealthy);
        Assert.True(report.CanRead);
        Assert.True(report.CanRewrite);
        Assert.Equal("Analysis source", report.Info.Metadata.Title);
        Assert.Equal(report.Info.PageCount, report.Diagnostics.PageCount);
        Assert.Same(report.Diagnostics, report.Optimization.Diagnostics);
        Assert.False(report.Signatures.HasSignatures);
        Assert.True(report.AppendOnlyMutation.CanAppendMetadata);
        Assert.Null(report.Compliance);
    }

    [Fact]
    public void AnalyzeCanIncludeComplianceReadbackForTheOpenedArtifact() {
        byte[] bytes = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Compliance analysis source"))
            .ToBytes();
        PdfDocument document = PdfDocument.Open(bytes);

        PdfAnalysisReport report = document.Analyze(PdfComplianceProfile.PdfA2B);
        PdfComplianceProofReport proof = document.AssessComplianceProof(PdfComplianceProfile.PdfA2B);

        Assert.NotNull(report.Compliance);
        Assert.Equal(PdfComplianceProfile.PdfA2B, report.Compliance!.Profile);
        Assert.Equal(bytes.LongLength, proof.ArtifactSizeBytes);
        Assert.False(string.IsNullOrWhiteSpace(proof.ArtifactSha256));
    }
}
