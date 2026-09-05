using OfficeIMO.Pdf;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfConversionLossPropagationTests {
    [Fact]
    public void ArtifactProofReadsTheSavedBytesWithoutSerializingAgain() {
        var result = new PdfDocumentConversionResult(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Original artifact")), new PdfConversionReport());
        byte[] original = result.ToBytes();
        result.Value.Paragraph(paragraph => paragraph.Text("Added after serialization"));
        var options = new PdfConversionProofOptions();
        options.RequiredTextMarkers.Add("Original artifact");
        PdfConversionProofReport proof = result.AssessArtifactProof(original, options);
        Assert.True(proof.IsSatisfied, proof.Summary);
        Assert.DoesNotContain("Added after serialization", proof.ExtractedText, StringComparison.Ordinal);
        Assert.Equal(original.Length, proof.ArtifactByteCount);
        using var sha = System.Security.Cryptography.SHA256.Create();
        Assert.Equal(BitConverter.ToString(sha.ComputeHash(original)).Replace("-", string.Empty).ToLowerInvariant(), proof.ArtifactSha256);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImageExportCapturesLateInformationalLossAndRejectsItBeforeDelivery(bool fluent) {
        var report = new PdfConversionReport();
        PdfDocument document = PdfDocument.Create().Deferred(_ => {
            report.Add(new PdfConversionWarning("renderer", "LateLoss", "page:1", "Late omission",
                PdfConversionWarningSeverity.Information, OfficeConversionLossKind.Omission));
            return item => item.Paragraph(paragraph => paragraph.Text("Content"));
        });
        var conversion = new PdfDocumentConversionResult(document, report);
        var builder = conversion.ToImages();
        Assert.Empty(conversion.Warnings);

        OfficeImageExportResult image = Assert.Single(fluent
            ? builder.AsSvg().Export()
            : conversion.ExportImages(OfficeImageExportFormat.Svg));
        Assert.Contains(image.Diagnostics, diagnostic => diagnostic.Code == "LateLoss" &&
            diagnostic.LossKind == OfficeConversionLossKind.Omission);
        var strict = new PdfImageExportOptions { Policy = new OfficeImageExportPolicy { RequireNoLoss = true } };
        int delivered = 0;
        var failure = Assert.Throws<OfficeImageExportPolicyException>(() => {
            if (fluent) conversion.ToImages(strict).AsSvg().ExportEach(_ => delivered++);
            else conversion.ExportImages(OfficeImageExportFormat.Svg, strict);
        });
        Assert.Equal(0, delivered);
        Assert.Contains(failure.Diagnostics, diagnostic => diagnostic.Code == "LateLoss");
    }

    [Fact]
    public void ImageExportRetainsLossFromAnEarlierSemanticProjectionStage() {
        var sourceReport = new PdfConversionReport();
        sourceReport.Add(new PdfConversionWarning("source", "UnsupportedObject", "source:1", "Object omitted",
            PdfConversionWarningSeverity.Information, OfficeConversionLossKind.Omission));
        var conversion = new PdfDocumentConversionResult(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Content")), new PdfConversionReport())
            .WithSourceConversionReport(sourceReport);
        OfficeImageExportResult image = Assert.Single(conversion.ExportImages(OfficeImageExportFormat.Svg));
        Assert.Contains(image.Diagnostics, diagnostic => diagnostic.Code == "SourceConversionLoss" &&
            diagnostic.LossKind != OfficeConversionLossKind.None);
        var strict = new PdfImageExportOptions { Policy = new OfficeImageExportPolicy { RequireNoLoss = true } };
        var failure = Assert.Throws<OfficeImageExportPolicyException>(() => conversion.ToImages(strict).AsSvg().Export());
        Assert.Contains(failure.Diagnostics, diagnostic => diagnostic.Code == "SourceConversionLoss");
    }

    [Theory]
    [InlineData(OfficeConversionLossKind.Approximation)]
    [InlineData(OfficeConversionLossKind.Omission)]
    [InlineData(OfficeConversionLossKind.Failure)]
    public void InformationalLossSurvivesResultSnapshotsAndStrictAcceptance(OfficeConversionLossKind lossKind) {
        var report = new PdfConversionReport();
        report.Add(new PdfConversionWarning("format-renderer", "RasterFallback", "source:1",
            "A vector was rasterized.", PdfConversionWarningSeverity.Information, lossKind));
        var result = new PdfDocumentConversionResult(PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Content")), report);

        Assert.Equal(lossKind, Assert.Single(result.Warnings).LossKind);
        Assert.True(result.HasLoss);
        Assert.Equal(PdfConversionFidelityStatus.Degraded, result.Report.FidelityStatus);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        result.ToBytes();
        Assert.Equal(lossKind, Assert.Single(result.Warnings).LossKind);
    }

    [Fact]
    public void DiagnosticSeverityRetainsItsLegacyDefaultsWithoutInventingInformationalLoss() {
        Assert.Equal(OfficeConversionLossKind.None,
            new PdfConversionWarning("renderer", "Info", "page:1", "Retained metadata", PdfConversionWarningSeverity.Information).LossKind);
        Assert.Equal(OfficeConversionLossKind.Approximation,
            new PdfConversionWarning("renderer", "Fallback", "page:1", "Approximate layout").LossKind);
        Assert.Equal(OfficeConversionLossKind.Failure,
            new PdfConversionWarning("renderer", "Error", "page:1", "Cannot render", PdfConversionWarningSeverity.Error).LossKind);
    }
}
