using OfficeIMO.Pdf;
using OfficeIMO.Drawing;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;
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

    [Fact]
    public void ArtifactProofPreservesTypedConversionLossForStrictAcceptance() {
        var sourceReport = new PdfConversionReport();
        sourceReport.Add(new PdfConversionWarning(
            "source",
            "SOURCE_OMISSION",
            "source:1",
            "Source content was omitted.",
            PdfConversionWarningSeverity.Warning,
            OfficeConversionLossKind.Omission));
        var conversion = new PdfDocumentConversionResult(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Content")),
            new PdfConversionReport())
            .WithSourceConversionReport(sourceReport);

        PdfConversionProofReport proof = conversion.AssessProof();
        IOfficeConversionReport commonReport = proof;

        OfficeConversionFidelityDiagnostic diagnostic = Assert.Single(commonReport.FidelityDiagnostics);
        Assert.True(commonReport.HasLoss);
        Assert.Equal(OfficeConversionLossKind.Omission, diagnostic.LossKind);
        Assert.Equal("source", diagnostic.Source);
        Assert.Equal("source:1", diagnostic.Location);
        Assert.Throws<InvalidOperationException>(commonReport.RequireNoLoss);
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

    [Theory]
    [InlineData(OfficeConversionLossKind.Approximation)]
    [InlineData(OfficeConversionLossKind.Omission)]
    [InlineData(OfficeConversionLossKind.Failure)]
    public void ImageExportRetainsExactLossFromAnEarlierSemanticProjectionStage(OfficeConversionLossKind lossKind) {
        var sourceReport = new PdfConversionReport();
        sourceReport.Add(new PdfConversionWarning("source", "UnsupportedObject", "source:1", "Object omitted",
            PdfConversionWarningSeverity.Information, lossKind));
        var conversion = new PdfDocumentConversionResult(
            PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Content")), new PdfConversionReport())
            .WithSourceConversionReport(sourceReport);
        var permissive = new PdfImageExportOptions {
            Policy = new OfficeImageExportPolicy { RequireNoFailures = false }
        };
        OfficeImageExportResult image = Assert.Single(conversion.ExportImages(OfficeImageExportFormat.Svg, permissive));
        OfficeImageExportDiagnostic projected = Assert.Single(image.Diagnostics,
            diagnostic => diagnostic.Code == "UnsupportedObject");
        Assert.Equal(lossKind, projected.LossKind);
        Assert.Equal("source:1", projected.Source);
        OfficeConversionFidelityDiagnostic fidelity = Assert.Single(image.CreateReport().FidelityDiagnostics,
            diagnostic => diagnostic.Code == "UnsupportedObject");
        Assert.Equal(lossKind, fidelity.LossKind);
        Assert.Equal("source", fidelity.Source);
        Assert.Equal("source:1", fidelity.Location);
        var strict = new PdfImageExportOptions { Policy = new OfficeImageExportPolicy { RequireNoLoss = true } };
        var failure = Assert.Throws<OfficeImageExportPolicyException>(() => conversion.ToImages(strict).AsSvg().Export());
        Assert.Contains(failure.Diagnostics, diagnostic =>
            diagnostic.Code == "UnsupportedObject" && diagnostic.LossKind == lossKind);
    }

    [Fact]
    public void ImageExportFailsClosedWhenAnUpstreamReportViolatesTheTypedLossContract() {
        var conversion = new PdfDocumentConversionResult(
                PdfDocument.Create().Paragraph(paragraph => paragraph.Text("Content")),
                new PdfConversionReport())
            .WithSourceConversionReport(new AggregateOnlyLossReport());

        OfficeImageExportPolicyException failure = Assert.Throws<OfficeImageExportPolicyException>(() =>
            conversion.ExportImages(OfficeImageExportFormat.Svg));

        OfficeImageExportDiagnostic diagnostic = Assert.Single(failure.Diagnostics,
            item => item.Code == "SourceConversionDiagnosticContractMismatch");
        Assert.Equal(OfficeConversionLossKind.Failure, diagnostic.LossKind);
        OfficeConversionFidelityDiagnostic contractFailure = Assert.Single(
            conversion.FidelityDiagnostics,
            item => item.Code == "CONVERSION_REPORT_UNTYPED_LOSS");
        Assert.Equal(OfficeConversionLossKind.Failure, contractFailure.LossKind);
        PdfConversionProofReport proof = conversion.AssessProof();
        Assert.Contains(proof.FidelityDiagnostics,
            item => item.Code == "CONVERSION_REPORT_UNTYPED_LOSS"
                && item.LossKind == OfficeConversionLossKind.Failure);
        Assert.Throws<InvalidOperationException>(proof.RequireNoLoss);
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

    [Fact]
    public void ExcelTableOnlyProjectionTreatsOmittedPageContentAsLoss() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Invoice reference INV-1001"))
            .ToBytes();
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);

        PdfExcelTableImportResult result = logical.ImportTablesToExcelDocumentResult();
        using (result.Value) {
            Assert.Empty(result.Report.Entries);
            Assert.True(result.Report.HasOmittedPageContent);
            Assert.True(result.HasLoss);
            Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        }
    }

    [Fact]
    public void WordProjectionRetainsInformationalLossSemantics() {
        var shared = new PdfConversionReport();
        shared.Add(new PdfConversionWarning(
            "OfficeIMO.Word.Pdf",
            "PdfVisualOnlyObject",
            "page:1",
            "The object was retained only as a visual fallback.",
            PdfConversionWarningSeverity.Information,
            OfficeConversionLossKind.Approximation));

        var report = new PdfWordConversionReport(shared);

        Assert.True(report.HasLoss);
        Assert.Throws<InvalidOperationException>(() => report.RequireNoLoss());
    }

    [Fact]
    public void SemanticHtmlReportsLayoutReconstructionAsTypedLoss() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Invoice reference INV-1001"))
            .ToBytes();
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);

        PdfHtmlConversionResult result = logical.ToHtmlResult(PdfToHtmlOptions.CreateSemanticProfile());

        Assert.True(result.Report.HasLoss);
        Assert.Contains(result.Report.Warnings, static warning =>
            warning.Code == "PdfSemanticLayoutReflowed" &&
            warning.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
    }

    [Fact]
    public void EditablePowerPointReportsReconstructionAsTypedLoss() {
        byte[] source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Invoice reference INV-1001"))
            .ToBytes();
        PdfDocumentReadResult logical = PdfDocumentReadResult.Load(source);

        PdfPowerPointConversionResult result = logical.ToPowerPointPresentationResult(
            PdfToPowerPointOptions.CreateEditableContent());
        using (result.Value) {
            Assert.True(result.Report.HasLoss);
            Assert.Contains(result.Report.Warnings, static warning =>
                warning.Code == "PdfEditableContentReconstructed" &&
                warning.LossKind == OfficeConversionLossKind.Approximation);
            Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        }
    }

    private sealed class AggregateOnlyLossReport : IOfficeConversionReport {
        public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics { get; } =
            Array.Empty<OfficeConversionFidelityDiagnostic>();

        public bool HasLoss => true;

        public void RequireNoLoss() => throw new InvalidOperationException("Aggregate loss.");
    }
}
