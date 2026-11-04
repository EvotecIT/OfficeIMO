using System.IO;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Theory]
    [InlineData(OfficeImageExportDiagnosticSeverity.Info, OfficeConversionLossKind.Omission, OfficeDiagnosticSeverity.Information)]
    [InlineData(OfficeImageExportDiagnosticSeverity.Warning, OfficeConversionLossKind.None, OfficeDiagnosticSeverity.Warning)]
    public void ImageConversionReport_PreservesIndependentSeverityAndLoss(
        OfficeImageExportDiagnosticSeverity severity, OfficeConversionLossKind loss, OfficeDiagnosticSeverity expected) {
        var result = CreateImageConversionEvidence(severity, loss);
        var report = OfficeImageExportEvidence.CreateReport("Word", result);
        var finding = Assert.Single(report.Diagnostics);
        Assert.Equal(expected, finding.Severity);
        Assert.Equal(loss, finding.LossKind);
        Assert.Equal("Page 2", finding.Location);
        Assert.Equal("Word", finding.Format);
        Assert.Equal(OfficeDiagnosticCodes.ImageSourceDecodeOmitted, finding.Code);
        Assert.Equal(loss != OfficeConversionLossKind.None, report.HasLoss);
    }

    [Fact]
    public void ImageConversionFailOnLoss_RejectsBeforeWritingCallerStream() {
        var result = CreateImageConversionEvidence(OfficeImageExportDiagnosticSeverity.Info, OfficeConversionLossKind.Omission);
        var options = new OfficeImageExportOptions { FailOnLoss = true };
        var operation = new OfficeImageExportOperation<OfficeImageExportOptions>(options, OfficeImageExportFormat.Png,
            OfficeImageExportFileConflictPolicy.FailIfExists, (_, _) => result);
        using var destination = new MemoryStream();
        var error = Assert.Throws<OfficeConversionException>(() => operation.Save(destination));
        Assert.True(error.Report.HasLoss);
        Assert.Equal(0, destination.Length);
        Assert.True(destination.CanWrite);
    }

    private static OfficeImageExportResult CreateImageConversionEvidence(OfficeImageExportDiagnosticSeverity severity, OfficeConversionLossKind loss) =>
        new OfficeImageExportResult(OfficeImageExportFormat.Png, 1, 1,
            OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White)), diagnostics: new[] {
                new OfficeImageExportDiagnostic(severity, OfficeDiagnosticCodes.ImageSourceDecodeOmitted,
                    "Image finding", "Page 2", loss)
            });
}
