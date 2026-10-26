using OfficeIMO;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public class OfficeConversionReportTests {
    [Fact]
    public void CompletedReportRetainsMetadataAndSnapshotsDiagnostics() {
        var diagnostics = new List<OfficeDiagnostic> { new OfficeDiagnostic("font-selected", OfficeDiagnosticSeverity.Warning, "Font selected.", "page 1", "Pdf") };
        var report = new OfficeConversionReport("input.docx", "Pdf", "output.pdf", 3, TimeSpan.FromSeconds(1), diagnostics);
        diagnostics.Clear();
        Assert.Equal("input.docx", report.Source);
        Assert.Equal("Pdf", report.Target);
        Assert.Equal("output.pdf", report.OutputPath);
        Assert.Equal(3, report.ItemsWritten);
        Assert.Equal(TimeSpan.FromSeconds(1), report.Elapsed);
        OfficeDiagnostic item = Assert.Single(report.Diagnostics);
        Assert.Equal("page 1", item.Location);
        Assert.Equal("Pdf", item.Format);
        Assert.False(report.HasLoss);
        report.ThrowIfLoss();
    }

    [Theory]
    [InlineData(OfficeConversionLossKind.Approximation)]
    [InlineData(OfficeConversionLossKind.Omission)]
    [InlineData(OfficeConversionLossKind.Unassessed)]
    [InlineData(OfficeConversionLossKind.Failure)]
    public void LossIsIndependentOfSeverityAndExceptionsCarryTheReport(OfficeConversionLossKind kind) {
        var finding = new OfficeDiagnostic("feature-affected", OfficeDiagnosticSeverity.Information, "Content affected.", lossKind: kind);
        var report = new OfficeConversionReport("Word", "Pdf", diagnostics: new[] { finding });
        Assert.True(report.HasLoss);
        Assert.Same(report, Assert.Throws<OfficeConversionException>(report.ThrowIfLoss).Report);
        IOfficeConversionReport legacy = report;
        Assert.Equal(kind, Assert.Single(legacy.FidelityDiagnostics).LossKind);
        Assert.Throws<OfficeConversionException>(legacy.RequireNoLoss);
        var result = new OfficeConversionResult<string, OfficeConversionReport>("text", report);
        Assert.True(result.HasLoss);
        Assert.Equal("text", result.RequireValue());
        Assert.Throws<OfficeConversionException>(() => result.RequireNoLoss());
    }

    [Fact]
    public void EmptyAndMalformedReportsHaveExplicitBehavior() {
        var report = new OfficeConversionReport("Markdown", "Word");
        Assert.Empty(report.Diagnostics);
        Assert.Null(report.OutputPath);
        Assert.False(report.HasLoss);
        report.ThrowIfLoss();
        Assert.Throws<ArgumentException>(() => new OfficeConversionReport("", "Pdf"));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeConversionReport("Word", "Pdf", itemsWritten: -1));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeConversionReport("Word", "Pdf", elapsed: TimeSpan.FromTicks(-1)));
        Assert.Throws<ArgumentException>(() => new OfficeConversionReport("Word", "Pdf", diagnostics: new OfficeDiagnostic[] { null! }));
        Assert.Throws<ArgumentException>(() => new OfficeDiagnostic(" ", OfficeDiagnosticSeverity.Warning, "message"));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeDiagnostic("x", (OfficeDiagnosticSeverity)99, "message"));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeDiagnostic("x", OfficeDiagnosticSeverity.Warning, "message", lossKind: (OfficeConversionLossKind)99));
    }
}
