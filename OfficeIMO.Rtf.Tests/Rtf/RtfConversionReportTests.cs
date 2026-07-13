using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfConversionReportTests {
    [Fact]
    public void Preserved_And_Substituted_Actions_Do_Not_Fail_Strict_Mode() {
        var report = new RtfConversionReport();
        report.Add(RtfConversionSeverity.Information, "Preserved", "Preserved.", RtfConversionAction.Preserved);
        report.Add(RtfConversionSeverity.Warning, "Substituted", "Substituted.", RtfConversionAction.Substituted);

        report.RequireNoLoss();

        Assert.False(report.HasLoss);
        Assert.Equal(2, report.Diagnostics.Count);
    }

    [Theory]
    [InlineData(RtfConversionAction.Flattened)]
    [InlineData(RtfConversionAction.Omitted)]
    [InlineData(RtfConversionAction.Blocked)]
    public void Loss_Actions_Fail_Strict_Mode(RtfConversionAction action) {
        var report = new RtfConversionReport();
        report.Add(RtfConversionSeverity.Warning, "Loss", "Loss occurred.", action, "Body/0", "feature", 2, "detail");

        RtfConversionLossException exception = Assert.Throws<RtfConversionLossException>(() => report.RequireNoLoss());

        Assert.Same(report, exception.Report);
        RtfConversionDiagnostic diagnostic = Assert.Single(report.Diagnostics);
        Assert.Equal("Body/0", diagnostic.SourcePath);
        Assert.Equal("feature", diagnostic.Feature);
        Assert.Equal(2, diagnostic.Count);
        Assert.Equal("detail", diagnostic.Detail);
    }

    [Fact]
    public void Generic_Result_Returns_Value_After_Strict_Check() {
        var report = new RtfConversionReport();
        var result = new RtfConversionResult<string>("value", report);

        Assert.True(result.Succeeded);
        Assert.False(result.HasLoss);
        Assert.Equal("value", result.RequireValue());
        Assert.Equal("value", result.RequireNoLoss());
    }

    [Fact]
    public void Generic_Result_Distinguishes_Errors_From_Fidelity_Loss() {
        var report = new RtfConversionReport();
        report.Add(RtfConversionSeverity.Warning, "Flattened", "Flattened.", RtfConversionAction.Flattened);
        var result = new RtfConversionResult<string>("value", report);

        Assert.True(result.Succeeded);
        Assert.True(result.HasLoss);
        Assert.Equal("value", result.RequireValue());
        Assert.Throws<RtfConversionLossException>(() => result.RequireNoLoss());
    }
}
