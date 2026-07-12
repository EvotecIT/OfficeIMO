using OfficeIMO.Core;
using Xunit;

namespace OfficeIMO.Tests;

public class OfficeCoreContractTests {
    [Fact]
    public void ConversionResultReportsLossAndReturnsSuccessfulValue() {
        var result = new OfficeConversionResult<string>("document", new[] {
            new OfficeDiagnostic(
                "TEST001",
                OfficeDiagnosticSeverity.Warning,
                "A feature was simplified.",
                OfficeDiagnosticImpact.Simplified)
        });

        Assert.True(result.Succeeded);
        Assert.True(result.HasLoss);
        Assert.Equal("document", result.RequireValue());
    }

    [Fact]
    public void OperationResultThrowsWithOriginalDiagnosticsWhenErrorsExist() {
        var diagnostic = new OfficeDiagnostic(
            "TEST002",
            OfficeDiagnosticSeverity.Error,
            "The operation failed.",
            OfficeDiagnosticImpact.Failure);
        var result = new OfficeOperationResult<string>(string.Empty, new[] { diagnostic });

        OfficeOperationException exception = Assert.Throws<OfficeOperationException>(() => result.RequireValue());

        Assert.False(result.Succeeded);
        Assert.Same(diagnostic, Assert.Single(exception.Diagnostics));
        Assert.Contains("TEST002", exception.Message, StringComparison.Ordinal);
    }
}
