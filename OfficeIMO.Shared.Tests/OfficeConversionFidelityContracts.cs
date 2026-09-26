using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class OfficeConversionFidelityContracts {
    [Fact]
    public void CommonProjectionPreservesLegacyDiagnosticsWithEmptyText() {
        var report = new WordMarkdownConversionReport(new[] {
            new WordMarkdownConversionDiagnostic(
                "WORD_EMPTY_MESSAGE",
                string.Empty,
                OfficeConversionLossKind.Omission)
        });

        OfficeConversionFidelityDiagnostic diagnostic = Assert.Single(report.FidelityDiagnostics);
        Assert.Equal(string.Empty, diagnostic.Message);
        Assert.Equal(OfficeConversionLossKind.Omission, diagnostic.LossKind);
        Assert.True(report.HasLoss);
        Assert.Throws<WordMarkdownConversionException>(() => report.RequireNoLoss());
    }
}
