using System;
using OfficeIMO;
using OfficeIMO.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentTypedSaveDiagnosticsTests {
    [Fact]
    public void FlatDocumentUnrepresentedEntriesAreTypedAsOmissions() {
        var report = new OdfSaveReport(
            Array.Empty<string>(),
            Array.Empty<string>(),
            Array.Empty<string>(),
            new[] { "Attachments/source.bin" });

        OfficeConversionFidelityDiagnostic diagnostic = Assert.Single(report.FidelityDiagnostics);

        Assert.Equal("ODF_ENTRY_OMITTED", diagnostic.Code);
        Assert.Equal(OfficeConversionLossKind.Omission, diagnostic.LossKind);
        Assert.Equal("Attachments/source.bin", diagnostic.Location);
        Assert.True(report.HasLoss);
        Assert.Throws<InvalidOperationException>(() => report.RequireNoLoss());
    }
}
