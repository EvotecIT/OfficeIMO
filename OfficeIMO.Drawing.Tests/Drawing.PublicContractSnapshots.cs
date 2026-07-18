using System.Collections.Generic;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingPublicContractSnapshotTests {
    [Fact]
    public void ImageExportResultSnapshotsCallerOwnedBuffersAndDiagnostics() {
        byte[] source = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.CornflowerBlue));
        var diagnostics = new List<OfficeImageExportDiagnostic> {
            new(OfficeImageExportDiagnosticSeverity.Warning, "Sample", "Sample diagnostic")
        };
        var result = new OfficeImageExportResult(
            OfficeImageExportFormat.Png,
            width: 1,
            height: 1,
            bytes: source,
            diagnostics: diagnostics);

        byte[] expected = (byte[])source.Clone();
        source[0] = 9;
        diagnostics.Clear();
        byte[] returned = result.Bytes;
        returned[1] = 9;

        Assert.Equal(expected, result.Bytes);
        Assert.Single(result.Diagnostics);
        Assert.Equal("image/png", result.MimeType);
        Assert.Equal(".png", result.FileExtension);
    }

    [Fact]
    public void ChartLayoutSnapshotsLabelSelectionsAndIsImmutable() {
        int[] seriesIndexes = { 1 };
        int[] pointIndexes = { 2 };
        var pointsBySeries = new Dictionary<int, IReadOnlyCollection<int>> {
            [1] = pointIndexes
        };
        var layout = new OfficeChartLayout(
            dataLabelSeriesIndexes: seriesIndexes,
            dataLabelPointIndexes: pointsBySeries);

        seriesIndexes[0] = 9;
        pointIndexes[0] = 9;
        pointsBySeries.Clear();

        Assert.Equal(new[] { 1 }, layout.DataLabelSeriesIndexes);
        Assert.Equal(new[] { 2 }, layout.DataLabelPointIndexes![1]);
        Assert.False(typeof(OfficeChartLayout).GetProperty(nameof(OfficeChartLayout.DataLabelSeriesIndexes))!.CanWrite);
        Assert.False(typeof(OfficeChartLayout).GetProperty(nameof(OfficeChartLayout.DataLabelPointIndexes))!.CanWrite);
    }

    [Fact]
    public void ChartStyleBorderVisibilityIsConstructorOwned() {
        Assert.True(OfficeChartStyle.Default.ShowBorder);
        Assert.False(new OfficeChartStyle(showBorder: false).ShowBorder);
        Assert.False(typeof(OfficeChartStyle).GetProperty(nameof(OfficeChartStyle.ShowBorder))!.CanWrite);
    }
}
