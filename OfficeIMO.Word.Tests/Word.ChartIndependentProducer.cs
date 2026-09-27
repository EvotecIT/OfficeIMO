using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordChartIndependentProducerTests {
    [Fact]
    public void LibreOfficeStatusPieRetainsPointAppearanceInNativeAndStaticPaths() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "Charts", "LibreOffice", "status-pie.docx");
        using WordDocument document = WordDocument.Load(path);
        WordChart chart = Assert.Single(document.Charts);
        string nativeXml = chart.ChartPart!.ChartSpace!.OuterXml;

        Assert.True(chart.TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot));
        Assert.Equal(OfficeChartKind.Pie, snapshot.ChartKind);
        Assert.Equal(new[] { "Pass", "Fail", "Unknown" }, snapshot.Data.Categories);
        var styles = Assert.Single(snapshot.Data.Series).PointStyles!;
        Assert.Equal(OfficeColor.Parse("#228844"), styles[0]!.FillColor);
        Assert.Equal(OfficeChartHatchPattern.WideForwardDiagonal, styles[1]!.Hatch);
        Assert.Equal(OfficeColor.Parse("#D97706"), styles[1]!.HatchColor);
        Assert.True(styles[2]!.NoFill);
        Assert.Equal(OfficeColor.Parse("#445566"), styles[2]!.OutlineColor);
        Assert.Equal(OfficeStrokeLineJoin.Round, styles[2]!.OutlineJoin);
        Assert.NotNull(snapshot.Style.LegendBackgroundColor);
        Assert.NotNull(snapshot.Style.LegendBorderColor);

        var rendered = document.ExportImage(OfficeImageExportFormat.Png,
            new WordImageExportOptions { Policy = new OfficeImageExportPolicy { RequireNoOmissions = true, RequireNoFailures = true } });
        Assert.NotEmpty(rendered.Bytes);
        Assert.Equal(nativeXml, chart.ChartPart.ChartSpace.OuterXml);

        using var package = document.ToStream();
        using WordDocument reopened = WordDocument.Load(package);
        Assert.True(Assert.Single(reopened.Charts).TryGetOfficeSnapshot(out var second));
        Assert.Equal(OfficeChartHatchPattern.WideForwardDiagonal, second.Data.Series[0].PointStyles![1]!.Hatch);
    }
}
