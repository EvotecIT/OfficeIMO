using System;
using System.IO;
using System.Linq;
using System.Reflection;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public sealed class WordSharedChartPdfTests {
    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void ExplodedSliceSurvivesNativePdfProjection(OfficeChartKind kind) {
        using var document = WordDocument.Create();
        WordChart chart = document.AddChart(kind, new OfficeChartData(new[] { "A", "B" },
            new[] { new OfficeChartSeries("Status", new[] { 7d, 3d }).WithPointExplosions(new[] { 25, 0 }) }));
        MethodInfo factory = typeof(WordPdfConverterExtensions).GetMethod("TryCreateNativeWordChartSnapshot",
            BindingFlags.NonPublic | BindingFlags.Static)!;
        object?[] arguments = { chart, null, null };
        Assert.True((bool)factory.Invoke(null, arguments)!);
        OfficeChartSnapshot projected = Assert.IsType<OfficeChartSnapshot>(arguments[1]);
        Assert.Equal(new[] { 25, 0 }, projected.Data.Series.Single().PointExplosions);
        var options = new WordToPdfOptions { IncludePageNumbers = false };
        Assert.NotEmpty(document.ToPdfBytes(options));
        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
    }

    [Theory]
    [InlineData("empty")]
    [InlineData("noFill")]
    [InlineData("noFillOutline")]
    [InlineData("emptyEffectList")]
    public void DefaultLegendShapeDoesNotDisableSingleChartPdfWithDataTable(string appearance) {
        using var document = WordDocument.Create();
        WordChart chart = document.AddChart(OfficeChartKind.ColumnClustered,
            new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Count", new[] { 4d }) }));
        C.Chart native = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!;
        native.PlotArea!.AddChild(new C.DataTable(), true);
        C.Legend legend = native.GetFirstChild<C.Legend>()!;
        var properties = new C.ChartShapeProperties();
        if (appearance != "empty") properties.Append(new A.NoFill());
        if (appearance == "noFillOutline") properties.Append(new A.Outline(new A.NoFill()));
        if (appearance == "emptyEffectList") properties.Append(new A.EffectList());
        legend.AddChild(properties, true);

        MethodInfo factory = typeof(WordPdfConverterExtensions).GetMethod("TryCreateNativeWordChartSnapshot",
            BindingFlags.NonPublic | BindingFlags.Static)!;
        object?[] arguments = { chart, null, null };
        Assert.True((bool)factory.Invoke(null, arguments)!);
        Assert.IsType<OfficeChartSnapshot>(arguments[1]);
        var options = new WordToPdfOptions { IncludePageNumbers = false };
        byte[] pdfBytes = document.ToPdfBytes(options);
        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        using var pdf = PdfPigDocument.Open(new MemoryStream(pdfBytes));
        Assert.Single(pdf.GetPages());
    }

    [Fact]
    public void ImportedPieLegendFrameUsesQualifiedSharedStyleInPdf() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "Charts", "LibreOffice", "status-pie.docx");
        using var document = WordDocument.Load(path);
        WordChart chart = Assert.Single(document.Charts);
        Assert.True(chart.TryGetOfficeSnapshot(out var shared));
        Assert.NotNull(shared.Style.LegendBorderColor);

        MethodInfo factory = typeof(WordPdfConverterExtensions).GetMethod("TryCreateNativeWordChartSnapshot",
            BindingFlags.NonPublic | BindingFlags.Static)!;
        object?[] arguments = { chart, null, null };
        Assert.True((bool)factory.Invoke(null, arguments)!);
        var projected = Assert.IsType<OfficeChartSnapshot>(arguments[1]);
        Assert.Equal(shared.Style.LegendBackgroundColor, projected.Style.LegendBackgroundColor);
        Assert.Equal(shared.Style.LegendBorderColor, projected.Style.LegendBorderColor);
        Assert.Equal(shared.Style.LegendBorderWidth, projected.Style.LegendBorderWidth);

        var options = new WordToPdfOptions { IncludePageNumbers = false };
        byte[] pdfBytes = document.ToPdfBytes(options);
        Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
        using var pdf = PdfPigDocument.Open(new MemoryStream(pdfBytes));
        Assert.Single(pdf.GetPages());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SharedCharts_RenderAllLayersAndBubbleDataInPdf(bool bubble) {
        using var document = WordDocument.Create();
        var data = bubble ? new OfficeChartData(new[] { "1", "2" }, new[] {
            OfficeChartSeries.CreateBubble("Measured", new[] { 1d, 2d }, new[] { 3d, 4d }, new[] { 5d, 20d }, OfficeColor.Parse("#224466")) }) :
            new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Columns", new[] { 100d, 200d }),
                new OfficeChartSeries("Ratio", new[] { 1d, 2d }, null, OfficeColor.Parse("#224466"), null, true,
                    renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary) });
        document.AddChart(bubble ? OfficeChartKind.Bubble : OfficeChartKind.ColumnClustered, data, "Shared chart");
        var options = new WordToPdfOptions { IncludePageNumbers = false };
        string path = Path.Combine(Path.GetTempPath(), "officeimo-shared-chart-" + Guid.NewGuid().ToString("N") + ".pdf");
        try {
            document.SaveAsPdf(path, options);
            Assert.DoesNotContain(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported");
            using var pdf = PdfPigDocument.Open(path);
            string text = string.Join(" ", pdf.GetPages().SelectMany(page => page.GetWords()).Select(word => word.Text));
            Assert.Contains("Shared chart", text);
            Assert.Contains(bubble ? "Measured" : "Ratio", text);
            if (!bubble) Assert.Contains("Columns", text);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
