using AngleSharp.Html.Parser;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System.Text;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class HtmlWordChartTests {
    [Fact]
    public void Export_PreservesPictureAndDistinctChartsInsideTheSameRun() {
        using var document = WordDocument.Create();
        document.AddChart(OfficeChartKind.ColumnClustered, Data(), title: "First chart");
        document.AddChart(OfficeChartKind.Doughnut, Data(), title: "Second chart");
        using var pixels = new MemoryStream(Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAABFSURBVEhLY1BNfv2flpgBXYDaeBhaILCzkSKMbt6oBRgY3bxRCzAwunmjFmBgdPNGLcDA6OaNWoCB0c3DsIDaeNQCghgAFxBXzP1LTe4AAAAASUVORK5CYII="));
        document.AddParagraph().AddImage(pixels, "marker.png", 10, 10, description: "Picture marker");
        var runs = document.Paragraphs.SelectMany(paragraph => paragraph.GetRuns()).Select(run => run._run!).ToArray();
        var secondDrawing = runs[1].GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Drawing>()!;
        var pictureDrawing = runs[2].GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Drawing>()!;
        secondDrawing.Remove(); pictureDrawing.Remove();
        runs[0].PrependChild(pictureDrawing); runs[0].Append(secondDrawing);
        runs[1].Remove(); runs[2].Remove();
        var result = document.ToHtmlResult();
        var images = new HtmlParser().ParseDocument(result.RequireValue()).QuerySelectorAll("img");
        Assert.Equal(3, images.Length);
        Assert.Equal(new[] { "Picture marker", "First chart", "Second chart" }, images.Select(image => image.GetAttribute("alt")));
        Assert.StartsWith("data:image/png;base64,", images[0].GetAttribute("src"));
    }

    [Fact]
    public void Export_PreservesTextAdjacentToChartInTheSameRun() {
        using var document = WordDocument.Create();
        document.AddChart(OfficeChartKind.ColumnClustered, Data());
        var run = document.Paragraphs.Single().GetRuns().Single()._run!;
        run.PrependChild(new DocumentFormat.OpenXml.Wordprocessing.Text("Before chart "));
        run.Append(new DocumentFormat.OpenXml.Wordprocessing.Text(" after chart"));
        var result = document.ToHtmlResult();
        string html = result.RequireValue();
        Assert.Contains("Before chart", html);
        Assert.Contains("after chart", html);
        Assert.True(html.IndexOf("Before chart", StringComparison.Ordinal) < html.IndexOf("data:image/svg+xml", StringComparison.Ordinal));
        Assert.True(html.IndexOf("data:image/svg+xml", StringComparison.Ordinal) < html.IndexOf("after chart", StringComparison.Ordinal));
    }

    [Fact]
    public void Export_TraversesChartOnlyHeaderEvenWhenItsRelationshipCannotBeProjected() {
        using var document = WordDocument.Create();
        document.AddChart(OfficeChartKind.ColumnClustered, Data());
        var source = document.Paragraphs.Single().GetRuns().Single()._run!;
        var header = document.Sections[0].GetOrCreateHeader(WordHeaderFooterType.Default).AddParagraph();
        header._paragraph.Append(source.CloneNode(true));
        var result = document.ToHtmlResult(new WordToHtmlOptions { ExportHeadersAndFooters = true });
        var dom = new HtmlParser().ParseDocument(result.RequireValue());
        Assert.NotNull(dom.QuerySelector("header"));
        Assert.True(dom.QuerySelector("header img") != null || result.Report.Diagnostics.Any(item => item.Code == "WordChartOmitted"));
    }

    [Theory]
    [InlineData(OfficeChartKind.Doughnut)]
    [InlineData(OfficeChartKind.Bubble)]
    [InlineData(OfficeChartKind.ColumnClustered)]
    public void Export_RendersNativeChartsInDocumentOrder(OfficeChartKind kind) {
        using var document = WordDocument.Create();
        document.AddParagraph("Before");
        var data = kind == OfficeChartKind.Bubble
            ? new OfficeChartData(new[] { "1", "2" }, new[] { OfficeChartSeries.CreateBubble("Values", new[] { 1d, 2d }, new[] { 3d, 4d }, new[] { 5d, 10d }, OfficeColor.Parse("#224466")) })
            : new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Values", new[] { 3d, 4d }, null, OfficeColor.Parse("#224466")) });
        document.AddChart(kind, data, title: "A <chart> & report", width: 600, height: 300);
        document.AddParagraph("After");
        var result = document.ToHtmlResult();
        string html = result.RequireValue();
        var dom = new HtmlParser().ParseDocument(html);
        var image = Assert.Single(dom.QuerySelectorAll("img"));
        Assert.Equal("A <chart> & report", image.GetAttribute("alt"));
        Assert.Equal("600", image.GetAttribute("width"));
        Assert.Equal("300", image.GetAttribute("height"));
        const string prefix = "data:image/svg+xml;base64,";
        Assert.StartsWith(prefix, image.GetAttribute("src"));
        string svg = Encoding.UTF8.GetString(Convert.FromBase64String(image.GetAttribute("src")!.Substring(prefix.Length)));
        Assert.Contains("<svg", svg);
        Assert.Contains("#224466", svg, StringComparison.OrdinalIgnoreCase);
        Assert.True(html.IndexOf("Before", StringComparison.Ordinal) < html.IndexOf("data:image/svg+xml", StringComparison.Ordinal));
        Assert.True(html.IndexOf("data:image/svg+xml", StringComparison.Ordinal) < html.IndexOf("After", StringComparison.Ordinal));
        Assert.Contains(result.Report.Diagnostics, item => item.Code == "WordChartRenderedAsImage");
        Assert.DoesNotContain(result.Report.Diagnostics, item => item.Code == "WordChartOmitted");
    }

    [Fact]
    public void Export_PreservesOutlinedAndHatchedRadialPointsWithoutMutatingTheNativeChart() {
        using var document = WordDocument.Create();
        var series = new OfficeChartSeries("Status", new[] { 3d, 2d, 1d }).WithPointStyles(new OfficeChartPointStyle?[] {
            new(fillColor: OfficeColor.Parse("#224466")),
            new(noFill: true, outlineColor: OfficeColor.Parse("#778899"), outlineWidth: 2, showOutline: true),
            new(hatch: OfficeChartHatchPattern.DiagonalCross, hatchColor: OfficeColor.Parse("#ABCDEF"))
        });
        var chart = document.AddChart(OfficeChartKind.Doughnut, new OfficeChartData(new[] { "Pass", "Unknown", "Fail" }, new[] { series }));
        string native = chart.ChartPart!.ChartSpace!.OuterXml;
        var result = document.ToHtmlResult(new WordToHtmlOptions { EmbedImagesAsBase64 = false });
        var image = new HtmlParser().ParseDocument(result.RequireValue()).QuerySelector("img")!;
        string svg = Encoding.UTF8.GetString(Convert.FromBase64String(image.GetAttribute("src")!.Substring("data:image/svg+xml;base64,".Length)));
        Assert.Contains("#224466", svg, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#778899", svg, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#ABCDEF", svg, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(native, chart.ChartPart.ChartSpace.OuterXml);
    }

    [Fact]
    public void Export_ReportsUnsupportedChartOmission() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Line, Data());
        chart.ChartPart!.ChartSpace!.Descendants<C.LineChartSeries>().Single().AddChild(new C.Smooth { Val = true }, true);
        var result = document.ToHtmlResult();
        Assert.Contains(result.Report.Diagnostics, item => item.Code == "WordChartOmitted");
        Assert.Empty(new HtmlParser().ParseDocument(result.RequireValue()).QuerySelectorAll("img"));
    }

    [Theory]
    [InlineData("image")]
    [InlineData("aggregate")]
    [InlineData("output")]
    public void Export_EnforcesImageAndOutputBudgets(string limit) {
        using var document = WordDocument.Create();
        document.AddChart(OfficeChartKind.ColumnClustered, Data());
        var options = new WordToHtmlOptions { IncludeDefaultCss = false };
        if (limit == "image") options.MaxEmbeddedImageBytes = 16;
        if (limit == "aggregate") options.MaxTotalEmbeddedImageBytes = 16;
        if (limit == "output") options.MaxOutputCharacters = 1024;
        var error = Assert.Throws<HtmlConversionLimitException>(() => document.ToHtml(options));
        Assert.Equal(limit == "output" ? "WordHtmlOutputLimitExceeded" : limit == "image" ? "WordImageSizeLimitExceeded" : "WordImageTotalSizeLimitExceeded", error.Code);
        Assert.Equal(limit == "output" ? options.MaxOutputCharacters : limit == "image" ? options.MaxEmbeddedImageBytes : options.MaxTotalEmbeddedImageBytes, error.Limit);
        Assert.True(error.Actual > error.Limit);
    }

    private static OfficeChartData Data() => new(new[] { "A", "B" }, new[] { new OfficeChartSeries("Values", new[] { 3d, 4d }) });
}
