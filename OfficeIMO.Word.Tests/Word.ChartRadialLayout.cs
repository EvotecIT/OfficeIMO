using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordChartRadialLayoutTests {
    [Theory]
    [InlineData("361", "50")]
    [InlineData("0", "9")]
    [InlineData("invalid", "50")]
    [InlineData("9999999999999", "50")]
    public void RadialGeometry_MalformedChartWarnsWithoutAbortingWordPdf(string angle, string hole) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Surrounding document content");
        WordChart chart = document.AddChart(OfficeChartKind.Doughnut,
            new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Status", new[] { 1d }) }))
            .SetRadialLayout(OfficeChartRadialLayout.Default);
        var native = chart.ChartPart!.ChartSpace!.Descendants<C.DoughnutChart>().Single();
        native.GetFirstChild<C.FirstSliceAngle>()!.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("", "val", "", angle));
        native.GetFirstChild<C.HoleSize>()!.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("", "val", "", hole));
        var options = new WordToPdfOptions();
        byte[] pdf = document.ToPdfBytes(options);
        Assert.Contains(options.Warnings, warning => warning.Code == "NativeBodyChartUnsupported" && warning.Message.Contains("invalid pie rotation"));
        using var rendered = UglyToad.PdfPig.PdfDocument.Open(pdf);
        Assert.Contains("Surrounding document content", rendered.GetPage(1).Text);
    }

    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void RadialGeometry_PersistsThroughReopenAndDataUpdate(OfficeChartKind kind) {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Status", new[] { 7d, 3d }) });
        using WordDocument authored = WordDocument.Create();
        authored.AddChart(kind, data).SetRadialLayout(new OfficeChartRadialLayout(135, 75));
        using var bytes = new MemoryStream();
        authored.Save(bytes);
        bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        WordChart chart = reopened.Charts.Single();
        chart.SetData(kind, data);
        Assert.True(chart.TryGetSnapshot(out WordChartSnapshot snapshot));
        Assert.Equal(135, snapshot.RadialLayout.FirstSliceAngleDegrees);
        Assert.Equal(kind == OfficeChartKind.Doughnut ? 75 : 50, snapshot.RadialLayout.DoughnutHolePercent);
        Assert.Empty(reopened.ValidateDocument());
    }
}
