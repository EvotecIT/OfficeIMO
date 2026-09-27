using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;

namespace OfficeIMO.Tests;

public sealed class WordChartWorkbookBindingsTests {
    [Fact]
    public void SharedUpdate_RejectsUnqualifiedWorkbookExtensionBeforeChangingChartOrWorkbook() {
        using WordDocument document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Line, Data());
        var extension = new C.ChartSpaceExtension { Uri = "urn:producer:label-range" };
        var range = new DocumentFormat.OpenXml.OpenXmlUnknownElement("p", "labelRange", "urn:producer:labels");
        range.InnerXml = "<p:f xmlns:p=\"urn:producer:labels\">Sheet1!$Z$1:$Z$2</p:f>";
        extension.Append(range);
        chart.ChartPart!.ChartSpace!.AddChild(new C.ChartSpaceExtensionList(extension), true);
        string before = chart.ChartPart.ChartSpace.OuterXml;
        byte[] workbookBefore = Workbook(chart);
        Assert.Throws<System.NotSupportedException>(() => chart.SetData(OfficeChartKind.Line, Data()));
        Assert.Equal(before, chart.ChartPart.ChartSpace.OuterXml);
        Assert.Equal(workbookBefore, Workbook(chart));
    }

    [Fact]
    public void SharedUpdate_MaterializesWorkbookLinkedCustomLabelsAndErrorBars() {
        using WordDocument document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Line, Data());
        var series = chart.ChartPart!.ChartSpace!.Descendants<C.LineChartSeries>().Single();
        series.AddChild(new C.DataLabels(new C.DataLabel(new C.Index { Val = 0 }, Title("Custom label").ChartText!.CloneNode(true))), true);
        series.AddChild(new C.ErrorBars(new C.ErrorBarType { Val = C.ErrorBarValues.Both },
            new C.ErrorBarValueType { Val = C.ErrorValues.Custom }, new C.Plus(new C.NumberReference(
                new C.Formula { Text = "Sheet1!$Z$2:$Z$3" }, new C.NumberingCache(new C.FormatCode { Text = "General" },
                    new C.PointCount { Val = 2 }, new C.NumericPoint { Index = 0, NumericValue = new C.NumericValue { Text = "0.5" } },
                    new C.NumericPoint { Index = 1, NumericValue = new C.NumericValue { Text = "1.5" } })))), true);
        Assert.Empty(document.ValidateDocument());
        chart.SetData(OfficeChartKind.Line, Data());
        series = chart.ChartPart.ChartSpace.Descendants<C.LineChartSeries>().Single();
        Assert.Equal("Custom label", series.Descendants<C.DataLabel>().Single().GetFirstChild<C.ChartText>()!.InnerText);
        var plus = series.Descendants<C.Plus>().Single();
        Assert.Null(plus.GetFirstChild<C.NumberReference>());
        Assert.Equal(new[] { "0.5", "1.5" }, plus.Descendants<C.NumericPoint>().Select(point => point.NumericValue!.Text));
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void SharedUpdate_RejectsUncachedLinkedTextBeforeChangingChartOrWorkbook() {
        using WordDocument document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Line, Data());
        var native = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!;
        native.AddChild(Title("Cached title"), true);
        native.PlotArea!.Elements<C.ValueAxis>().Single().AddChild(new C.Title(new C.ChartText(new C.StringReference(
            new C.Formula { Text = "Sheet1!$Z$1" }))), true);
        string before = chart.ChartPart.ChartSpace.OuterXml;
        byte[] workbookBefore = Workbook(chart);
        Assert.Throws<System.NotSupportedException>(() => chart.SetData(OfficeChartKind.Line, Data()));
        Assert.Equal(before, chart.ChartPart.ChartSpace.OuterXml);
        Assert.Equal(workbookBefore, Workbook(chart));
    }

    [Theory]
    [InlineData(OfficeChartKind.Scatter, OfficeChartKind.Line)]
    [InlineData(OfficeChartKind.Line, OfficeChartKind.Scatter)]
    public void SharedUpdate_RejectsNumericFamilyOverridesBeforeAddingDrawing(OfficeChartKind defaultKind, OfficeChartKind renderKind) {
        using WordDocument document = WordDocument.Create();
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Status", new[] { 1d, 2d },
            new[] { 1d, 2d }, null, null, true, renderKind: renderKind) });
        Assert.Throws<System.NotSupportedException>(() => document.AddChart(defaultKind, data));
        Assert.Empty(document.Charts);
    }

    [Fact]
    public void SharedUpdate_RebuildsScatterWhenAnUnsupportedNativeLayerIsPresent() {
        using WordDocument document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Scatter, Data());
        var pie = document.AddChart(OfficeChartKind.Pie, Data());
        var plot = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        var sourcePie = pie.ChartPart!.ChartSpace!.Descendants<C.PieChart>().Single();
        var extra = new C.Pie3DChart(sourcePie.Elements<C.PieChartSeries>().Single().CloneNode(true));
        plot.InsertBefore(extra, plot.Elements<C.ValueAxis>().First());
        Assert.Empty(document.ValidateDocument());
        chart.SetData(OfficeChartKind.Scatter, Data());
        plot = chart.ChartPart.ChartSpace.GetFirstChild<C.Chart>()!.PlotArea!;
        Assert.Single(plot.Elements<C.ScatterChart>());
        Assert.Empty(plot.Elements<C.Pie3DChart>());
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void SharedUpdate_MaterializesCachedWorkbookLinkedChartAndAxisTitles() {
        using WordDocument document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Line, Data());
        var native = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!;
        native.AddChild(Title("Original title"), true);
        native.PlotArea!.Elements<C.ValueAxis>().Single().AddChild(Title("Original axis"), true);
        Assert.Empty(document.ValidateDocument());
        chart.SetData(OfficeChartKind.Line, Data());
        Assert.Equal("Original title", native.Title!.InnerText);
        Assert.Equal("Original axis", native.PlotArea!.Elements<C.ValueAxis>().Single().GetFirstChild<C.Title>()!.InnerText);
        Assert.All(native.Descendants<C.Title>(), title => Assert.Empty(title.Descendants<C.StringReference>()));
        Assert.Empty(document.ValidateDocument());
        using var bytes = new MemoryStream();
        document.Save(bytes); bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        Assert.Equal("Original title", reopened.Charts.First().Title);
        Assert.Empty(reopened.ValidateDocument());
    }

    private static C.Title Title(string text) => new C.Title(new C.ChartText(new C.StringReference(
        new C.Formula { Text = "Sheet1!$Z$1" }, new C.StringCache(new C.PointCount { Val = 1 },
            new C.StringPoint { Index = 0, NumericValue = new C.NumericValue { Text = text } }))));

    private static OfficeChartData Data() => new OfficeChartData(new[] { "A", "B" },
        new[] { new OfficeChartSeries("Status", new[] { 1d, 2d }, new[] { 1d, 2d }) });

    private static byte[] Workbook(WordChart chart) {
        using var output = new MemoryStream();
        using (Stream source = chart.ChartPart!.GetPartsOfType<EmbeddedPackagePart>().Single().GetStream()) source.CopyTo(output);
        return output.ToArray();
    }
}
