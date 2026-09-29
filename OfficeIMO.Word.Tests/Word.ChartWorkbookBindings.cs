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
    public void SharedPieAuthoringRejectsMultipleSeriesInsteadOfSilentlyOmittingOneInSnapshots() {
        using WordDocument document = WordDocument.Create();
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("First", new[] { 1d, 2d }),
            new OfficeChartSeries("Second", new[] { 3d, 4d }) });
        Assert.Throws<System.NotSupportedException>(() => document.AddChart(OfficeChartKind.Pie, data));
        var doughnut = document.AddChart(OfficeChartKind.Doughnut, data);
        Assert.Equal(2, doughnut.ChartPart!.ChartSpace!.Descendants<C.PieChartSeries>().Count());
    }

    [Fact]
    public void ScatterWorkbook_StoresOnlyActualPointsInUnequalLengthSeries() {
        var data = new OfficeChartData(new[] { "1", "2", "3" }, new[] {
            new OfficeChartSeries("Long", new[] { 4d, 5d, 6d }, new[] { 1d, 2d, 3d }),
            new OfficeChartSeries("Short", new[] { 7d }, new[] { 8d }) });
        byte[] bytes = OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartWriter.BuildScatterWorkbook(data);
        using var stream = new MemoryStream(bytes);
        using var workbook = SpreadsheetDocument.Open(stream, false);
        var rows = workbook.WorkbookPart!.WorksheetParts.Single().Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>()!
            .Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().ToArray();
        Assert.Equal(new[] { 4, 4, 2, 2 }, rows.Select(row => row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Count()));
        Assert.Equal(new[] { "A2", "B2", "C2", "D2" }, rows[1].Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Select(cell => cell.CellReference!.Value));
        Assert.Equal(new[] { "A4", "B4" }, rows[3].Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Select(cell => cell.CellReference!.Value));
        Assert.Equal("6", rows[3].Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Last().CellValue!.Text);
    }
    [Fact]
    public void SharedUpdate_PreservesNativeCombinationLayerOrder() {
        using var document = WordDocument.Create();
        var data = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Columns", new[] { 3d, 4d }),
            new OfficeChartSeries("Line", new[] { 1d, 2d }, null, null, null, true, renderKind: OfficeChartKind.Line) });
        var chart = document.AddChart(OfficeChartKind.ColumnClustered, data);
        var plot = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        var line = plot.GetFirstChild<C.LineChart>()!;
        line.Remove(); plot.InsertBefore(line, plot.GetFirstChild<C.BarChart>());
        chart.SetData(OfficeChartKind.ColumnClustered, data);
        Assert.Equal(new[] { "lineChart", "barChart" }, chart.ChartPart.ChartSpace.GetFirstChild<C.Chart>()!.PlotArea!.ChildElements
            .Where(element => element.LocalName.EndsWith("Chart", System.StringComparison.Ordinal)).Select(element => element.LocalName));
        Assert.Empty(document.ValidateDocument());
    }
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void SharedUpdate_RejectsIncompleteErrorBarCachesBeforeMutation(int missing) {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Line, Data());
        var cache = new C.NumberingCache(new C.PointCount { Val = 2 },
            new C.NumericPoint { Index = 0, NumericValue = new C.NumericValue { Text = "0.5" } },
            new C.NumericPoint { Index = 1, NumericValue = new C.NumericValue { Text = "1.5" } });
        if (missing == 0) cache.Elements<C.NumericPoint>().Last().Remove();
        else if (missing == 1) cache.Elements<C.NumericPoint>().Last().NumericValue!.Remove();
        else cache.Elements<C.NumericPoint>().Last().NumericValue!.Text = "NaN";
        chart.ChartPart!.ChartSpace!.Descendants<C.LineChartSeries>().Single().AddChild(new C.ErrorBars(
            new C.ErrorBarType { Val = C.ErrorBarValues.Both }, new C.ErrorBarValueType { Val = C.ErrorValues.Custom },
            new C.Plus(new C.NumberReference(new C.Formula { Text = "Sheet1!$Z$2:$Z$3" }, cache))), true);
        string before = chart.ChartPart.ChartSpace.OuterXml;
        byte[] workbook = Workbook(chart);
        Assert.Throws<System.NotSupportedException>(() => chart.SetData(OfficeChartKind.Line, Data()));
        Assert.Equal(before, chart.ChartPart.ChartSpace.OuterXml);
        Assert.Equal(workbook, Workbook(chart));
    }

    [Theory]
    [InlineData(1)]
    [InlineData(3)]
    public void SharedUpdate_RejectsErrorBarCachesWithDifferentReplacementLengthBeforeMutation(int replacementCount) {
        using WordDocument document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Line, Data());
        chart.ChartPart!.ChartSpace!.Descendants<C.LineChartSeries>().Single().AddChild(new C.ErrorBars(
            new C.ErrorBarType { Val = C.ErrorBarValues.Both }, new C.ErrorBarValueType { Val = C.ErrorValues.Custom },
            new C.Plus(new C.NumberReference(new C.Formula { Text = "Sheet1!$Z$2:$Z$3" },
                new C.NumberingCache(new C.PointCount { Val = 2 },
                    new C.NumericPoint { Index = 0, NumericValue = new C.NumericValue { Text = "0.5" } },
                    new C.NumericPoint { Index = 1, NumericValue = new C.NumericValue { Text = "1.5" } })))), true);
        string before = chart.ChartPart.ChartSpace.OuterXml;
        byte[] workbook = Workbook(chart);
        var replacement = new OfficeChartData(Enumerable.Range(0, replacementCount).Select(index => $"C{index}"),
            new[] { new OfficeChartSeries("Updated", Enumerable.Range(0, replacementCount).Select(index => (double)index)) });
        Assert.Throws<System.NotSupportedException>(() => chart.SetData(OfficeChartKind.Line, replacement));
        Assert.Equal(before, chart.ChartPart.ChartSpace.OuterXml);
        Assert.Equal(workbook, Workbook(chart));
    }

    [Fact]
    public void ScatterGrowth_PreservesExistingPointMetadataWithoutCopyingItToNewSeries() {
        using var document = WordDocument.Create();
        var styled = Data().Series.Single().WithPointStyles(new OfficeChartPointStyle?[] { new(OfficeColor.White), null });
        var chart = document.AddChart(OfficeChartKind.Scatter, new OfficeChartData(Data().Categories, new[] { styled }));
        var source = chart.ChartPart!.ChartSpace!.Descendants<C.ScatterChartSeries>().Single();
        source.AddChild(new C.DataLabels(new C.DataLabel(new C.Index { Val = 0 },
            (C.ChartText)Title("Original point").GetFirstChild<C.ChartText>()!.CloneNode(true)), new C.ShowValue { Val = false }), true);
        Assert.Empty(document.ValidateDocument());
        chart.SetData(OfficeChartKind.Scatter, new OfficeChartData(Data().Categories, new[] { styled,
            new OfficeChartSeries("Added", new[] { 3d, 4d }, new[] { 1d, 2d }) }));
        var series = chart.ChartPart.ChartSpace.Descendants<C.ScatterChartSeries>().ToArray();
        Assert.Single(series[0].Elements<C.DataPoint>());
        Assert.Single(series[0].Elements<C.DataLabels>());
        Assert.Empty(series[1].Elements<C.DataPoint>());
        Assert.Empty(series[1].Elements<C.DataLabels>());
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void ScatterGrowth_PreservesExistingErrorBarsWithoutCopyingMagnitudeDataToNewSeries() {
        using WordDocument document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.Scatter, Data());
        var source = chart.ChartPart!.ChartSpace!.Descendants<C.ScatterChartSeries>().Single();
        source.AddChild(new C.ErrorBars(new C.ErrorDirection { Val = C.ErrorBarDirectionValues.Y },
            new C.ErrorBarType { Val = C.ErrorBarValues.Both }, new C.ErrorBarValueType { Val = C.ErrorValues.Custom },
            new C.Plus(new C.NumberReference(new C.Formula { Text = "Sheet1!$Z$2:$Z$3" },
                new C.NumberingCache(new C.FormatCode { Text = "General" }, new C.PointCount { Val = 2 },
                    new C.NumericPoint { Index = 0, NumericValue = new C.NumericValue { Text = "0.5" } },
                    new C.NumericPoint { Index = 1, NumericValue = new C.NumericValue { Text = "1.5" } })))), true);
        Assert.Empty(document.ValidateDocument());
        chart.SetData(OfficeChartKind.Scatter, new OfficeChartData(Data().Categories, new[] {
            Data().Series.Single(), new OfficeChartSeries("Added", new[] { 3d, 4d }, new[] { 1d, 2d }) }));
        var series = chart.ChartPart.ChartSpace.Descendants<C.ScatterChartSeries>().ToArray();
        Assert.Single(series[0].Elements<C.ErrorBars>());
        Assert.Empty(series[1].Elements<C.ErrorBars>());
        Assert.DoesNotContain(series[0].Descendants<C.NumberReference>(), reference => reference.Parent is C.Plus);
        Assert.Empty(document.ValidateDocument());
    }

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
