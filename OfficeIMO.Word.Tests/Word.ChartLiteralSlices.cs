using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordChartLiteralSliceTests {
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void SliceAppend_CreatesMissingDataSourcesBeforeSeriesExtensions(int family) {
        using WordDocument document = WordDocument.Create();
        WordChart chart = document.AddChart();
        Append(chart, family, "A", 8);
        PieChartSeries series = chart.ChartPart!.ChartSpace.Descendants<PieChartSeries>().Single();
        series.RemoveAllChildren<CategoryAxisData>();
        series.RemoveAllChildren<Values>();
        series.Append(new PieSerExtensionList());
        Assert.Empty(document.ValidateDocument());
        using var stream = new MemoryStream();
        document.Save(stream);
        stream.Position = 0;
        using WordDocument reopened = WordDocument.Load(stream);
        chart = reopened.Charts.Single();
        Append(chart, family, "B", 2);
        series = chart.ChartPart!.ChartSpace.Descendants<PieChartSeries>().Single();
        Assert.Equal("extLst", series.LastChild!.LocalName);
        Assert.Equal("B", series.GetFirstChild<CategoryAxisData>()!.InnerText);
        Assert.Equal("2", series.GetFirstChild<Values>()!.Descendants<NumericValue>().Single().Text);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void SliceAppend_RejectsAdditionalPlotGroupsBeforeMutation(int family) {
        using WordDocument document = WordDocument.Create();
        WordChart chart = document.AddChart();
        Append(chart, family, "A", 8);
        PlotArea plot = chart.ChartPart!.ChartSpace.GetFirstChild<Chart>()!.PlotArea!;
        OpenXmlCompositeElement extra = family == 0 ? new DoughnutChart(new HoleSize { Val = (byte)50 }) : new PieChart();
        plot.Append(extra);
        string before = chart.ChartPart.ChartSpace.OuterXml;
        Assert.Throws<NotSupportedException>(() => Append(chart, family, "B", 2));
        Assert.Equal(before, chart.ChartPart.ChartSpace.OuterXml);
    }

    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(0, true)]
    [InlineData(1, true)]
    [InlineData(2, true)]
    public void SliceAppend_PreservesSparsePositionsNumericCategoriesAndDisabledLabels(int family, bool numeric) {
        using WordDocument document = WordDocument.Create();
        WordChart chart = document.AddChart();
        Append(chart, family, "1", 8);
        Append(chart, family, "2", 2);
        Append(chart, family, "3", 1);
        Chart native = chart.ChartPart!.ChartSpace.GetFirstChild<Chart>()!;
        var layer = (OpenXmlCompositeElement)native.PlotArea!.ChildElements.First(element =>
            element is PieChart || element is Pie3DChart || element is DoughnutChart);
        layer.RemoveAllChildren<DataLabels>();
        PieChartSeries series = layer.GetFirstChild<PieChartSeries>()!;
        StringLiteral labels = series.GetFirstChild<CategoryAxisData>()!.GetFirstChild<StringLiteral>()!;
        labels.Elements<StringPoint>().Single(point => point.Index!.Value == 1).Remove();
        if (numeric) {
            var numbers = new NumberLiteral(new FormatCode { Text = "General" }, new PointCount { Val = 3 });
            foreach (StringPoint point in labels.Elements<StringPoint>())
                numbers.Append(new NumericPoint { Index = point.Index!.Value, NumericValue = new NumericValue { Text = point.NumericValue!.Text } });
            series.GetFirstChild<CategoryAxisData>()!.ReplaceChild(numbers, labels);
        }
        NumberLiteral values = series.GetFirstChild<Values>()!.GetFirstChild<NumberLiteral>()!;
        values.Elements<NumericPoint>().Single(point => point.Index!.Value == 1).Remove();
        // Declared trailing gaps are also logical positions and must survive the append.
        OpenXmlCompositeElement categories = numeric ? series.GetFirstChild<CategoryAxisData>()!.GetFirstChild<NumberLiteral>()! : labels;
        uint next = numeric ? 3U : 5U;
        if (numeric) {
            categories.RemoveAllChildren<PointCount>();
            values.RemoveAllChildren<PointCount>();
        } else {
            categories.GetFirstChild<PointCount>()!.Val = next;
            values.GetFirstChild<PointCount>()!.Val = next;
        }
        categories.Append(numeric ? (OpenXmlElement)new ExtensionList() : new StrDataExtensionList());
        values.Append(new ExtensionList());
        using var stream = new MemoryStream();
        document.Save(stream);
        stream.Position = 0;
        using WordDocument imported = WordDocument.Load(stream);
        chart = imported.Charts.Single();
        native = chart.ChartPart!.ChartSpace.GetFirstChild<Chart>()!;
        layer = (OpenXmlCompositeElement)native.PlotArea!.ChildElements.First(element =>
            element is PieChart || element is Pie3DChart || element is DoughnutChart);
        series = layer.GetFirstChild<PieChartSeries>()!;
        categories = numeric ? series.GetFirstChild<CategoryAxisData>()!.GetFirstChild<NumberLiteral>()! :
            series.GetFirstChild<CategoryAxisData>()!.GetFirstChild<StringLiteral>()!;
        values = series.GetFirstChild<Values>()!.GetFirstChild<NumberLiteral>()!;
        Append(chart, family, "6", 4);
        Assert.Equal(next + 1, categories.GetFirstChild<PointCount>()!.Val!.Value);
        Assert.Equal(next + 1, values.GetFirstChild<PointCount>()!.Val!.Value);
        uint[] categoryIndexes = numeric
            ? categories.Elements<NumericPoint>().Select(point => point.Index!.Value).ToArray()
            : categories.Elements<StringPoint>().Select(point => point.Index!.Value).ToArray();
        Assert.Equal(new uint[] { 0, 2, next }, categoryIndexes);
        Assert.Equal(new uint[] { 0, 2, next }, values.Elements<NumericPoint>().Select(point => point.Index!.Value));
        Assert.Equal("4", values.Elements<NumericPoint>().Last().NumericValue!.Text);
        Assert.Equal("extLst", categories.LastChild!.LocalName);
        Assert.Equal("extLst", values.LastChild!.LocalName);
        Assert.Null(layer.GetFirstChild<DataLabels>());
        Assert.Empty(imported.ValidateDocument());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void SliceCreation_PreservesDefaultValueAndPercentLabels(int family) {
        using WordDocument document = WordDocument.Create();
        WordChart chart = document.AddChart();
        Append(chart, family, "A", 8);
        DataLabels labels = chart.ChartPart!.ChartSpace.Descendants<DataLabels>().Single();
        Assert.True(labels.GetFirstChild<ShowValue>()!.Val!.Value);
        Assert.True(labels.GetFirstChild<ShowPercent>()!.Val!.Value);
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void SliceAppend_RejectsMisalignedPositionsBeforeMutatingNativeData(int family) {
        using WordDocument document = WordDocument.Create();
        WordChart chart = document.AddChart();
        Append(chart, family, "A", 8);
        Append(chart, family, "B", 2);
        PieChartSeries series = chart.ChartPart!.ChartSpace.Descendants<PieChartSeries>().Single();
        series.GetFirstChild<Values>()!.Descendants<NumericPoint>().Last().Index = 3U;
        string before = chart.ChartPart.ChartSpace.OuterXml;
        Assert.Throws<InvalidOperationException>(() => Append(chart, family, "C", 1));
        Assert.Equal(before, chart.ChartPart.ChartSpace.OuterXml);
    }

    private static void Append(WordChart chart, int family, string category, int value) {
        if (family == 0) chart.AddPie(category, value);
        else if (family == 1) chart.AddPie3D(category, value);
        else chart.AddDoughnut(category, value);
    }
}
