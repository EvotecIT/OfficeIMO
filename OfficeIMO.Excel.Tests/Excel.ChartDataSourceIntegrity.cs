using DocumentFormat.OpenXml;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class ExcelChartDataSourceIntegrityTests {
    [Theory]
    [InlineData(ExcelChartType.ColumnClustered)]
    [InlineData(ExcelChartType.Line)]
    [InlineData(ExcelChartType.Area)]
    [InlineData(ExcelChartType.Pie)]
    public void Update_ReplacesNumericCategoriesAndRestoresMissingValuesInSchemaOrder(ExcelChartType kind) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelChart chart = document.AddWorksheet("Summary").AddChart(Data(kind), row: 1, column: 6, type: kind);
        var part = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single(p => p.DrawingsPart != null).DrawingsPart!.ChartParts.Single();
        var series = part.ChartSpace.Descendants<OpenXmlCompositeElement>().Single(e => e.LocalName == "ser");
        series.AddChild(new C.SeriesText(new C.NumericValue { Text = "Literal title" }), true);
        series.AddChild(new C.CategoryAxisData(new C.NumberLiteral(new C.FormatCode { Text = "General" },
            new C.PointCount { Val = 2 }, new C.NumericPoint { Index = 0, NumericValue = new C.NumericValue { Text = "1" } },
            new C.NumericPoint { Index = 1, NumericValue = new C.NumericValue { Text = "2" } })), true);
        series.RemoveAllChildren<C.Values>();
        if (kind == ExcelChartType.Line) series.AddChild(new C.Smooth { Val = false }, true);
        Assert.Empty(document.ValidateOpenXml());
        chart.UpdateData(Data(kind));
        Assert.Single(series.GetFirstChild<C.SeriesText>()!.ChildElements);
        Assert.Single(series.GetFirstChild<C.CategoryAxisData>()!.ChildElements);
        Assert.NotNull(series.GetFirstChild<C.CategoryAxisData>()!.GetFirstChild<C.StringReference>());
        Assert.NotNull(series.GetFirstChild<C.Values>()!.GetFirstChild<C.NumberReference>());
        AssertValidAfterReopen(document);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ScatterUpdate_ReplacesStringXChoiceAndRestoresSourcesBeforeTrailingMetadata(bool missingX) {
        using ExcelDocument document = ExcelDocument.Create();
        ExcelChart chart = document.AddWorksheet("Summary").AddChart(Data(ExcelChartType.Scatter), row: 1, column: 6, type: ExcelChartType.Scatter);
        var part = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single(p => p.DrawingsPart != null).DrawingsPart!.ChartParts.Single();
        var series = part.ChartSpace.Descendants<C.ScatterChartSeries>().Single();
        series.AddChild(new C.SeriesText(new C.NumericValue { Text = "Literal title" }), true);
        series.RemoveAllChildren<C.XValues>();
        if (!missingX) series.AddChild(new C.XValues(new C.StringLiteral(new C.PointCount { Val = 2 },
            new C.StringPoint { Index = 0, NumericValue = new C.NumericValue { Text = "1" } },
            new C.StringPoint { Index = 1, NumericValue = new C.NumericValue { Text = "2" } })), true);
        series.RemoveAllChildren<C.YValues>();
        series.AddChild(new C.Smooth { Val = false }, true);
        series.AddChild(new C.ScatterSerExtensionList(), true);
        Assert.Empty(document.ValidateOpenXml());
        chart.UpdateData(Data(ExcelChartType.Scatter));
        Assert.Single(series.GetFirstChild<C.SeriesText>()!.ChildElements);
        Assert.Single(series.GetFirstChild<C.XValues>()!.ChildElements);
        Assert.NotNull(series.GetFirstChild<C.XValues>()!.GetFirstChild<C.NumberReference>());
        Assert.NotNull(series.GetFirstChild<C.YValues>()!.GetFirstChild<C.NumberReference>());
        Assert.Equal("extLst", series.LastChild!.LocalName);
        AssertValidAfterReopen(document);
    }

    private static ExcelChartData Data(ExcelChartType kind) => new ExcelChartData(new[] { "1", "2" },
        new[] { new ExcelChartSeries("Updated", new[] { 5d, 6d }, kind) });

    private static void AssertValidAfterReopen(ExcelDocument document) {
        Assert.Empty(document.ValidateOpenXml());
        using var stream = new MemoryStream();
        document.Save(stream);
        stream.Position = 0;
        using ExcelDocument reopened = ExcelDocument.Load(stream);
        Assert.Empty(reopened.ValidateOpenXml());
    }
}
