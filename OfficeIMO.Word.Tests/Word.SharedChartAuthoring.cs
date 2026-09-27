using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests;

public sealed class WordSharedChartAuthoringTests {
    public static IEnumerable<object[]> SupportedKinds() =>
        Enum.GetValues(typeof(OfficeChartKind)).Cast<OfficeChartKind>().Select(kind => new object[] { kind });

    [Theory]
    [MemberData(nameof(SupportedKinds))]
    public void SharedChart_AuthorsReopensAndUpdatesNativeCachesAndEmbeddedWorkbook(OfficeChartKind kind) {
        using WordDocument document = WordDocument.Create();
        WordChart chart = document.AddChart(kind, Data(kind, 0), "Results", roundedCorners: true, width: 360, height: 180);
        chart.Name = "Status chart";
        chart.AltText = "Status distribution";
        Assert.Empty(document.ValidateDocument());
        using var bytes = new MemoryStream();
        document.Save(bytes);
        bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        WordChart restored = reopened.Charts.Single();
        restored.SetData(kind, Data(kind, 10));
        Assert.Equal("Status chart", restored.Name);
        Assert.Equal("Status distribution", restored.AltText);
        Assert.True(restored.TryGetLayoutSnapshot(out WordDrawingLayoutSnapshot layout));
        Assert.Equal(270D, layout.WidthPoints);
        Assert.Equal(135D, layout.HeightPoints);
        C.ChartSpace space = restored.ChartPart!.ChartSpace!;
        Assert.True(space.GetFirstChild<C.RoundedCorners>()!.Val!.Value);
        Assert.Contains("Results", space.GetFirstChild<C.Chart>()!.Title!.InnerText);
        C.Values? values = space.Descendants<C.Values>().FirstOrDefault();
        var points = values == null ? space.Descendants<C.YValues>().Single().Descendants<C.NumericPoint>() :
            values.Descendants<C.NumericPoint>();
        Assert.Equal(new[] { "11", "12", "13" }, points.Select(point => point.NumericValue!.Text));
        EmbeddedPackagePart embedded = restored.ChartPart.GetPartsOfType<EmbeddedPackagePart>().Single();
        Assert.Equal(restored.ChartPart.GetIdOfPart(embedded), space.GetFirstChild<C.ExternalData>()!.Id!.Value);
        using var workbookBytes = new MemoryStream();
        using (Stream source = embedded.GetStream()) source.CopyTo(workbookBytes);
        workbookBytes.Position = 0;
        using SpreadsheetDocument workbook = SpreadsheetDocument.Open(workbookBytes, false);
        Assert.Empty(new OpenXmlValidator().Validate(workbook));
        S.Cell lastValue = workbook.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<S.Cell>()
            .Single(cell => cell.CellReference!.Value == "B4");
        Assert.Equal("13", lastValue.CellValue!.Text);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Fact]
    public void SharedChart_ParagraphEntryPointPreflightsBeforeInsertingAndPreservesText() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("Introduction");
        int count = document.Paragraphs.Count;
        Assert.Throws<ArgumentOutOfRangeException>(() => paragraph.AddChart(OfficeChartKind.Pie, Data(OfficeChartKind.Pie, 0), width: 0));
        Assert.Equal(count, document.Paragraphs.Count);
        paragraph.AddChart(OfficeChartKind.Pie, Data(OfficeChartKind.Pie, 0));
        Assert.Single(document.Charts);
        Assert.Equal("Introduction", paragraph.Text);
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void SharedChart_ChangesFamiliesWithoutLeavingOldPlotLayers() {
        using WordDocument document = WordDocument.Create();
        WordChart chart = document.AddChart(OfficeChartKind.Pie, Data(OfficeChartKind.Pie, 0));
        foreach (OfficeChartKind kind in new[] { OfficeChartKind.Scatter, OfficeChartKind.Bubble, OfficeChartKind.Doughnut }) {
            chart.SetData(kind, Data(kind, 10));
            C.PlotArea plot = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
            Assert.Single(plot.ChildElements, element => element.LocalName.EndsWith("Chart", StringComparison.Ordinal));
            Assert.Equal(kind == OfficeChartKind.Scatter ? "scatterChart" : kind == OfficeChartKind.Bubble ? "bubbleChart" : "doughnutChart",
                plot.ChildElements.Single(element => element.LocalName.EndsWith("Chart", StringComparison.Ordinal)).LocalName);
            Assert.Empty(document.ValidateDocument());
        }
    }

    [Fact]
    public void SharedChart_AuthorsAndUpdatesComboAxesAndExplicitPointStyleReset() {
        using WordDocument document = WordDocument.Create();
        OfficeChartData Combo(int offset) => new OfficeChartData(new[] { "A", "B", "C" }, new[] {
            new OfficeChartSeries("Count", new[] { 1d + offset, 2d + offset, 3d + offset }, null, null, null, false,
                renderKind: OfficeChartKind.ColumnClustered),
            new OfficeChartSeries("Rate", new[] { 10d + offset, 20d + offset, 30d + offset }, null, null, null, true,
                renderKind: OfficeChartKind.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        });
        WordChart combo = document.AddChart(OfficeChartKind.ColumnClustered, Combo(0));
        combo.SetData(OfficeChartKind.ColumnClustered, Combo(10));
        C.PlotArea plot = combo.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        Assert.Single(plot.Elements<C.BarChart>());
        Assert.Single(plot.Elements<C.LineChart>());
        Assert.Equal(2, plot.Elements<C.ValueAxis>().Count());
        Assert.Equal(new[] { "20", "30", "40" }, plot.Descendants<C.LineChartSeries>().Single()
            .Descendants<C.NumericPoint>().Select(point => point.NumericValue!.Text));

        var styled = new OfficeChartSeries("Status", new[] { 1d, 2d, 3d }).WithPointStyles(new OfficeChartPointStyle?[] {
            new OfficeChartPointStyle(fillColor: OfficeColor.Parse("#008000")),
            new OfficeChartPointStyle(noFill: true, outlineColor: OfficeColor.Black, outlineWidth: 2),
            new OfficeChartPointStyle(hatch: OfficeChartHatchPattern.Cross, hatchColor: OfficeColor.Black)
        });
        WordChart pie = document.AddChart(OfficeChartKind.Pie, new OfficeChartData(new[] { "A", "B", "C" }, new[] { styled }));
        Assert.Single(pie.ChartPart!.ChartSpace!.Descendants<A.NoFill>(), fill => fill.Parent is C.ChartShapeProperties);
        Assert.Single(pie.ChartPart.ChartSpace.Descendants<A.PatternFill>());
        pie.SetData(OfficeChartKind.Pie, new OfficeChartData(new[] { "A", "B", "C" }, new[] {
            new OfficeChartSeries("Status", new[] { 4d, 5d, 6d }).WithPointStyles(new OfficeChartPointStyle?[3])
        }));
        Assert.Empty(pie.ChartPart.ChartSpace.Descendants<A.PatternFill>());
        Assert.DoesNotContain(pie.ChartPart.ChartSpace.Descendants<A.NoFill>(), fill => fill.Parent is C.ChartShapeProperties);
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void SharedChart_InvalidInputDoesNotAddADrawingOrChangeExistingNativeData() {
        using WordDocument document = WordDocument.Create();
        Assert.Throws<ArgumentOutOfRangeException>(() => document.AddChart((OfficeChartKind)999, Data(OfficeChartKind.Pie, 0)));
        Assert.Empty(document.Charts);
        WordChart chart = document.AddChart(OfficeChartKind.Pie, Data(OfficeChartKind.Pie, 0));
        string before = chart.ChartPart!.ChartSpace!.OuterXml;
        Assert.Throws<NotSupportedException>(() => chart.SetData(OfficeChartKind.Pie,
            new OfficeChartData(new[] { "A" }, new[] {
                new OfficeChartSeries("Secondary", new[] { 1d }, null, null, null, true,
                    axisGroup: OfficeChartAxisGroup.Secondary)
            })));
        Assert.Equal(before, chart.ChartPart.ChartSpace.OuterXml);
    }

    private static OfficeChartData Data(OfficeChartKind kind, int offset) {
        var values = new[] { 1d + offset, 2d + offset, 3d + offset };
        OfficeChartSeries series = kind == OfficeChartKind.Bubble
            ? OfficeChartSeries.CreateBubble("Results", new[] { 1d, 2d, 3d }, values, new[] { 1d, 2d, 3d })
            : new OfficeChartSeries("Results", values, kind == OfficeChartKind.Scatter ? new[] { 1d, 2d, 3d } : null);
        return new OfficeChartData(new[] { "A", "B", "C" }, new[] { series });
    }
}
