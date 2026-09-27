using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class WordChartProjectionQualificationTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Snapshot_RejectsHiddenWorkbookValuesOnlyWhenVisibleOnlyIsEnabled(bool visibleOnly) {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartKind.ColumnClustered, new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Values", new[] { 3d, 4d }) }));
        var part = chart.ChartPart!;
        part.ChartSpace!.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotVisibleOnly>()!.Val = visibleOnly;
        var embedded = part.GetPartsOfType<DocumentFormat.OpenXml.Packaging.EmbeddedPackagePart>().Single();
        using var bytes = new MemoryStream();
        using (var source = embedded.GetStream()) source.CopyTo(bytes);
        bytes.Position = 0;
        using (var workbook = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(bytes, true)) {
            workbook.WorkbookPart!.WorksheetParts.Single().Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Row>().Skip(1).First().Hidden = true;
        }
        bytes.Position = 0; embedded.FeedData(bytes);
        Assert.Equal(!visibleOnly, chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData("trendline")]
    [InlineData("errorBars")]
    [InlineData("dropLines")]
    [InlineData("sparse")]
    [InlineData("negative")]
    [InlineData("dateAxis")]
    [InlineData("style")]
    [InlineData("labelFormat")]
    [InlineData("legendEntry")]
    [InlineData("rotation")]
    [InlineData("labelSkip")]
    [InlineData("unequal")]
    [InlineData("emptyValues")]
    [InlineData("unequalSeries")]
    [InlineData("differentCategories")]
    [InlineData("titleParagraphs")]
    [InlineData("titleBreak")]
    [InlineData("dataTable")]
    public void Snapshot_RejectsUnrepresentedNativeChartContent(string feature) {
        using var document = WordDocument.Create();
        var kind = feature == "negative" ? OfficeChartKind.ColumnClustered : OfficeChartKind.Line;
        var chart = document.AddChart(kind, new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Values", new[] { 3d, -4d }) }), title: "Revenue");
        var space = chart.ChartPart!.ChartSpace!;
        var native = space.GetFirstChild<C.Chart>()!;
        var plot = native.PlotArea!;
        var layer = plot.ChildElements.OfType<OpenXmlCompositeElement>().First(item => item.LocalName.EndsWith("Chart", StringComparison.Ordinal));
        var series = layer.ChildElements.OfType<OpenXmlCompositeElement>().Single(item => item.LocalName == "ser");
        if (feature == "trendline") series.AddChild(new C.Trendline(new C.TrendlineType { Val = C.TrendlineValues.Linear }), true);
        else if (feature == "errorBars") series.AddChild(new C.ErrorBars(), true);
        else if (feature == "dropLines") layer.AddChild(new C.DropLines(), true);
        else if (feature == "sparse") series.Descendants<C.NumericPoint>().Last().Remove();
        else if (feature == "negative") series.AddChild(new C.InvertIfNegative { Val = true }, true);
        else if (feature == "style") space.AddChild(new C.Style { Val = 42 }, true);
        else if (feature == "labelFormat") layer.AddChild(new C.DataLabels(new C.NumberingFormat { FormatCode = "m/d/yyyy", SourceLinked = false }, new C.ShowValue { Val = true }), true);
        else if (feature == "legendEntry") native.GetFirstChild<C.Legend>()!.AddChild(new C.LegendEntry(new C.Index { Val = 0 }, new C.TextProperties(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.ParagraphProperties(new A.DefaultRunProperties { FontSize = 2400 })))), true);
        else if (feature == "rotation") space.AddChild(new C.TextProperties(new A.BodyProperties { Rotation = 5400000 }, new A.ListStyle(), new A.Paragraph()), true);
        else if (feature == "labelSkip") plot.GetFirstChild<C.CategoryAxis>()!.AddChild(new C.TickLabelSkip { Val = 2 }, true);
        else if (feature == "unequal") { var cache = series.Descendants<C.Values>().Single(); cache.Descendants<C.NumericPoint>().Last().Remove(); cache.Descendants<C.PointCount>().Single().Val = 1; }
        else if (feature == "emptyValues") series.GetFirstChild<C.Values>()!.Remove();
        else if (feature == "unequalSeries" || feature == "differentCategories") {
            var sibling = (OpenXmlCompositeElement)series.CloneNode(true);
            sibling.GetFirstChild<C.Index>()!.Val = 1;
            sibling.GetFirstChild<C.Order>()!.Val = 1;
            if (feature == "unequalSeries") {
                sibling.Descendants<C.NumericPoint>().Last().Remove();
                sibling.Descendants<C.Values>().Single().Descendants<C.PointCount>().Single().Val = 1;
            } else sibling.GetFirstChild<C.CategoryAxisData>()!.Descendants<C.StringPoint>().Last().NumericValue!.Text = "Different category";
            layer.AddChild(sibling, true);
        }
        else if (feature == "titleParagraphs") native.GetFirstChild<C.Title>()!.Descendants<C.RichText>().Single().Append(new A.Paragraph(new A.Run(new A.Text("2026"))));
        else if (feature == "titleBreak") native.GetFirstChild<C.Title>()!.Descendants<A.Paragraph>().Single().Append(new A.Break(), new A.Run(new A.Text("2026")));
        else if (feature == "dataTable") plot.AddChild(new C.DataTable(new C.ShowHorizontalBorder { Val = true }), true);
        else if (feature == "dateAxis") {
            var category = plot.GetFirstChild<C.CategoryAxis>()!;
            var replacement = new C.DateAxis();
            foreach (var child in category.ChildElements) replacement.Append(child.CloneNode(true));
            plot.ReplaceChild(replacement, category);
        }
        string before = space.OuterXml;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        Assert.Equal(before, space.OuterXml);
    }
}
