using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Tests;

public sealed class WordChartProjectionQualificationTests {
    [Fact]
    public void Snapshot_RejectsUnrepresentedChartTextColorTransform() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartType.ColumnClustered,            new OfficeChartData(new[] { "A" }, new[] { new OfficeSeries("Values", new[] { 3d }) }));
        var space = chart.ChartPart!.ChartSpace!;
        space.AddChild(new C.TextProperties(new A.BodyProperties(), new A.ListStyle(),
            new A.Paragraph(new A.ParagraphProperties(new A.DefaultRunProperties(
                new A.SolidFill(new A.RgbColorModelHex(new A.Gamma()) { Val = "112233" }))))), true);
        string before = space.OuterXml;
        Assert.False(chart.TryGetOfficeSnapshot(out _));
        Assert.Equal(before, space.OuterXml);
    }

    [Fact]
    public void Snapshot_RejectsDifferentChartScriptTypeface() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartType.ColumnClustered,            new OfficeChartData(new[] { "A" }, new[] { new OfficeSeries("Values", new[] { 3d }) }));
        var space = chart.ChartPart!.ChartSpace!;
        space.AddChild(new C.TextProperties(new A.BodyProperties(), new A.ListStyle(),
            new A.Paragraph(new A.ParagraphProperties(new A.DefaultRunProperties(
                new A.LatinFont { Typeface = "Arial" }, new A.EastAsianFont { Typeface = "Yu Gothic" })))), true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void Snapshot_RejectsEmptyNativeLabelSeparator() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartType.Pie,            new OfficeChartData(new[] { "A" }, new[] { new OfficeSeries("Values", new[] { 3d }) }));
        chart.ChartPart!.ChartSpace!.Descendants<C.DataLabels>().Single().AddChild(new C.Separator(string.Empty), true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void Snapshot_RejectsUnresolvedFormulaBasedSeriesName() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartType.ColumnClustered,            new OfficeChartData(new[] { "A" }, new[] { new OfficeSeries("Values", new[] { 3d }) }));
        var text = chart.ChartPart!.ChartSpace!.Descendants<C.BarChartSeries>().Single().GetFirstChild<C.SeriesText>()!;
        text.GetFirstChild<C.StringReference>()!.StringCache!.Remove();
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void Snapshot_TreatsOmittedBarGroupingAsClusteredForOverlap() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartType.ColumnClustered,            new OfficeChartData(new[] { "A" }, new[] { new OfficeSeries("Values", new[] { 3d }) }));
        var bars = chart.ChartPart!.ChartSpace!.Descendants<C.BarChart>().Single();
        bars.GetFirstChild<C.BarGrouping>()!.Remove();
        bars.AddChild(new C.Overlap { Val = 0 }, true);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(OfficeChartType.ColumnClustered, snapshot.ChartKind);
    }

    [Fact]
    public void Snapshot_BoundsBubbleOverridesSeparatelyFromCachedPoints() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartType.Bubble,            new OfficeChartData(new[] { "1" }, new[] { OfficeSeries.CreateBubble("Values", new[] { 1d }, new[] { 3d }, new[] { 5d }) }));
        var series = chart.ChartPart!.ChartSpace!.Descendants<C.BubbleChartSeries>().Single();
        for (int index = 0; index < 10001; index++)
            series.InsertBefore(new C.DataPoint(new C.Index { Val = 99 }), series.GetFirstChild<C.XValues>());
        Assert.True(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void AuthoredBarDirectionChangeKeepsNativeAxesProjectable() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartType.BarClustered,            new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeSeries("Values", new[] { 3D, 4D }) }));
        chart.BarDirection = WordChartBarDirection.Column;
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(OfficeChartType.ColumnClustered, snapshot.ChartKind);
        var plot = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        Assert.Equal(C.AxisPositionValues.Bottom, plot.GetFirstChild<C.CategoryAxis>()!.AxisPosition!.Val!.Value);
        Assert.Equal(C.AxisPositionValues.Left, plot.GetFirstChild<C.ValueAxis>()!.AxisPosition!.Val!.Value);
        chart.BarDirection = WordChartBarDirection.Bar;
        chart.BarDirection = null;
        Assert.True(chart.TryGetOfficeSnapshot(out var omitted));
        Assert.Equal(OfficeChartType.ColumnClustered, omitted.ChartKind);
        Assert.Equal(C.AxisPositionValues.Bottom, plot.GetFirstChild<C.CategoryAxis>()!.AxisPosition!.Val!.Value);
        Assert.Equal(C.AxisPositionValues.Left, plot.GetFirstChild<C.ValueAxis>()!.AxisPosition!.Val!.Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Snapshot_RejectsFormulaLinkedTitleWithoutCachedText(bool axisTitle) {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(new[] { "A" }, new[] { new OfficeSeries("Values", new[] { 3d }) }), new OfficeIMO.Word.WordChartOptions { Title = "Results" });
        if (axisTitle) chart.SetXAxisTitle("Units");
        C.Title title = axisTitle
            ? chart.ChartPart!.ChartSpace!.Descendants<C.CategoryAxis>().Single().GetFirstChild<C.Title>()!
            : chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.GetFirstChild<C.Title>()!;
        title.RemoveAllChildren<C.ChartText>();
        title.AddChild(new C.ChartText(new C.StringReference(new C.Formula("Sheet1!$A$1"))), true);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void Snapshot_RejectsSeparateStackedLayersSharingAnAxisPair() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartType.ColumnStacked,            new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeSeries("First", new[] { 2d, 3d }) }));
        C.PlotArea plot = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!.PlotArea!;
        C.BarChart second = (C.BarChart)plot.GetFirstChild<C.BarChart>()!.CloneNode(true);
        C.BarChartSeries series = second.Elements<C.BarChartSeries>().Single();
        series.GetFirstChild<C.Index>()!.Val = 1U;
        series.GetFirstChild<C.Order>()!.Val = 1U;
        plot.InsertAfter(second, plot.GetFirstChild<C.BarChart>());
        Assert.Equal(2, plot.Elements<C.BarChart>().Count());
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    [InlineData(null)]
    public void Snapshot_LeavesOutHiddenWorkbookRowsUnlessTheChartShowsHiddenCells(bool? visibleOnly) {
        using var document = WordDocument.Create();
        var chart = AddChartWithTwoSeries(document);
        C.Chart native = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!;
        // Office plots visible cells only when c:plotVisOnly is absent.
        if (visibleOnly.HasValue) native.GetFirstChild<C.PlotVisibleOnly>()!.Val = visibleOnly.Value;
        else native.GetFirstChild<C.PlotVisibleOnly>()!.Remove();
        EditEmbeddedWorksheet(chart, worksheet => Row(worksheet, 3).Hidden = true);

        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));

        bool leftOut = visibleOnly != false;
        Assert.Equal(leftOut ? new[] { "A", "C" } : new[] { "A", "B", "C" }, snapshot.Data.Categories);
        Assert.Equal(leftOut ? new[] { 3d, 5d } : new[] { 3d, 4d, 5d }, snapshot.Data.Series[0].Values);
        Assert.Equal(leftOut ? new[] { 6d, 8d } : new[] { 6d, 7d, 8d }, snapshot.Data.Series[1].Values);
    }

    [Fact]
    public void Snapshot_LeavesOutSeriesWhoseWorkbookColumnIsHidden() {
        using var document = WordDocument.Create();
        var chart = AddChartWithTwoSeries(document);
        EditEmbeddedWorksheet(chart, worksheet => {
            var columns = worksheet.GetFirstChild<S.Columns>() ?? worksheet.InsertBefore(new S.Columns(), worksheet.GetFirstChild<S.SheetData>());
            columns.Append(new S.Column { Min = 2U, Max = 2U, Hidden = true });
        });

        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));

        var series = Assert.Single(snapshot.Data.Series);
        Assert.Equal("Second", series.Name);
        Assert.Equal(new[] { 6d, 7d, 8d }, series.Values);
    }

    [Fact]
    public void Snapshot_LeavesOutRowsWithoutRecordsWhenTheSheetHidesRowsByDefault() {
        using var document = WordDocument.Create();
        var chart = AddChartWithTwoSeries(document);
        EditEmbeddedWorksheet(chart, worksheet => {
            // With zeroHeight a row without a record is hidden; a row record keeps its row visible.
            var format = worksheet.GetFirstChild<S.SheetFormatProperties>();
            if (format == null) {
                format = new S.SheetFormatProperties { DefaultRowHeight = 15D };
                worksheet.InsertBefore(format, worksheet.GetFirstChild<S.SheetData>());
            }
            format.ZeroHeight = true;
            Row(worksheet, 3).Remove();
        });

        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));

        Assert.Equal(new[] { "A", "C" }, snapshot.Data.Categories);
        Assert.Equal(new[] { 3d, 5d }, snapshot.Data.Series[0].Values);
    }

    [Fact]
    public void Snapshot_LeavesCategoryLabelsEmptyWhenTheirColumnIsHidden() {
        using var document = WordDocument.Create();
        var chart = AddChartWithTwoSeries(document);
        // Excel draws the visible values under empty labels, not numbered categories.
        EditEmbeddedWorksheet(chart, worksheet => {
            var columns = worksheet.GetFirstChild<S.Columns>() ?? worksheet.InsertBefore(new S.Columns(), worksheet.GetFirstChild<S.SheetData>());
            columns.Append(new S.Column { Min = 1U, Max = 1U, Hidden = true });
        });

        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));

        Assert.Equal(new[] { "", "", "" }, snapshot.Data.Categories);
        Assert.Equal(new[] { 3d, 4d, 5d }, snapshot.Data.Series[0].Values);
    }

    [Fact]
    public void Snapshot_IgnoresAHiddenSeriesWhenZeroBlanksCompareSeriesLengths() {
        using var document = WordDocument.Create();
        var chart = AddChartWithTwoSeries(document);
        C.Chart native = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!;
        native.GetFirstChild<C.DisplayBlanksAs>()!.Val = C.DisplayBlanksAsValues.Zero;
        EditEmbeddedWorksheet(chart, worksheet => {
            var columns = worksheet.GetFirstChild<S.Columns>() ?? worksheet.InsertBefore(new S.Columns(), worksheet.GetFirstChild<S.SheetData>());
            columns.Append(new S.Column { Min = 2U, Max = 2U, Hidden = true });
        });

        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));

        Assert.Equal(new[] { 6d, 7d, 8d }, Assert.Single(snapshot.Data.Series).Values);
    }

    [Fact]
    public void Snapshot_RejectsZeroBlanksWhenASeriesIsShorterThanItsCategories() {
        using var document = WordDocument.Create();
        var chart = AddChartWithTwoSeries(document);
        C.Chart native = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!;
        native.GetFirstChild<C.DisplayBlanksAs>()!.Val = C.DisplayBlanksAsValues.Zero;
        // The categories read rows that stay visible, so they keep three labels while the remaining series has two.
        foreach (C.Formula categories in native.Descendants<C.CategoryAxisData>().SelectMany(data => data.Descendants<C.Formula>()))
            categories.Text = categories.Text.Replace("$2:", "$6:").Replace("$4", "$8");
        EditEmbeddedWorksheet(chart, worksheet => {
            var columns = worksheet.GetFirstChild<S.Columns>() ?? worksheet.InsertBefore(new S.Columns(), worksheet.GetFirstChild<S.SheetData>());
            columns.Append(new S.Column { Min = 2U, Max = 2U, Hidden = true });
            Row(worksheet, 3).Hidden = true;
        });

        // The zero blank rule would draw the missing third point as zero.
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void SourceVisibility_IgnoresTheCategoriesOfAHiddenSeriesWhenZeroBlanksCompareLengths() {
        using var document = WordDocument.Create();
        var chart = AddChartWithTwoSeries(document);
        C.Chart native = chart.ChartPart!.ChartSpace!.GetFirstChild<C.Chart>()!;
        native.GetFirstChild<C.DisplayBlanksAs>()!.Val = C.DisplayBlanksAsValues.Zero;
        // The first series, whose values are hidden, labels a fourth category the remaining series never reaches.
        C.CategoryAxisData first = native.Descendants<C.BarChartSeries>().First().GetFirstChild<C.CategoryAxisData>()!;
        C.Formula formula = first.Descendants<C.Formula>().Single();
        formula.Text = formula.Text.Replace("$4", "$5");
        C.StringCache cache = first.Descendants<C.StringCache>().Single();
        cache.GetFirstChild<C.PointCount>()!.Val = 4U;
        cache.Append(new C.StringPoint(new C.NumericValue("D")) { Index = 3U });
        EditEmbeddedWorksheet(chart, worksheet => {
            var columns = worksheet.GetFirstChild<S.Columns>() ?? worksheet.InsertBefore(new S.Columns(), worksheet.GetFirstChild<S.SheetData>());
            columns.Append(new S.Column { Min = 2U, Max = 2U, Hidden = true });
        });

        // The hidden-cell projection accepts it: nothing it keeps needs padding. (The flat Word projection still
        // declines series whose category caches differ, which is a separate rule.)
        Assert.True(OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartSeriesReader.TryReadSourceVisibility(
            chart.ChartPart!, native, out _, out string? reason), reason);
    }

    [Fact]
    public void LegacySnapshot_RejectsASeriesShortenedByHiddenCells() {
        using var document = WordDocument.Create();
        var chart = AddChartWithTwoSeries(document);
        // The second series reads rows below the hidden one, so only the first loses a point.
        C.Formula second = chart.ChartPart!.ChartSpace!.Descendants<C.Values>().ElementAt(1).Descendants<C.Formula>().Single();
        second.Text = second.Text.Replace("$2:", "$6:").Replace("$4", "$8");
        EditEmbeddedWorksheet(chart, worksheet => Row(worksheet, 3).Hidden = true);

        Assert.True(chart.TryGetOfficeSnapshot(out var shared));
        Assert.True(double.IsNaN(shared.Data.Series[0].Values[2]));
        // The legacy projection has no missing values, so it must not show the gap as zero.
        Assert.False(chart.TryGetSnapshot(out _));
    }

    [Fact]
    public void Snapshot_RejectsHiddenWorkbookCellsItCannotMatchToCachedPoints() {
        using var document = WordDocument.Create();
        var chart = AddChartWithTwoSeries(document);
        // The cache covers fewer points than the hidden range has cells.
        var cache = chart.ChartPart!.ChartSpace!.Descendants<C.Values>().First().Descendants<C.NumberingCache>().Single();
        cache.GetFirstChild<C.PointCount>()!.Val = 2U;
        cache.Elements<C.NumericPoint>().Last().Remove();
        EditEmbeddedWorksheet(chart, worksheet => Row(worksheet, 3).Hidden = true);

        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    private static WordChart AddChartWithTwoSeries(WordDocument document) =>
        document.AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(new[] { "A", "B", "C" }, new[] {
            new OfficeSeries("Values", new[] { 3d, 4d, 5d }), new OfficeSeries("Second", new[] { 6d, 7d, 8d }) }));

    private static S.Row Row(S.Worksheet worksheet, uint index) =>
        worksheet.Descendants<S.Row>().Single(row => row.RowIndex?.Value == index);

    private static void EditEmbeddedWorksheet(WordChart chart, Action<S.Worksheet> edit) {
        var embedded = chart.ChartPart!.GetPartsOfType<DocumentFormat.OpenXml.Packaging.EmbeddedPackagePart>().Single();
        using var bytes = new MemoryStream();
        using (var source = embedded.GetStream()) source.CopyTo(bytes);
        bytes.Position = 0;
        using (var workbook = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(bytes, true)) {
            S.Worksheet worksheet = workbook.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            edit(worksheet);
            worksheet.Save();
        }
        bytes.Position = 0;
        embedded.FeedData(bytes);
    }

    [Fact]
    public void Snapshot_IgnoresHiddenRowsAndColumnsOutsideReferencedChartRanges() {
        using var document = WordDocument.Create();
        var chart = document.AddChart(OfficeChartType.ColumnClustered,            new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeSeries("Values", new[] { 3d, 4d }) }));
        var part = chart.ChartPart!;
        part.ChartSpace!.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotVisibleOnly>()!.Val = true;
        var embedded = part.GetPartsOfType<DocumentFormat.OpenXml.Packaging.EmbeddedPackagePart>().Single();
        using var bytes = new MemoryStream();
        using (var source = embedded.GetStream()) source.CopyTo(bytes);
        bytes.Position = 0;
        using (var workbook = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(bytes, true)) {
            S.Worksheet worksheet = workbook.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            worksheet.GetFirstChild<S.SheetData>()!.Append(new S.Row { RowIndex = 1000U, Hidden = true });
            var columns = worksheet.GetFirstChild<S.Columns>() ?? new S.Columns();
            columns.Append(new S.Column { Min = 20U, Max = 20U, Hidden = true });
            if (columns.Parent == null) worksheet.AddChild(columns, true);
            worksheet.Save();
        }
        bytes.Position = 0;
        embedded.FeedData(bytes);
        Assert.True(chart.TryGetOfficeSnapshot(out _));
    }

    [Theory]
    [InlineData("dropLines")]
    [InlineData("negative")]
    [InlineData("dateAxis")]
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
    [InlineData("hierarchy")]
    [InlineData("secondaryCategory")]
    [InlineData("labelOffset")]
    [InlineData("rounded")]
    [InlineData("labelAlignment")]
    [InlineData("barShape")]
    [InlineData("barSeriesShape")]
    [InlineData("seriesLines")]
    [InlineData("userShapes")]
    [InlineData("radialInsideEnd")]
    [InlineData("radialNoVary")]
    public void Snapshot_RejectsUnrepresentedNativeChartContent(string feature) {
        using var document = WordDocument.Create();
        var kind = feature.StartsWith("radial", StringComparison.Ordinal) ? OfficeChartType.Pie :
            feature == "negative" || feature.StartsWith("bar", StringComparison.Ordinal) || feature == "seriesLines" ? OfficeChartType.ColumnClustered : OfficeChartType.Line;
        var chart = document.AddChart(kind, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeSeries("Values", new[] { 3d, feature == "negative" ? -4d : 4d }) }), new OfficeIMO.Word.WordChartOptions { Title = "Revenue" });
        if (feature == "secondaryCategory") chart.SetData(kind, new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeSeries("Values", new[] { 3d, -4d }),
            new OfficeSeries("Secondary", new[] { 1d, 2d }, null, null, null, true, renderKind: OfficeChartType.Line, axisGroup: OfficeChartAxisGroup.Secondary) }));
        var space = chart.ChartPart!.ChartSpace!;
        var native = space.GetFirstChild<C.Chart>()!;
        var plot = native.PlotArea!;
        var layer = plot.ChildElements.OfType<OpenXmlCompositeElement>().First(item => item.LocalName.EndsWith("Chart", StringComparison.Ordinal));
        var series = layer.ChildElements.OfType<OpenXmlCompositeElement>().First(item => item.LocalName == "ser");
        if (feature == "dropLines") layer.AddChild(new C.DropLines(), true);
        else if (feature == "negative") series.AddChild(new C.InvertIfNegative { Val = true }, true);
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
            layer.InsertAfter(sibling, series);
        }
        else if (feature == "titleParagraphs") native.GetFirstChild<C.Title>()!.Descendants<C.RichText>().Single().Append(new A.Paragraph(new A.Run(new A.Text("2026"))));
        else if (feature == "titleBreak") native.GetFirstChild<C.Title>()!.GetFirstChild<C.ChartText>()!.GetFirstChild<C.RichText>()!.Descendants<A.Paragraph>().Single().Append(new A.Break(), new A.Run(new A.Text("2026")));
        else if (feature == "dataTable") plot.AddChild(new C.DataTable(new C.ShowHorizontalBorder { Val = true }), true);
        else if (feature == "hierarchy") series.GetFirstChild<C.CategoryAxisData>()!.Append(new C.MultiLevelStringReference());
        else if (feature == "secondaryCategory") plot.Elements<C.CategoryAxis>().Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Top).GetFirstChild<C.Delete>()!.Val = false;
        else if (feature == "labelOffset") plot.GetFirstChild<C.CategoryAxis>()!.GetFirstChild<C.LabelOffset>()!.Val = 200;
        else if (feature == "rounded") space.GetFirstChild<C.RoundedCorners>()!.Val = true;
        else if (feature == "labelAlignment") plot.GetFirstChild<C.CategoryAxis>()!.AddChild(new C.LabelAlignment { Val = C.LabelAlignmentValues.Left }, true);
        else if (feature == "barShape" || feature == "barSeriesShape") {
            var shape = new OpenXmlUnknownElement("c", "shape", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            shape.SetAttribute(new OpenXmlAttribute("val", "", "cylinder"));
            (feature == "barShape" ? layer : series).Append(shape);
        }
        else if (feature == "seriesLines") layer.Append(new C.SeriesLines());
        else if (feature == "userShapes") {
            var drawingPart = chart.ChartPart.AddNewPart<DocumentFormat.OpenXml.Packaging.ChartDrawingPart>();
            var shapes = new OpenXmlUnknownElement("c", "userShapes", "http://schemas.openxmlformats.org/drawingml/2006/chart");
            shapes.SetAttribute(new OpenXmlAttribute("r", "id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", chart.ChartPart.GetIdOfPart(drawingPart)));
            space.Append(shapes);
        }
        else if (feature == "radialOutside" || feature == "radialInsideEnd") layer.AddChild(new C.DataLabels(
            new C.DataLabelPosition { Val = feature == "radialOutside" ? C.DataLabelPositionValues.OutsideEnd : C.DataLabelPositionValues.InsideEnd },
            new C.ShowValue { Val = true }), true);
        else if (feature == "radialNoVary") layer.RemoveAllChildren<C.VaryColors>();
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
