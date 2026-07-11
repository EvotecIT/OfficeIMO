using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ExcelCharts_SharedContractPreservesComboKindsAndSecondaryAxis() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.SharedContract.Combo.xlsx");
            var sharedData = new OfficeChartData(
                new[] { "Q1", "Q2", "Q3" },
                new[] {
                    new OfficeChartSeries(
                        "Adoption",
                        new[] { 42d, 55d, 68d },
                        xValues: null,
                        color: null,
                        pointColors: null,
                        showMarkers: false,
                        showInLegend: true,
                        renderKind: OfficeChartKind.ColumnClustered),
                    new OfficeChartSeries(
                        "Conversion",
                        new[] { 3.2d, 5.1d, 8.4d },
                        xValues: null,
                        color: null,
                        pointColors: null,
                        showMarkers: true,
                        showInLegend: true,
                        renderKind: OfficeChartKind.Line,
                        axisGroup: OfficeChartAxisGroup.Secondary)
                });

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Shared");
                sheet.AddChart(
                    OfficeChartKind.ColumnClustered,
                    sharedData,
                    row: 1,
                    column: 5,
                    widthPixels: 640,
                    heightPixels: 360,
                    title: "Shared performance");
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                OpenXmlValidator validator = new();
                Assert.Empty(validator.Validate(spreadsheet));
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, readOnly: true)) {
                ExcelSheet sheet = document.Sheets.Single(item => item.Name == "Shared");
                ExcelChart chart = Assert.Single(sheet.Charts);
                Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
                Assert.Equal("Shared performance", snapshot.Title);
                Assert.Equal(new[] { "Q1", "Q2", "Q3" }, snapshot.Data.Categories);
                Assert.Collection(
                    snapshot.Data.Series,
                    series => {
                        Assert.Equal("Adoption", series.Name);
                        Assert.Equal(ExcelChartType.ColumnClustered, series.ChartType);
                        Assert.Equal(ExcelChartAxisGroup.Primary, series.AxisGroup);
                    },
                    series => {
                        Assert.Equal("Conversion", series.Name);
                        Assert.Equal(ExcelChartType.Line, series.ChartType);
                        Assert.Equal(ExcelChartAxisGroup.Secondary, series.AxisGroup);
                    });
            }
        }

        [Fact]
        public void Test_ExcelCharts_HorizontalSharedContractPreservesSecondaryAxis() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.SharedContract.HorizontalSecondary.xlsx");
            var sharedData = new OfficeChartData(new[] { "North", "South" }, new[] {
                new OfficeChartSeries("Primary", new[] { 42D, 55D },
                    xValues: null, color: null, pointColors: null, showMarkers: true,
                    renderKind: OfficeChartKind.BarClustered),
                new OfficeChartSeries("Secondary", new[] { 4.2D, 5.5D },
                    xValues: null, color: null, pointColors: null, showMarkers: true,
                    renderKind: OfficeChartKind.BarClustered,
                    axisGroup: OfficeChartAxisGroup.Secondary)
            });

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Shared");
                sheet.AddChart(OfficeChartKind.BarClustered, sharedData,
                    row: 1, column: 5, title: "Horizontal secondary");
                document.Save();
            }

            using ExcelDocument reopened = ExcelDocument.Load(filePath, readOnly: true);
            ExcelSheet reopenedSheet = reopened.Sheets.Single(item => item.Name == "Shared");
            ExcelChart chart = Assert.Single(reopenedSheet.Charts);
            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            Assert.Equal(ExcelChartAxisGroup.Primary, snapshot.Data.Series[0].AxisGroup);
            Assert.Equal(ExcelChartAxisGroup.Secondary, snapshot.Data.Series[1].AxisGroup);
        }

        [Fact]
        public void Test_ExcelCharts_SharedScatterUsesExplicitXValuesWithDisplayLabels() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.SharedContract.ScatterLabels.xlsx");
            var sharedData = new OfficeChartData(new[] { "Discovery", "Delivery" }, new[] {
                new OfficeChartSeries("Actual", new[] { 10D, 14D }, new[] { 1.25D, 2.75D })
            });

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Shared");
                sheet.AddChart(OfficeChartKind.Scatter, sharedData,
                    row: 1, column: 5, title: "Explicit X values");
                document.Save();
            }

            using ExcelDocument reopened = ExcelDocument.Load(filePath, readOnly: true);
            ExcelSheet reopenedSheet = reopened.Sheets.Single(item => item.Name == "Shared");
            ExcelChart chart = Assert.Single(reopenedSheet.Charts);
            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            Assert.Equal(new[] { 1.25D, 2.75D }, Assert.Single(snapshot.Data.Series).XValues);
        }

        [Fact]
        public void Test_ExcelCharts_SharedContractPersistsAuthoredSeriesStyles() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.SharedContract.Styles.xlsx");
            var sharedData = new OfficeChartData(new[] { "Q1", "Q2" }, new[] {
                new OfficeChartSeries("Styled", new[] { 12D, 18D }, xValues: null,
                    color: OfficeColor.ParseHex("#2563EB"),
                    pointColors: new OfficeColor?[] { OfficeColor.ParseHex("#DC2626"), OfficeColor.ParseHex("#16A34A") },
                    showMarkers: true, connectLine: true, markerSize: 9,
                    markerShape: OfficeChartMarkerShape.Diamond,
                    markerOutlineColor: OfficeColor.ParseHex("#111827"), markerOutlineWidth: 1.5D,
                    strokeWidth: 2.25D, strokeDashStyle: OfficeStrokeDashStyle.Dash,
                    showInLegend: false, renderKind: OfficeChartKind.Line)
            });

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Shared");
                sheet.AddChart(OfficeChartKind.Line, sharedData, row: 1, column: 5);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                ChartPart chartPart = spreadsheet.WorkbookPart!.WorksheetParts
                    .Where(worksheet => worksheet.DrawingsPart != null)
                    .SelectMany(worksheet => worksheet.DrawingsPart!.ChartParts)
                    .Single();
                C.LegendEntry entry = Assert.Single(chartPart.ChartSpace.Descendants<C.LegendEntry>());
                Assert.Equal(0U, entry.Index!.Val!.Value);
                Assert.True(entry.GetFirstChild<C.Delete>()!.Val!.Value);
                C.Legend legend = entry.Ancestors<C.Legend>().Single();
                Assert.IsType<C.LegendPosition>(legend.ChildElements[0]);
                Assert.IsType<C.LegendEntry>(legend.ChildElements[1]);
                var validationErrors = new OpenXmlValidator().Validate(spreadsheet).ToList();
                Assert.True(validationErrors.Count == 0, string.Join(Environment.NewLine,
                    validationErrors.Select(error => error.Description + Environment.NewLine +
                        error.Node?.OuterXml)));
            }

            using ExcelDocument reopened = ExcelDocument.Load(filePath, readOnly: true);
            ExcelChart chart = Assert.Single(reopened.Sheets[0].Charts);
            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            ExcelChartSeries series = Assert.Single(snapshot.Data.Series);
            Assert.Equal("2563EB", series.SeriesColorArgb);
            Assert.Equal(2.25D, series.SeriesLineWidth);
            Assert.Equal(OfficeStrokeDashStyle.Dash, series.SeriesLineDashStyle);
            Assert.Equal(new[] { "DC2626", "16A34A" }, series.PointColorArgb);
            Assert.Equal(9, series.MarkerSize);
            Assert.Equal(OfficeChartMarkerShape.Diamond, series.MarkerShape);
            Assert.Equal("111827", series.MarkerOutlineColorArgb);
            Assert.Equal(1.5D, series.MarkerOutlineWidth);
        }

        [Fact]
        public void Test_ExcelCharts_SharedComboStylesMatchNativeSeriesIndexes() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.SharedContract.InterleavedStyles.xlsx");
            var sharedData = new OfficeChartData(new[] { "Q1", "Q2" }, new[] {
                new OfficeChartSeries("Columns A", new[] { 12D, 18D }, null,
                    OfficeColor.ParseHex("#DC2626"), null, showMarkers: false,
                    renderKind: OfficeChartKind.ColumnClustered),
                new OfficeChartSeries("Trend", new[] { 14D, 20D }, null,
                    OfficeColor.ParseHex("#16A34A"), null, showMarkers: true,
                    renderKind: OfficeChartKind.Line),
                new OfficeChartSeries("Columns B", new[] { 10D, 16D }, null,
                    OfficeColor.ParseHex("#2563EB"), null, showMarkers: false,
                    renderKind: OfficeChartKind.ColumnClustered)
            });

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorkSheet("Shared").AddChart(OfficeChartKind.ColumnClustered, sharedData,
                    row: 1, column: 5);
                document.Save();
            }

            using ExcelDocument reopened = ExcelDocument.Load(filePath, readOnly: true);
            ExcelChart chart = Assert.Single(reopened.Sheets[0].Charts);
            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            Assert.Equal("DC2626", snapshot.Data.Series.Single(series => series.Name == "Columns A").SeriesColorArgb);
            Assert.Equal("16A34A", snapshot.Data.Series.Single(series => series.Name == "Trend").SeriesColorArgb);
            Assert.Equal("2563EB", snapshot.Data.Series.Single(series => series.Name == "Columns B").SeriesColorArgb);
        }

        [Fact]
        public void Test_ExcelCharts_UpdateRejectsScatterMixedWithCategoryCharts() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Update.ScatterCombo.xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Shared");
            ExcelChart chart = sheet.AddChart(new ExcelChartData(
                new[] { "Q1", "Q2" },
                new[] {
                    new ExcelChartSeries("Columns", new[] { 1D, 2D }, ExcelChartType.ColumnClustered),
                    new ExcelChartSeries("Trend", new[] { 2D, 3D }, ExcelChartType.Line)
                }),
                row: 1, column: 5, type: ExcelChartType.ColumnClustered);

            NotSupportedException exception = Assert.Throws<NotSupportedException>(() => chart.UpdateData(
                new ExcelChartData(new[] { "Q1", "Q2" }, new[] {
                    new ExcelChartSeries("Scatter", new[] { 2D, 4D }, new[] { 1D, 2D },
                        ExcelChartType.Scatter),
                    new ExcelChartSeries("Line", new[] { 3D, 5D }, ExcelChartType.Line)
                })));

            Assert.Equal("Scatter charts cannot be combined with other chart types.", exception.Message);
        }
    }
}
