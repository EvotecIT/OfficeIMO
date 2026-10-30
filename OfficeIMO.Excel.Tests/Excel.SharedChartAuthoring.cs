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
                    new OfficeSeries(
                        "Adoption",
                        new[] { 42d, 55d, 68d },
                        xValues: null,
                        color: null,
                        pointColors: null,
                        showMarkers: false,
                        showInLegend: true,
                        renderKind: OfficeChartType.ColumnClustered),
                    new OfficeSeries(
                        "Conversion",
                        new[] { 3.2d, 5.1d, 8.4d },
                        xValues: null,
                        color: null,
                        pointColors: null,
                        showMarkers: true,
                        showInLegend: true,
                        renderKind: OfficeChartType.Line,
                        axisGroup: OfficeChartAxisGroup.Secondary)
                });

            using (var document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Shared");
                sheet.AddChart(OfficeChartType.ColumnClustered, sharedData, new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 5, Width = OfficeLength.Pixels(640), Height = OfficeLength.Pixels(360), Title = "Shared performance" });
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                ChartPart chartPart = spreadsheet.WorkbookPart!.WorksheetParts
                    .Where(worksheet => worksheet.DrawingsPart != null)
                    .SelectMany(worksheet => worksheet.DrawingsPart!.ChartParts)
                    .Single();
                Assert.Null(chartPart.ChartSpace.Descendants<C.BarChartSeries>().Single()
                    .GetFirstChild<C.Marker>());
                Assert.NotNull(chartPart.ChartSpace.Descendants<C.LineChartSeries>().Single()
                    .GetFirstChild<C.Marker>());
                OpenXmlValidator validator = new();
                var validationErrors = validator.Validate(spreadsheet).ToList();
                Assert.True(validationErrors.Count == 0, string.Join(Environment.NewLine,
                    validationErrors.Select(error => error.Description + Environment.NewLine +
                        error.Node?.OuterXml)));
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                ExcelSheet sheet = document.Sheets.Single(item => item.Name == "Shared");
                ExcelChart chart = Assert.Single(sheet.Charts);
                Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));
                Assert.Equal("Shared performance", snapshot.Title);
                Assert.Equal(new[] { "Q1", "Q2", "Q3" }, snapshot.Data.Categories);
                Assert.Collection(
                    snapshot.Data.Series,
                    series => {
                        Assert.Equal("Adoption", series.Name);
                        Assert.Equal(OfficeChartType.ColumnClustered, series.ChartType);
                        Assert.Equal(OfficeChartAxisGroup.Primary, series.AxisGroup);
                    },
                    series => {
                        Assert.Equal("Conversion", series.Name);
                        Assert.Equal(OfficeChartType.Line, series.ChartType);
                        Assert.Equal(OfficeChartAxisGroup.Secondary, series.AxisGroup);
                    });
            }
        }

        [Fact]
        public void Test_ExcelCharts_HorizontalSharedContractPreservesSecondaryAxis() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.SharedContract.HorizontalSecondary.xlsx");
            var sharedData = new OfficeChartData(new[] { "North", "South" }, new[] {
                new OfficeSeries("Primary", new[] { 42D, 55D },
                    xValues: null, color: null, pointColors: null, showMarkers: true,
                    renderKind: OfficeChartType.BarClustered),
                new OfficeSeries("Secondary", new[] { 4.2D, 5.5D },
                    xValues: null, color: null, pointColors: null, showMarkers: true,
                    renderKind: OfficeChartType.BarClustered,
                    axisGroup: OfficeChartAxisGroup.Secondary)
            });

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Shared");
                sheet.AddChart(OfficeChartType.BarClustered, sharedData, new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 5, Title = "Horizontal secondary" });
                document.Save();
            }

            using ExcelDocument reopened = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            ExcelSheet reopenedSheet = reopened.Sheets.Single(item => item.Name == "Shared");
            ExcelChart chart = Assert.Single(reopenedSheet.Charts);
            Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));
            Assert.Equal(OfficeChartAxisGroup.Primary, snapshot.Data.Series[0].AxisGroup);
            Assert.Equal(OfficeChartAxisGroup.Secondary, snapshot.Data.Series[1].AxisGroup);
        }

        [Fact]
        public void Test_ExcelCharts_SharedScatterUsesExplicitXValuesWithDisplayLabels() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.SharedContract.ScatterLabels.xlsx");
            var sharedData = new OfficeChartData(new[] { "Discovery", "Delivery" }, new[] {
                new OfficeSeries("Actual", new[] { 10D, 14D }, new[] { 1.25D, 2.75D })
            });

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Shared");
                sheet.AddChart(OfficeChartType.Scatter, sharedData, new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 5, Title = "Explicit X values" });
                document.Save();
            }

            using ExcelDocument reopened = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            ExcelSheet reopenedSheet = reopened.Sheets.Single(item => item.Name == "Shared");
            ExcelChart chart = Assert.Single(reopenedSheet.Charts);
            Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));
            Assert.Equal(new[] { 1.25D, 2.75D }, Assert.Single(snapshot.Data.Series).XValues);
        }

        [Fact]
        public void Test_ExcelCharts_SharedBubbleKindIsRejectedInsteadOfFallingBack() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddSheet("Shared");
            var data = new OfficeChartData(new[] { "1", "2" }, new[] {
                OfficeSeries.CreateBubble("Bubbles",
                    new[] { 1D, 2D }, new[] { 10D, 20D }, new[] { 4D, 9D })
            });

            Assert.Throws<NotSupportedException>(() =>
                sheet.AddChart(OfficeChartType.Bubble, data, new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 5 }));
            Assert.Empty(sheet.Charts);
            Assert.Single(document.Sheets);
        }

        [Fact]
        public void Test_ExcelCharts_SharedContractPersistsAuthoredSeriesStyles() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.SharedContract.Styles.xlsx");
            var sharedData = new OfficeChartData(new[] { "Q1", "Q2" }, new[] {
                new OfficeSeries("Styled", new[] { 12D, 18D }, xValues: null,
                    color: OfficeColor.ParseHex("#2563EB"),
                    pointColors: new OfficeColor?[] { OfficeColor.ParseHex("#DC2626"), OfficeColor.ParseHex("#16A34A") },
                    showMarkers: true, connectLine: true, markerSize: 9,
                    markerShape: OfficeChartMarkerShape.Diamond,
                    markerOutlineColor: OfficeColor.ParseHex("#111827"), markerOutlineWidth: 1.5D,
                    strokeWidth: 2.25D, strokeDashStyle: OfficeStrokeDashStyle.Dash,
                    showInLegend: false, renderKind: OfficeChartType.Line)
            });

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddSheet("Shared");
                sheet.AddChart(OfficeChartType.Line, sharedData, new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 5 });
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

            using ExcelDocument reopened = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            ExcelChart chart = Assert.Single(reopened.Sheets[0].Charts);
            Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));
            ExcelChartSeries series = Assert.Single(snapshot.Data.Series);
            Assert.Equal("2563EB", series.SeriesColorArgb);
            Assert.Equal(2.25D, series.SeriesLineWidth?.ToPoints());
            Assert.Equal(OfficeStrokeDashStyle.Dash, series.SeriesLineDashStyle);
            Assert.Equal(new[] { "DC2626", "16A34A" }, series.PointColorArgb);
            Assert.Equal(9, series.MarkerSize?.ToPoints());
            Assert.Equal(OfficeChartMarkerShape.Diamond, series.MarkerShape);
            Assert.Equal("111827", series.MarkerOutlineColorArgb);
            Assert.Equal(1.5D, series.MarkerOutlineWidth?.ToPoints());
        }

        [Fact]
        public void Test_ExcelCharts_SharedContractClampsMarkerSizeToChartSchemaLimit() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.SharedContract.MarkerSize.xlsx");
            var sharedData = new OfficeChartData(new[] { "Q1", "Q2" }, new[] {
                new OfficeSeries("Trend", new[] { 12D, 18D }, xValues: null,
                    color: null, pointColors: null, showMarkers: true, markerSize: 300,
                    renderKind: OfficeChartType.Line)
            });

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddSheet("Shared").AddChart(OfficeChartType.Line, sharedData, new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 5 });
                document.Save();
            }

            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, false);
            C.Marker marker = spreadsheet.WorkbookPart!.WorksheetParts
                .Where(worksheet => worksheet.DrawingsPart != null)
                .SelectMany(worksheet => worksheet.DrawingsPart!.ChartParts)
                .Single().ChartSpace.Descendants<C.Marker>().Single();
            Assert.Equal((byte)72, marker.Size!.Val!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(spreadsheet));
        }

        [Fact]
        public void Test_ExcelCharts_SharedComboStylesMatchNativeSeriesIndexes() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.SharedContract.InterleavedStyles.xlsx");
            var sharedData = new OfficeChartData(new[] { "Q1", "Q2" }, new[] {
                new OfficeSeries("Columns A", new[] { 12D, 18D }, null,
                    OfficeColor.ParseHex("#DC2626"), null, showMarkers: false,
                    renderKind: OfficeChartType.ColumnClustered),
                new OfficeSeries("Trend", new[] { 14D, 20D }, null,
                    OfficeColor.ParseHex("#16A34A"), null, showMarkers: true,
                    renderKind: OfficeChartType.Line),
                new OfficeSeries("Columns B", new[] { 10D, 16D }, null,
                    OfficeColor.ParseHex("#2563EB"), null, showMarkers: false,
                    renderKind: OfficeChartType.ColumnClustered)
            });

            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddSheet("Shared").AddChart(OfficeChartType.ColumnClustered, sharedData, new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 5 });
                document.Save();
            }

            using ExcelDocument reopened = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            ExcelChart chart = Assert.Single(reopened.Sheets[0].Charts);
            Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));
            Assert.Equal("DC2626", snapshot.Data.Series.Single(series => series.Name == "Columns A").SeriesColorArgb);
            Assert.Equal("16A34A", snapshot.Data.Series.Single(series => series.Name == "Trend").SeriesColorArgb);
            Assert.Equal("2563EB", snapshot.Data.Series.Single(series => series.Name == "Columns B").SeriesColorArgb);
        }

        [Fact]
        public void Test_ExcelCharts_UpdateRejectsScatterMixedWithCategoryCharts() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Update.ScatterCombo.xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddSheet("Shared");
            ExcelChart chart = sheet.AddChart(OfficeChartType.ColumnClustered, new ExcelChartData(
                new[] { "Q1", "Q2" },
                new[] {
                    new ExcelChartSeries("Columns", new[] { 1D, 2D }, OfficeChartType.ColumnClustered),
                    new ExcelChartSeries("Trend", new[] { 2D, 3D }, OfficeChartType.Line)
                }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 5 });

            NotSupportedException exception = Assert.Throws<NotSupportedException>(() => chart.UpdateData(
                new ExcelChartData(new[] { "Q1", "Q2" }, new[] {
                    new ExcelChartSeries("Scatter", new[] { 2D, 4D }, new[] { 1D, 2D },
                        OfficeChartType.Scatter),
                    new ExcelChartSeries("Line", new[] { 3D, 5D }, OfficeChartType.Line)
                })));

            Assert.Equal("Scatter charts cannot be combined with other chart types.", exception.Message);
        }
    }
}
