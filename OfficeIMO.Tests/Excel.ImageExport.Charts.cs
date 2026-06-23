using DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class ExcelImageExportTests {
        [Fact]
        public void ExcelRange_ImageExportCarriesChartBodyTextColorsIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("ChartBodyText");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Body Text");
            chart.SetLegendTextStyle(color: "0F766E");
            chart.SetDataLabels(
                showLegendKey: false,
                showValue: true,
                showCategoryName: false,
                showSeriesName: false,
                showPercent: false,
                position: DataLabelPositionValues.OutsideEnd,
                numberFormat: "0");
            chart.SetDataLabelTextStyle(color: "EA580C");

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 3D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Style);
            Assert.Equal(OfficeColor.FromRgb(15, 118, 110), visualChart.Snapshot.Style!.LegendTextColor);
            Assert.Equal(OfficeColor.FromRgb(234, 88, 12), visualChart.Snapshot.Style.DataLabelTextColor);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartTextStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("#0F766E", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#EA580C", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(15, 118, 110),
                    tolerance: 42),
                "Expected the exported chart to include the authored legend text color.");
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(234, 88, 12),
                    tolerance: 42),
                "Expected the exported chart to include the authored data-label text color.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartAxisTextColorIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("ChartAxisText");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Axis Text");
            chart.SetCategoryAxisTitle("Month");
            chart.SetValueAxisTitle("Actual");
            chart.SetCategoryAxisLabelTextStyle(color: "7C3AED");
            chart.SetValueAxisLabelTextStyle(color: "7C3AED");
            chart.SetCategoryAxisTitleTextStyle(color: "DC2626");
            chart.SetValueAxisTitleTextStyle(color: "DC2626");

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 3D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Style);
            Assert.Equal(OfficeColor.FromRgb(124, 58, 237), visualChart.Snapshot.Style!.MutedTextColor);
            Assert.Equal(OfficeColor.FromRgb(220, 38, 38), visualChart.Snapshot.Style.AxisTitleColor);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartTextStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("#7C3AED", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#DC2626", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(124, 58, 237),
                    tolerance: 42),
                "Expected the exported chart to include the authored axis label text color.");
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(220, 38, 38),
                    tolerance: 42),
                "Expected the exported chart to include the authored axis title text color.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartTextFontSizesIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("ChartTextSizes");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Text Sizes");
            chart.SetLegend(LegendPositionValues.Right);
            chart.SetLegendTextStyle(fontSizePoints: 9D, color: "0F766E");
            chart.SetDataLabels(
                showLegendKey: false,
                showValue: true,
                showCategoryName: false,
                showSeriesName: false,
                showPercent: false,
                position: DataLabelPositionValues.OutsideEnd,
                numberFormat: "0");
            chart.SetDataLabelTextStyle(fontSizePoints: 11D, color: "0F766E");
            chart.SetCategoryAxisLabelTextStyle(fontSizePoints: 8D, color: "7C3AED");
            chart.SetValueAxisLabelTextStyle(fontSizePoints: 8D, color: "7C3AED");
            chart.SetCategoryAxisTitle("Month")
                .SetValueAxisTitle("Actual")
                .SetCategoryAxisTitleTextStyle(fontSizePoints: 10D)
                .SetValueAxisTitleTextStyle(fontSizePoints: 10D);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.Equal(9D, visualChart.Snapshot.Layout!.LegendFontSize, 3);
            Assert.Equal(11D, visualChart.Snapshot.Layout.DataLabelFontSize, 3);
            Assert.Equal(8D, visualChart.Snapshot.Layout.AxisLabelFontSize, 3);
            Assert.Equal(10D, visualChart.Snapshot.Layout.AxisTitleFontSize!.Value, 3);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartTextStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("font-size=\"9\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("font-size=\"11\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("font-size=\"8\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("font-size=\"10\"", svg, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartVerticalValueAxisNumberFormatIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("ChartAxisFormat");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 1200);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 1800);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 1600);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Axis Format");
            chart.SetValueAxisNumberFormat("#,##0.0");

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.Equal("#,##0.0", visualChart.Snapshot.Layout!.VerticalAxisNumberFormat);
            Assert.Null(visualChart.Snapshot.Layout.HorizontalAxisNumberFormat);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartAxisNumberFormatApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartCategoryAxisNumberFormatUnsupported);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("1,800.0", svg, StringComparison.Ordinal);
            Assert.Contains("0.0", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartHorizontalBarValueAxisNumberFormatIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("ChartBarAxisFormat");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 1200);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 1800);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 1600);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.BarClustered, title: "Bar Axis Format");
            chart.SetValueAxisNumberFormat("#,##0.0");

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.Equal("#,##0.0", visualChart.Snapshot.Layout!.HorizontalAxisNumberFormat);
            Assert.Null(visualChart.Snapshot.Layout.VerticalAxisNumberFormat);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartAxisNumberFormatApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartCategoryAxisNumberFormatUnsupported);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("1,800.0", svg, StringComparison.Ordinal);
            Assert.Contains("0.0", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportReportsUnsupportedChartAxisNumberFormat() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("ChartAxisFormatDiag");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 1200);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 1800);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 1600);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Axis Format Diagnostic");
            chart.SetValueAxisNumberFormat("yyyy-mm-dd");

            OfficeImageExportResult png = sheet.Range("A1:H9").ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D });

            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartAxisNumberFormatApproximation);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("ChartAxisFormatDiag!" + chart.Name, diagnostic.Source);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartCategoryAxisNumberFormatUnsupported);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        [Fact]
        public void ExcelRange_ImageExportReportsUnsupportedChartCategoryAxisNumberFormat() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("ChartCategoryFormatDiag");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "1");
            sheet.CellValue(2, 2, 1200);
            sheet.CellValue(3, 1, "2");
            sheet.CellValue(3, 2, 1800);
            sheet.CellValue(4, 1, "3");
            sheet.CellValue(4, 2, 1600);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Category Axis Format Diagnostic");
            chart.SetCategoryAxisNumberFormat("0");

            OfficeImageExportResult png = sheet.Range("A1:H9").ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D });

            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartCategoryAxisNumberFormatUnsupported);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("ChartCategoryFormatDiag!" + chart.Name, diagnostic.Source);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartAxisNumberFormatApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesSuppressedCategoryAxisLabelsIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("NoCategoryLabels");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "No Category Labels")
                .SetCategoryAxisTickLabelPosition(TickLabelPositionValues.None);

            ExcelRange range = sheet.Range("D1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.False(visualChart.Snapshot.Layout!.ShowCategoryAxisLabels);
            Assert.DoesNotContain("Jan", svg, StringComparison.Ordinal);
            Assert.DoesNotContain("Feb", svg, StringComparison.Ordinal);
            Assert.DoesNotContain("Mar", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartAxisTickLabelPositionApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesSuppressedValueAxisLabelsIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("NoValueLabels");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "No Value Labels")
                .SetValueAxisNumberFormat("0.0")
                .SetValueAxisTickLabelPosition(TickLabelPositionValues.None);

            ExcelRange range = sheet.Range("D1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.False(visualChart.Snapshot.Layout!.ShowValueAxisLabels);
            Assert.Equal("0.0", visualChart.Snapshot.Layout.VerticalAxisNumberFormat);
            Assert.DoesNotContain("180.0", svg, StringComparison.Ordinal);
            Assert.DoesNotContain("0.0", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartAxisTickLabelPositionApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesHighChartAxisTickLabelPositionIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("ChartAxisTicks");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Axis Ticks");
            chart.SetValueAxisTickLabelPosition(TickLabelPositionValues.High);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeChartData chartData = new OfficeChartData(
                visualChart.Snapshot.Data.Categories,
                visualChart.Snapshot.Data.Series.Select(series => new OfficeChartSeries(series.Name, series.Values)));
            OfficeChartSnapshot officeChartSnapshot = new OfficeChartSnapshot(
                visualChart.Snapshot.Name,
                visualChart.Snapshot.Title,
                OfficeChartKind.ColumnClustered,
                chartData,
                visualChart.Width,
                visualChart.Height,
                visualChart.Snapshot.Style,
                visualChart.Snapshot.Layout);
            OfficeDrawing chartDrawing = OfficeChartDrawingRenderer.Render(officeChartSnapshot);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.Equal(OfficeChartAxisTickLabelPosition.High, visualChart.Snapshot.Layout!.VerticalAxisTickLabelPosition);
            Assert.Contains(
                chartDrawing.Elements.OfType<OfficeDrawingText>(),
                text => text.Text == "180" && text.X > chartDrawing.Width / 2D && text.Alignment == OfficeTextAlignment.Left);
            Assert.Contains(">180<", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartAxisTickLabelPositionApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartAxisMajorTickMarksIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("ChartTickMarks");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Tick Marks");
            SetFirstChartValueAxisMajorTickMark(document, TickMarkValues.Outside);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeChartData chartData = new OfficeChartData(
                visualChart.Snapshot.Data.Categories,
                visualChart.Snapshot.Data.Series.Select(series => new OfficeChartSeries(series.Name, series.Values)));
            OfficeChartSnapshot officeChartSnapshot = new OfficeChartSnapshot(
                visualChart.Snapshot.Name,
                visualChart.Snapshot.Title,
                OfficeChartKind.ColumnClustered,
                chartData,
                visualChart.Width,
                visualChart.Height,
                visualChart.Snapshot.Style,
                visualChart.Snapshot.Layout);
            OfficeDrawing chartDrawing = OfficeChartDrawingRenderer.Render(officeChartSnapshot);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.Equal(OfficeChartAxisTickMark.Outside, visualChart.Snapshot.Layout!.VerticalAxisMajorTickMark);
            Assert.True(
                chartDrawing.Shapes.Count(shape =>
                    shape.Shape.Kind == OfficeShapeKind.Line &&
                    Math.Abs(shape.Shape.Width - 4D) < 0.001D &&
                    Math.Abs(shape.Shape.Height - 1D) < 0.001D &&
                    shape.Shape.Points.Count == 2 &&
                    Math.Abs(shape.Shape.Points[0].Y - shape.Shape.Points[1].Y) < 0.001D) >= 5,
                "Expected the shared chart renderer to draw vertical value-axis major tick marks.");
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == "ExcelChartAxisTickMarkUnsupported");
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("<line", svg, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartAxisMinorTickMarksWithPlacementApproximation() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("ChartMinorTicks");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Minor Ticks");
            SetFirstChartValueAxisMinorTickMark(document, TickMarkValues.Outside);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeChartData chartData = new OfficeChartData(
                visualChart.Snapshot.Data.Categories,
                visualChart.Snapshot.Data.Series.Select(series => new OfficeChartSeries(series.Name, series.Values)));
            OfficeChartSnapshot officeChartSnapshot = new OfficeChartSnapshot(
                visualChart.Snapshot.Name,
                visualChart.Snapshot.Title,
                OfficeChartKind.ColumnClustered,
                chartData,
                visualChart.Width,
                visualChart.Height,
                visualChart.Snapshot.Style,
                visualChart.Snapshot.Layout);
            OfficeDrawing chartDrawing = OfficeChartDrawingRenderer.Render(officeChartSnapshot);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);

            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.Equal(OfficeChartAxisTickMark.Outside, visualChart.Snapshot.Layout!.VerticalAxisMinorTickMark);
            Assert.True(
                chartDrawing.Shapes.Count(shape =>
                    shape.Shape.Kind == OfficeShapeKind.Line &&
                    Math.Abs(shape.Shape.Width - 4D) < 0.001D &&
                    Math.Abs(shape.Shape.Height - 1D) < 0.001D &&
                    shape.Shape.Points.Count == 2 &&
                    Math.Abs(shape.Shape.Points[0].Y - shape.Shape.Points[1].Y) < 0.001D) >= 4,
                "Expected the shared chart renderer to draw vertical value-axis minor tick marks.");
            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartAxisMinorTickMarkPlacementApproximation);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("ChartMinorTicks!" + chart.Name, diagnostic.Source);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == "ExcelChartAxisTickMarkUnsupported");
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesMaximumValueAxisCrossingIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("ChartAxisCrossing");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Axis Crossing");
            chart.SetValueAxisCrossing(CrossesValues.Maximum);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeChartData chartData = new OfficeChartData(
                visualChart.Snapshot.Data.Categories,
                visualChart.Snapshot.Data.Series.Select(series => new OfficeChartSeries(series.Name, series.Values)));
            OfficeChartSnapshot officeChartSnapshot = new OfficeChartSnapshot(
                visualChart.Snapshot.Name,
                visualChart.Snapshot.Title,
                OfficeChartKind.ColumnClustered,
                chartData,
                visualChart.Width,
                visualChart.Height,
                visualChart.Snapshot.Style,
                visualChart.Snapshot.Layout);
            OfficeDrawing chartDrawing = OfficeChartDrawingRenderer.Render(officeChartSnapshot);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.Equal(OfficeChartAxisCrossingPosition.Maximum, visualChart.Snapshot.Layout!.VerticalAxisCrossingPosition);
            Assert.Contains(
                chartDrawing.Elements.OfType<OfficeDrawingText>(),
                text => text.Text == "180" && text.X > chartDrawing.Width / 2D && text.Alignment == OfficeTextAlignment.Left);
            Assert.Contains(">180<", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartAxisCrossingApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesMaximumCategoryAxisCrossingIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("ChartCategoryCrossing");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Category Crossing");
            chart.SetCategoryAxisCrossing(CrossesValues.Maximum);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeChartData chartData = new OfficeChartData(
                visualChart.Snapshot.Data.Categories,
                visualChart.Snapshot.Data.Series.Select(series => new OfficeChartSeries(series.Name, series.Values)));
            OfficeChartSnapshot officeChartSnapshot = new OfficeChartSnapshot(
                visualChart.Snapshot.Name,
                visualChart.Snapshot.Title,
                OfficeChartKind.ColumnClustered,
                chartData,
                visualChart.Width,
                visualChart.Height,
                visualChart.Snapshot.Style,
                visualChart.Snapshot.Layout);
            OfficeDrawing chartDrawing = OfficeChartDrawingRenderer.Render(officeChartSnapshot);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.Equal(OfficeChartAxisCrossingPosition.Maximum, visualChart.Snapshot.Layout!.HorizontalAxisCrossingPosition);
            Assert.Contains(
                chartDrawing.Elements.OfType<OfficeDrawingText>(),
                text => text.Text == "Jan" && text.Y < chartDrawing.Height / 2D);
            Assert.Contains(">Jan<", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartAxisCrossingApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartAxisScaleIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("ChartAxisScale");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Axis Scale");
            chart.SetValueAxisScale(minimum: 100D, maximum: 220D, majorUnit: 40D, minorUnit: 20D);
            chart.SetValueAxisGridlines(showMajor: false, showMinor: true, lineColor: "14B8A6", lineWidthPoints: 1.5D);
            SetFirstChartValueAxisMinorTickMark(document, TickMarkValues.Outside);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeChartData chartData = new OfficeChartData(
                visualChart.Snapshot.Data.Categories,
                visualChart.Snapshot.Data.Series.Select(series => new OfficeChartSeries(series.Name, series.Values)));
            OfficeChartSnapshot officeChartSnapshot = new OfficeChartSnapshot(
                visualChart.Snapshot.Name,
                visualChart.Snapshot.Title,
                OfficeChartKind.ColumnClustered,
                chartData,
                visualChart.Width,
                visualChart.Height,
                visualChart.Snapshot.Style,
                visualChart.Snapshot.Layout);
            OfficeDrawing chartDrawing = OfficeChartDrawingRenderer.Render(officeChartSnapshot);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.Equal(100D, visualChart.Snapshot.Layout!.VerticalAxisMinimum);
            Assert.Equal(220D, visualChart.Snapshot.Layout.VerticalAxisMaximum);
            Assert.Equal(40D, visualChart.Snapshot.Layout.VerticalAxisMajorUnit);
            Assert.Equal(20D, visualChart.Snapshot.Layout.VerticalAxisMinorUnit);
            Assert.Equal(OfficeChartAxisTickMark.Outside, visualChart.Snapshot.Layout.VerticalAxisMinorTickMark);
            Assert.Equal(true, visualChart.Snapshot.Style.ShowValueMinorGridLines);
            Assert.Equal(OfficeColor.FromRgb(20, 184, 166), visualChart.Snapshot.Style.ValueMinorGridLineColor);
            Assert.True(
                chartDrawing.Shapes.Count(shape =>
                    shape.Shape.Kind == OfficeShapeKind.Line &&
                    shape.Shape.StrokeColor == OfficeColor.FromRgb(20, 184, 166)) == 3,
                "Expected the shared chart renderer to draw value-axis minor gridlines at 120, 160, and 200.");
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartAxisScaleApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartAxisMinorTickMarkPlacementApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains(">100<", svg, StringComparison.Ordinal);
            Assert.Contains(">140<", svg, StringComparison.Ordinal);
            Assert.Contains(">180<", svg, StringComparison.Ordinal);
            Assert.Contains(">220<", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartAxisDisplayUnitsIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("ChartDisplayUnits");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120000);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180000);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160000);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Display Units");
            chart.SetValueAxisDisplayUnits(BuiltInUnitValues.Thousands, "Thousands", showLabel: true);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.Equal(1000D, visualChart.Snapshot.Layout!.VerticalAxisDisplayUnitDivisor);
            Assert.Equal("Thousands", visualChart.Snapshot.Layout.VerticalAxisDisplayUnitLabel);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == "ExcelChartAxisDisplayUnitsUnsupported");
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains(">180<", svg, StringComparison.Ordinal);
            Assert.Contains("Thousands", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesCategoryAxisReverseOrderIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("ChartAxisReverse");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Axis Reverse");
            chart.SetCategoryAxisReverseOrder();

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeChartData chartData = new OfficeChartData(
                visualChart.Snapshot.Data.Categories,
                visualChart.Snapshot.Data.Series.Select(series => new OfficeChartSeries(series.Name, series.Values)));
            OfficeChartSnapshot officeChartSnapshot = new OfficeChartSnapshot(
                visualChart.Snapshot.Name,
                visualChart.Snapshot.Title,
                OfficeChartKind.ColumnClustered,
                chartData,
                visualChart.Width,
                visualChart.Height,
                visualChart.Snapshot.Style,
                visualChart.Snapshot.Layout);
            OfficeDrawing chartDrawing = OfficeChartDrawingRenderer.Render(officeChartSnapshot);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.True(visualChart.Snapshot.Layout!.ReverseCategoryAxis);
            OfficeDrawingText janLabel = Assert.Single(chartDrawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "Jan");
            OfficeDrawingText marLabel = Assert.Single(chartDrawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "Mar");
            Assert.True(janLabel.X > marLabel.X, "Expected the shared renderer to place the first source category on the right when the category axis is reversed.");
            Assert.Contains(">Jan<", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartAxisScaleApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        [Fact]
        public void ExcelRange_ImageExportAllowsDistinctChartBodyTextColorBuckets() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("ChartTextConflict");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Text Conflict");
            chart.SetLegendTextStyle(color: "0F766E");
            chart.SetDataLabels(
                showLegendKey: false,
                showValue: true,
                showCategoryName: false,
                showSeriesName: false,
                showPercent: false,
                position: DataLabelPositionValues.OutsideEnd,
                numberFormat: "0");
            chart.SetDataLabelTextStyle(color: "EA580C");

            OfficeImageExportResult png = sheet.Range("A1:H9").ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D });

            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartTextStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        private static void SetFirstChartValueAxisMajorTickMark(ExcelDocument document, TickMarkValues value) {
            var chartPart = GetFirstChartPart(document);
            ValueAxis valueAxis = chartPart.ChartSpace.Descendants<ValueAxis>().First();
            MajorTickMark majorTickMark = valueAxis.GetFirstChild<MajorTickMark>() ?? new MajorTickMark();
            majorTickMark.Val = value;
            if (majorTickMark.Parent == null) {
                valueAxis.Append(majorTickMark);
            }

            chartPart.ChartSpace.Save();
        }

        private static void SetFirstChartValueAxisMinorTickMark(ExcelDocument document, TickMarkValues value) {
            var chartPart = GetFirstChartPart(document);
            ValueAxis valueAxis = chartPart.ChartSpace.Descendants<ValueAxis>().First();
            MinorTickMark minorTickMark = valueAxis.GetFirstChild<MinorTickMark>() ?? new MinorTickMark();
            minorTickMark.Val = value;
            if (minorTickMark.Parent == null) {
                valueAxis.Append(minorTickMark);
            }

            chartPart.ChartSpace.Save();
        }
    }
}
