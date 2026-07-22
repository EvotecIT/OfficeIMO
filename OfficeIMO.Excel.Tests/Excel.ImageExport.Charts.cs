using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using X = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class ExcelImageExportTests {
        [Fact]
        public void ExcelChart_ImageExportClampsUntrustedLineWidths() {
            var properties = new ChartShapeProperties(
                new A.Outline { Width = int.MaxValue });
            System.Reflection.MethodInfo method = typeof(ExcelChart).GetMethod(
                "TryGetLineWidth",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
            object?[] arguments = { properties, 0D };

            bool resolved = (bool)method.Invoke(null, arguments)!;

            Assert.True(resolved);
            Assert.Equal(64D, Assert.IsType<double>(arguments[1]));
        }

        [Fact]
        public void ExcelChart_ImageExportClampsUntrustedMarkerOutlineWidths() {
            var marker = new Marker(
                new ChartShapeProperties(
                    new A.Outline { Width = int.MaxValue }));
            System.Reflection.MethodInfo method = typeof(ExcelChart).GetMethod(
                "GetImageExportMarkerOutlineWidth",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;

            double? width = (double?)method.Invoke(null, new object?[] { marker });

            Assert.Equal(64D, width);
        }

        [Fact]
        public void ExcelChart_ImageExportPreservesOnePixelTwoCellAnchors() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("TinyAnchor");
            sheet.CellValue(1, 1, "Category");
            sheet.CellValue(1, 2, "Value");
            sheet.CellValue(2, 1, "Only");
            sheet.CellValue(2, 2, 1);
            ExcelChart chart = sheet.AddChartFromRange("A1:B2", row: 1, column: 3);
            ReplaceChartAnchorWithTwoCell(
                document,
                new Xdr.FromMarker(new Xdr.ColumnId("0"), new Xdr.ColumnOffset("0"), new Xdr.RowId("0"), new Xdr.RowOffset("0")),
                new Xdr.ToMarker(new Xdr.ColumnId("0"), new Xdr.ColumnOffset("9525"), new Xdr.RowId("0"), new Xdr.RowOffset("9525")));

            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            Assert.Equal(1, snapshot.WidthPixels);
            Assert.Equal(1, snapshot.HeightPixels);
        }

        [Fact]
        public void ExcelChart_ImageExportRefreshesTwoCellGeometryAfterWorksheetMutation() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("MutableGeometry");
            sheet.CellValue(1, 1, "Category");
            sheet.CellValue(1, 2, "Value");
            sheet.CellValue(2, 1, "Only");
            sheet.CellValue(2, 2, 1);
            ExcelChart chart = sheet.AddChartFromRange("A1:B2", row: 1, column: 3);
            ReplaceChartAnchorWithTwoCell(
                document,
                new Xdr.FromMarker(new Xdr.ColumnId("0"), new Xdr.ColumnOffset("0"), new Xdr.RowId("0"), new Xdr.RowOffset("0")),
                new Xdr.ToMarker(new Xdr.ColumnId("1"), new Xdr.ColumnOffset("0"), new Xdr.RowId("1"), new Xdr.RowOffset("0")));

            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot before));
            sheet.SetColumnWidth(1, 30D);
            sheet.SetRowHeight(1, 45D);
            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot after));

            Assert.True(after.WidthPixels > before.WidthPixels);
            Assert.True(after.HeightPixels > before.HeightPixels);
        }

        [Fact]
        public void ExcelChart_ImageExportUsesCustomRowGeometryAfterInitialIndexBudget() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("LateGeometry");
            sheet.CellValue(1, 1, "Category");
            sheet.CellValue(1, 2, "Value");
            sheet.CellValue(2, 1, "Only");
            sheet.CellValue(2, 2, 1);
            ExcelChart chart = sheet.AddChartFromRange("A1:B2", row: 1, column: 3);
            ReplaceChartAnchorWithTwoCell(
                document,
                new Xdr.FromMarker(new Xdr.ColumnId("0"), new Xdr.ColumnOffset("0"), new Xdr.RowId("100000"), new Xdr.RowOffset("0")),
                new Xdr.ToMarker(new Xdr.ColumnId("1"), new Xdr.ColumnOffset("0"), new Xdr.RowId("100001"), new Xdr.RowOffset("0")));

            X.SheetData sheetData = sheet.WorksheetPart.Worksheet!.GetFirstChild<X.SheetData>()!;
            sheetData.RemoveAllChildren<X.Row>();
            for (uint index = 1; index <= 100000U; index++) {
                sheetData.Append(new X.Row { RowIndex = index });
            }
            sheetData.Append(new X.Row { RowIndex = 100001U, Height = 60D, CustomHeight = true });
            sheetData.Append(new X.Row { RowIndex = 100001U, Height = 30D, CustomHeight = true });

            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            Assert.Equal(80, snapshot.HeightPixels);
        }

        [Fact]
        public void ExcelRange_ImageExportPreservesScatterSeriesCachedXYValues() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ScatterXY");
            sheet.CellValue(1, 1, "X1");
            sheet.CellValue(1, 2, "Y1");
            sheet.CellValue(1, 3, "X2");
            sheet.CellValue(1, 4, "Y2");
            sheet.CellValue(2, 1, 1);
            sheet.CellValue(3, 1, 2);
            sheet.CellValue(2, 2, 10);
            sheet.CellValue(3, 2, 20);
            sheet.CellValue(2, 3, 100);
            sheet.CellValue(3, 3, 200);
            sheet.CellValue(2, 4, 30);
            sheet.CellValue(3, 4, 40);
            sheet.AddScatterChartFromRanges(
                new[] {
                    new ExcelChartSeriesRange("First", "A2:A3", "B2:B3"),
                    new ExcelChartSeriesRange("Second", "C2:C3", "D2:D3")
                },
                row: 1,
                column: 6,
                widthPixels: 260,
                heightPixels: 170,
                title: "Scatter");
            SetScatterChartSeriesIndexes(document, 1U, 3U);

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:J10").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);

            Assert.Equal(new[] { 1D, 2D }, visualChart.Snapshot.Data.Series[0].XValues);
            Assert.Equal(new[] { 10D, 20D }, visualChart.Snapshot.Data.Series[0].Values);
            Assert.Equal(new[] { 100D, 200D }, visualChart.Snapshot.Data.Series[1].XValues);
            Assert.Equal(new[] { 30D, 40D }, visualChart.Snapshot.Data.Series[1].Values);
        }

        [Fact]
        public void ExcelRange_ImageExportPreservesVariableLengthScatterSeries() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ScatterVariable");
            sheet.CellValue(1, 1, "X1");
            sheet.CellValue(1, 2, "Y1");
            sheet.CellValue(1, 3, "X2");
            sheet.CellValue(1, 4, "Y2");
            sheet.CellValue(2, 1, 1);
            sheet.CellValue(3, 1, 2);
            sheet.CellValue(4, 1, 3);
            sheet.CellValue(2, 2, 10);
            sheet.CellValue(3, 2, 20);
            sheet.CellValue(4, 2, 30);
            sheet.CellValue(2, 3, 100);
            sheet.CellValue(3, 3, 200);
            sheet.CellValue(2, 4, 40);
            sheet.CellValue(3, 4, 50);
            sheet.AddScatterChartFromRanges(
                new[] {
                    new ExcelChartSeriesRange("First", "A2:A4", "B2:B4"),
                    new ExcelChartSeriesRange("Second", "C2:C3", "D2:D3")
                },
                row: 1,
                column: 6,
                widthPixels: 260,
                heightPixels: 170,
                title: "Scatter");

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:J10").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);

            Assert.Equal(new[] { 1D, 2D, 3D }, visualChart.Snapshot.Data.Series[0].XValues);
            Assert.Equal(new[] { 10D, 20D, 30D }, visualChart.Snapshot.Data.Series[0].Values);
            Assert.Equal(new[] { 100D, 200D }, visualChart.Snapshot.Data.Series[1].XValues);
            Assert.Equal(new[] { 40D, 50D }, visualChart.Snapshot.Data.Series[1].Values);
        }

        [Fact]
        public void ExcelRange_ImageExportPreservesChartLevelScatterMarkers() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ScatterMarkers");
            sheet.CellValue(1, 1, "X");
            sheet.CellValue(1, 2, "Y");
            sheet.CellValue(2, 1, 1);
            sheet.CellValue(3, 1, 2);
            sheet.CellValue(2, 2, 10);
            sheet.CellValue(3, 2, 20);
            sheet.AddScatterChartFromRanges(
                new[] { new ExcelChartSeriesRange("Points", "A2:A3", "B2:B3") },
                row: 1,
                column: 4,
                widthPixels: 260,
                heightPixels: 170,
                title: "Markers");

            ExcelVisualChart visualChart = Assert.Single(sheet.Range("A1:H10").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false }).Charts);

            Assert.True(visualChart.Snapshot.Data.Series[0].ShowMarkers);
            Assert.True(visualChart.Snapshot.Data.Series[0].ConnectLine);
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsMarkerOnlyScatterStyle() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ScatterMarkerOnly");
            sheet.CellValue(1, 1, "X");
            sheet.CellValue(1, 2, "Y");
            sheet.CellValue(2, 1, 1);
            sheet.CellValue(3, 1, 2);
            sheet.CellValue(2, 2, 10);
            sheet.CellValue(3, 2, 20);
            sheet.AddScatterChartFromRanges(
                new[] { new ExcelChartSeriesRange("Points", "A2:A3", "B2:B3") },
                row: 1,
                column: 4,
                widthPixels: 260,
                heightPixels: 170,
                title: "Marker Only");
            SetFirstScatterStyle(document, ScatterStyleValues.Marker);

            ExcelVisualChart visualChart = Assert.Single(sheet.Range("A1:H10").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false }).Charts);

            Assert.True(visualChart.Snapshot.Data.Series[0].ShowMarkers);
            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.False(visualChart.Snapshot.Layout!.ConnectScatterPoints);
        }

        [Fact]
        public void ExcelRange_ImageExportRendersSupportedComboChartSeriesTypesIndependently() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Combo");
            var data = new ExcelChartData(
                new[] { "Jan", "Feb", "Mar" },
                new[] {
                    new ExcelChartSeries("Sales", new[] { 10D, 20D, 30D }, ExcelChartType.ColumnClustered),
                    new ExcelChartSeries("Trend", new[] { 12D, 18D, 28D }, ExcelChartType.Line)
                });
            ExcelChart chart = sheet.AddChart(data, row: 1, column: 4, widthPixels: 260, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Combo");
            chart.SetSeriesLineColor(1, "DC2626", widthPoints: 2.5D);

            ExcelRange range = sheet.Range("A1:H10");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 3D };
            ExcelVisualChart visualChart = Assert.Single(range.CreateVisualSnapshot(options).Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);

            Assert.DoesNotContain(png.Diagnostics, diagnostic =>
                diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartKindApproximated &&
                diagnostic.Message.Contains("combo chart", StringComparison.OrdinalIgnoreCase));
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(220, 38, 38),
                    tolerance: 36),
                "Expected the supported combo chart to render the line-series color over the column series.");
        }

        [Fact]
        public void ExcelRange_ImageExportMapsComboSeriesTypesByOrderedSeriesWhenIndexesAreNonContiguous() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ComboIdx");
            var data = new ExcelChartData(
                new[] { "Jan", "Feb" },
                new[] {
                    new ExcelChartSeries("Sales", new[] { 10D, 20D }, ExcelChartType.ColumnClustered),
                    new ExcelChartSeries("Trend", new[] { 12D, 22D }, ExcelChartType.Line)
                });
            sheet.AddChart(data, row: 1, column: 4, widthPixels: 260, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Combo Idx");
            SetChartSeriesIndexes(document, 1U, 2U);

            ExcelVisualChart visualChart = Assert.Single(sheet.Range("A1:H10").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false }).Charts);

            Assert.Collection(
                visualChart.Snapshot.Data.Series,
                series => Assert.Equal(ExcelChartType.ColumnClustered, series.ChartType),
                series => Assert.Equal(ExcelChartType.Line, series.ChartType));
        }

        [Fact]
        public void ExcelRange_ImageExportMapsSeriesStylesByOrderedSeriesWhenIndexesAreNonContiguous() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("StyleIdx");
            var data = new ExcelChartData(
                new[] { "Jan", "Feb" },
                new[] {
                    new ExcelChartSeries("Sales", new[] { 10D, 20D }),
                    new ExcelChartSeries("Trend", new[] { 12D, 22D })
                });
            ExcelChart chart = sheet.AddChart(data, row: 1, column: 4, widthPixels: 260, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Style Idx");
            chart.SetSeriesFillColor(0, "22C55E");
            chart.SetSeriesFillColor(1, "F97316");
            SetChartSeriesIndexes(document, 1U, 2U);

            ExcelVisualChart visualChart = Assert.Single(sheet.Range("A1:H10").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false }).Charts);

            Assert.Equal("22C55E", visualChart.Snapshot.Data.Series[0].SeriesColorArgb);
            Assert.Equal("F97316", visualChart.Snapshot.Data.Series[1].SeriesColorArgb);
        }

        [Fact]
        public void ExcelRange_ImageExportReferencedNumberReaderPreservesBlankPoints() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("BlankPoints");
            sheet.CellValue(2, 2, 10);
            sheet.CellValue(4, 2, 30);
            Type utilities = typeof(ExcelDocument).Assembly.GetType("OfficeIMO.Excel.ExcelChartUtils")!;
            System.Reflection.MethodInfo method = utilities.GetMethod("TryReadReferencedNumberValues", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
            object budget = CreateChartDataPointBudget(utilities);
            object?[] args = { sheet, "BlankPoints!$B$2:$B$4", budget, null };

            bool read = (bool)method.Invoke(null, args)!;

            Assert.True(read);
            Assert.Equal(new[] { 10D, 0D, 30D }, (IReadOnlyList<double>)args[3]!);
        }

        [Fact]
        public void ExcelRange_ImageExportRejectsUnboundedChartFormulaRangesBeforeEnumeration() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            Type utilities = typeof(ExcelDocument).Assembly.GetType("OfficeIMO.Excel.ExcelChartUtils")!;
            System.Reflection.MethodInfo method = utilities.GetMethod("TryReadReferencedNumberValues", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
            object budget = CreateChartDataPointBudget(utilities);
            object?[] args = { sheet, "Data!$A$1:$XFD$1048576", budget, null };

            bool read = (bool)method.Invoke(null, args)!;

            Assert.False(read);
            Assert.Null(args[3]);
        }

        [Fact]
        public void ExcelRange_ImageExportRejectsOversizedChartCachePointCountsBeforeAllocation() {
            Type utilities = typeof(ExcelDocument).Assembly.GetType("OfficeIMO.Excel.ExcelChartUtils")!;
            System.Reflection.MethodInfo method = utilities.GetMethod("TryReadNumberPoints", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
            var cache = new DocumentFormat.OpenXml.Drawing.Charts.NumberingCache(
                new DocumentFormat.OpenXml.Drawing.Charts.PointCount { Val = 1_000_001U });
            object budget = CreateChartDataPointBudget(utilities);
            object?[] args = { cache, budget, null };

            bool read = (bool)method.Invoke(null, args)!;

            Assert.False(read);
            Assert.Null(args[2]);
        }

        [Fact]
        public void ExcelRange_ImageExportChargesActualCachePointsWhenPointCountIsUnderstated() {
            Type utilities = typeof(ExcelDocument).Assembly.GetType("OfficeIMO.Excel.ExcelChartUtils")!;
            System.Reflection.MethodInfo method = utilities.GetMethod("TryReadNumberPoints", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
            var cache = new DocumentFormat.OpenXml.Drawing.Charts.NumberingCache(
                new DocumentFormat.OpenXml.Drawing.Charts.PointCount { Val = 1U },
                new NumericPoint(new NumericValue("1")) { Index = 0U },
                new NumericPoint(new NumericValue("2")) { Index = 1U });
            object budget = CreateChartDataPointBudget(utilities);
            object?[] args = { cache, budget, null };

            Assert.True((bool)method.Invoke(null, args)!);

            System.Reflection.FieldInfo remaining = budget.GetType().GetField("_remaining", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!;
            Assert.Equal(999_998L, (long)remaining.GetValue(budget)!);
        }

        [Fact]
        public void ExcelRange_ImageExportSharesOneAggregateBudgetAcrossChartReferences() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("AggregateBudget");
            sheet.CellValue(1, 1, 1);
            sheet.CellValue(2, 1, 2);
            sheet.CellValue(1, 2, 3);
            Type utilities = typeof(ExcelDocument).Assembly.GetType("OfficeIMO.Excel.ExcelChartUtils")!;
            System.Reflection.MethodInfo method = utilities.GetMethod("TryReadReferencedNumberValues", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
            object budget = CreateChartDataPointBudget(utilities);
            object?[] first = { sheet, "AggregateBudget!$A$1:$A$2", budget, null };
            object?[] second = { sheet, "AggregateBudget!$B$1", budget, null };

            Assert.True((bool)method.Invoke(null, first)!);
            Assert.True((bool)method.Invoke(null, second)!);

            System.Reflection.FieldInfo remaining = budget.GetType().GetField("_remaining", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!;
            Assert.Equal(999_997L, (long)remaining.GetValue(budget)!);
        }

        [Fact]
        public void ExcelRange_ImageExportDoesNotChargeUnusedSeriesCaches() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("SourceBudget");
            sheet.CellValue(1, 1, "Category");
            sheet.CellValue(1, 2, "First");
            sheet.CellValue(1, 3, "Second");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(2, 2, 10D);
            sheet.CellValue(3, 2, 20D);
            sheet.CellValue(2, 3, 30D);
            sheet.CellValue(3, 3, 40D);
            sheet.AddChartFromRange(
                "A1:C3",
                row: 1,
                column: 5,
                widthPixels: 260,
                heightPixels: 170,
                type: ExcelChartType.ColumnClustered,
                title: "Source budget");

            ChartPart chartPart = GetFirstChartPart(document);
            Type utilities = typeof(ExcelDocument).Assembly.GetType("OfficeIMO.Excel.ExcelChartUtils")!;
            Type budgetType = utilities.GetNestedType("ChartDataPointBudget", System.Reflection.BindingFlags.NonPublic)!;
            object budget = Activator.CreateInstance(
                budgetType,
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                binder: null,
                args: new object[] { 6L },
                culture: null)!;
            System.Reflection.MethodInfo method = utilities.GetMethod(
                "TryReadChartDataCore",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;

            ExcelChartData? data = (ExcelChartData?)method.Invoke(null, new[] { chartPart, sheet, budget });

            Assert.NotNull(data);
            Assert.Equal(new[] { "Jan", "Feb" }, data!.Categories);
            Assert.Equal(2, data.Series.Count);
            Assert.Equal(new[] { 10D, 20D }, data.Series[0].Values);
            Assert.Equal(new[] { 30D, 40D }, data.Series[1].Values);
        }

        [Fact]
        public void ExcelRange_ImageExportDoesNotChargeMissingChartSourcesBeforeCacheFallback() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            Type utilities = typeof(ExcelDocument).Assembly.GetType("OfficeIMO.Excel.ExcelChartUtils")!;
            System.Reflection.MethodInfo readReference = utilities.GetMethod("TryReadReferencedNumberValues", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
            System.Reflection.MethodInfo readCache = utilities.GetMethod("TryReadNumberPoints", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
            Type budgetType = utilities.GetNestedType("ChartDataPointBudget", System.Reflection.BindingFlags.NonPublic)!;
            object budget = Activator.CreateInstance(
                budgetType,
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                binder: null,
                args: new object[] { 2L },
                culture: null)!;
            object?[] missingReference = { sheet, "Missing!$A$1:$A$1000000", budget, null };
            var cache = new NumberingCache(
                new PointCount { Val = 2U },
                new NumericPoint(new NumericValue("10")) { Index = 0U },
                new NumericPoint(new NumericValue("20")) { Index = 1U });
            object?[] cachedValues = { cache, budget, null };

            Assert.False((bool)readReference.Invoke(null, missingReference)!);
            Assert.True((bool)readCache.Invoke(null, cachedValues)!);
            Assert.Equal(new[] { 10D, 20D }, (IReadOnlyList<double>)cachedValues[2]!);
        }

        [Fact]
        public void ExcelRange_ImageExportRestoresUnusedReadsAfterInvalidChartSourceFallback() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "not-a-number");
            sheet.CellValue(2, 1, 20D);
            sheet.CellValue(3, 1, 30D);
            sheet.CellValue(4, 1, 40D);
            sheet.CellValue(1, 2, 50D);
            sheet.CellValue(2, 2, 60D);
            Type utilities = typeof(ExcelDocument).Assembly.GetType("OfficeIMO.Excel.ExcelChartUtils")!;
            System.Reflection.MethodInfo readReference = utilities.GetMethod("TryReadReferencedNumberValues", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
            System.Reflection.MethodInfo readCache = utilities.GetMethod("TryReadNumberPoints", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
            Type budgetType = utilities.GetNestedType("ChartDataPointBudget", System.Reflection.BindingFlags.NonPublic)!;
            object budget = Activator.CreateInstance(
                budgetType,
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                binder: null,
                args: new object[] { 4L },
                culture: null)!;
            object?[] invalidReference = { sheet, "Data!$A$1:$A$4", budget, null };
            var cache = new NumberingCache(
                new PointCount { Val = 1U },
                new NumericPoint(new NumericValue("10")) { Index = 0U });
            object?[] cachedValues = { cache, budget, null };
            object?[] validReference = { sheet, "Data!$B$1:$B$2", budget, null };

            Assert.False((bool)readReference.Invoke(null, invalidReference)!);
            Assert.True((bool)readCache.Invoke(null, cachedValues)!);
            Assert.True((bool)readReference.Invoke(null, validReference)!);
            Assert.Equal(new[] { 10D }, (IReadOnlyList<double>)cachedValues[2]!);
            Assert.Equal(new[] { 50D, 60D }, (IReadOnlyList<double>)validReference[3]!);
            System.Reflection.FieldInfo remainingSourceReads = budgetType.GetField("_remainingSourceReads", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance)!;
            Assert.Equal(1L, (long)remainingSourceReads.GetValue(budget)!);
        }

        [Fact]
        public void ExcelRange_ImageExportReusesFirstScatterXValuesWithinAggregateBudget() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("ScatterBudget");
            sheet.CellValue(1, 1, 1D);
            sheet.CellValue(2, 1, 2D);
            sheet.CellValue(1, 3, 10D);
            sheet.CellValue(2, 3, 20D);
            sheet.AddScatterChartFromRanges(
                new[] { new ExcelChartSeriesRange("Points", "A1:A2", "C1:C2") },
                row: 1,
                column: 5,
                widthPixels: 260,
                heightPixels: 170,
                title: "Budget");
            ChartPart chartPart = GetFirstChartPart(document);
            Type utilities = typeof(ExcelDocument).Assembly.GetType("OfficeIMO.Excel.ExcelChartUtils")!;
            Type budgetType = utilities.GetNestedType("ChartDataPointBudget", System.Reflection.BindingFlags.NonPublic)!;
            object budget = Activator.CreateInstance(
                budgetType,
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                binder: null,
                args: new object[] { 4L },
                culture: null)!;
            System.Reflection.MethodInfo method = utilities.GetMethod(
                "TryReadChartDataCore",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;

            ExcelChartData? data = (ExcelChartData?)method.Invoke(null, new[] { chartPart, sheet, budget });

            Assert.NotNull(data);
            ExcelChartSeries series = Assert.Single(data!.Series);
            Assert.Equal(new[] { 1D, 2D }, series.XValues);
            Assert.Equal(new[] { 10D, 20D }, series.Values);
        }

        [Fact]
        public void ExcelRange_ImageExportDoesNotRechargeMaterializedScatterYValues() {
            const uint pointsPerSeries = 200_000U;
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("ScatterXYBudget");
            for (int column = 1; column <= 8; column++) {
                sheet.CellValue(1, column, column);
            }
            sheet.AddScatterChartFromRanges(
                new[] {
                    new ExcelChartSeriesRange("One", "A1:A1", "B1:B1"),
                    new ExcelChartSeriesRange("Two", "C1:C1", "D1:D1"),
                    new ExcelChartSeriesRange("Three", "E1:E1", "F1:F1"),
                    new ExcelChartSeriesRange("Four", "G1:G1", "H1:H1")
                },
                row: 1,
                column: 10,
                widthPixels: 260,
                heightPixels: 170,
                title: "XY budget");

            ChartPart chartPart = GetFirstChartPart(document);
            foreach (ScatterChartSeries chartSeries in chartPart.ChartSpace!.Descendants<ScatterChartSeries>()) {
                ReplaceWithLargeNumberCache(chartSeries.GetFirstChild<XValues>()!.GetFirstChild<NumberReference>()!, pointsPerSeries);
                ReplaceWithLargeNumberCache(chartSeries.GetFirstChild<YValues>()!.GetFirstChild<NumberReference>()!, pointsPerSeries);
            }
            chartPart.ChartSpace.Save();

            string[] categories = new string[(int)pointsPerSeries];
            double[] values = new double[(int)pointsPerSeries];
            var input = new ExcelChartData(
                categories,
                Enumerable.Range(1, 4).Select(index => new ExcelChartSeries("Series " + index, values, ExcelChartType.Scatter)));

            ExcelChartData result = ExcelChartUtils.ApplyScatterSeriesXValues(chartPart, input, sheet);

            Assert.All(result.Series, series => {
                Assert.NotNull(series.XValues);
                Assert.Equal((int)pointsPerSeries, series.XValues!.Count);
            });
        }

        [Fact]
        public void ExcelRange_ImageExportDoesNotChargeReferencedSeriesNamesToDataPointBudget() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("CaptionBudget");
            sheet.CellValue(1, 1, 1D);
            sheet.CellValue(2, 1, 2D);
            sheet.CellValue(1, 2, "Referenced caption");
            sheet.CellValue(1, 3, 10D);
            sheet.CellValue(2, 3, 20D);
            sheet.AddScatterChartFromRanges(
                new[] { new ExcelChartSeriesRange("Placeholder", "A1:A2", "C1:C2") },
                row: 1,
                column: 5,
                widthPixels: 260,
                heightPixels: 170,
                title: "Budget");
            ChartPart chartPart = GetFirstChartPart(document);
            ScatterChartSeries chartSeries = chartPart.ChartSpace!.Descendants<ScatterChartSeries>().First();
            SeriesText seriesText = chartSeries.GetFirstChild<SeriesText>()!;
            seriesText.RemoveAllChildren();
            seriesText.Append(new StringReference(new Formula("CaptionBudget!$B$1")));
            chartPart.ChartSpace.Save();

            Type utilities = typeof(ExcelDocument).Assembly.GetType("OfficeIMO.Excel.ExcelChartUtils")!;
            Type budgetType = utilities.GetNestedType("ChartDataPointBudget", System.Reflection.BindingFlags.NonPublic)!;
            object budget = Activator.CreateInstance(
                budgetType,
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic,
                binder: null,
                args: new object[] { 4L },
                culture: null)!;
            System.Reflection.MethodInfo method = utilities.GetMethod(
                "TryReadChartDataCore",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;

            ExcelChartData? data = (ExcelChartData?)method.Invoke(null, new[] { chartPart, sheet, budget });

            Assert.NotNull(data);
            ExcelChartSeries series = Assert.Single(data!.Series);
            Assert.Equal("Referenced caption", series.Name);
            Assert.Equal(new[] { 1D, 2D }, series.XValues);
            Assert.Equal(new[] { 10D, 20D }, series.Values);
        }

        private static object CreateChartDataPointBudget(Type utilities) {
            Type budgetType = utilities.GetNestedType("ChartDataPointBudget", System.Reflection.BindingFlags.NonPublic)!;
            return Activator.CreateInstance(budgetType, nonPublic: true)!;
        }

        private static void ReplaceWithLargeNumberCache(NumberReference reference, uint pointCount) {
            reference.RemoveAllChildren();
            reference.Append(new NumberingCache(
                new PointCount { Val = pointCount },
                new NumericPoint(new NumericValue("1")) { Index = 0U }));
        }

        private static void ReplaceChartAnchorWithTwoCell(
            ExcelDocument document,
            Xdr.FromMarker fromMarker,
            Xdr.ToMarker toMarker) {
            ChartPart chartPart = GetFirstChartPart(document);
            DrawingsPart drawingsPart = chartPart.GetParentParts().OfType<DrawingsPart>().Single();
            Xdr.OneCellAnchor oneCellAnchor = Assert.Single(drawingsPart.WorksheetDrawing!.Elements<Xdr.OneCellAnchor>());
            Xdr.GraphicFrame frame = oneCellAnchor.GetFirstChild<Xdr.GraphicFrame>()!;
            frame.Remove();
            oneCellAnchor.Remove();
            drawingsPart.WorksheetDrawing.Append(new Xdr.TwoCellAnchor(
                fromMarker,
                toMarker,
                frame,
                new Xdr.ClientData()));
        }

        [Fact]
        public void ExcelRange_ImageExportUsesReferencedValuesWhenVerticalSeriesIsNonAdjacent() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("VerticalGap");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Wrong");
            sheet.CellValue(1, 3, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(2, 2, 999);
            sheet.CellValue(3, 2, 999);
            sheet.CellValue(4, 2, 999);
            sheet.CellValue(2, 3, 10);
            sheet.CellValue(3, 3, 20);
            sheet.CellValue(4, 3, 30);
            sheet.AddChartFromRange("A1:B4", row: 1, column: 5, widthPixels: 260, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Vertical Gap");
            SetFirstBarSeriesValueFormula(document, "VerticalGap!$C$2:$C$4");

            ExcelVisualChart visualChart = Assert.Single(sheet.Range("A1:H10").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false }).Charts);

            ExcelChartSeries series = Assert.Single(visualChart.Snapshot.Data.Series);
            Assert.Equal(new[] { 10D, 20D, 30D }, series.Values);
        }

        [Fact]
        public void ExcelRange_ImageExportPreservesScatterCacheOrderInMixedCharts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ComboScatter");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Sales");
            sheet.CellValue(1, 3, "Trend");
            sheet.CellValue(1, 4, "Trend X");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(2, 2, 10);
            sheet.CellValue(3, 2, 20);
            sheet.CellValue(2, 3, 30);
            sheet.CellValue(3, 3, 40);
            sheet.CellValue(2, 4, 100);
            sheet.CellValue(3, 4, 200);
            sheet.AddChartFromRange("A1:C3", row: 1, column: 6, widthPixels: 260, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Combo Scatter");
            ConvertSecondBarSeriesToScatter(document);

            ExcelVisualChart visualChart = Assert.Single(sheet.Range("A1:J10").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false }).Charts);

            Assert.Equal(new[] { 10D, 20D }, visualChart.Snapshot.Data.Series[0].Values);
            Assert.Null(visualChart.Snapshot.Data.Series[0].XValues);
            Assert.Equal(new[] { 30D, 40D }, visualChart.Snapshot.Data.Series[1].Values);
            Assert.Equal(new[] { 100D, 200D }, visualChart.Snapshot.Data.Series[1].XValues);
        }

        [Fact]
        public void ExcelRange_ImageExportReadsNumericCategoryReferences() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("NumericCats");
            sheet.CellValue(1, 1, "Week");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, 1);
            sheet.CellValue(3, 1, 2);
            sheet.CellValue(4, 1, 3);
            sheet.CellValue(2, 2, 10);
            sheet.CellValue(3, 2, 20);
            sheet.CellValue(4, 2, 30);
            sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 260, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Numeric Cats");
            ReplaceFirstChartCategoryWithNumberReference(document, "NumericCats!$A$2:$A$4", 1D, 2D, 3D);

            ExcelVisualChart visualChart = Assert.Single(sheet.Range("A1:H10").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false }).Charts);

            Assert.Equal(new[] { "1", "2", "3" }, visualChart.Snapshot.Data.Categories);
            ExcelChartSeries series = Assert.Single(visualChart.Snapshot.Data.Series);
            Assert.Equal("Actual", series.Name);
            Assert.Equal(new[] { 10D, 20D, 30D }, series.Values);
        }

        [Fact]
        public void ExcelRange_ImageExportReadsRowOrientedChartReferences() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("RowChart");
            sheet.CellValue(1, 2, "Jan");
            sheet.CellValue(1, 3, "Feb");
            sheet.CellValue(1, 4, "Mar");
            sheet.CellValue(2, 1, "Actual");
            sheet.CellValue(2, 2, 10);
            sheet.CellValue(2, 3, 20);
            sheet.CellValue(2, 4, 30);
            sheet.CellValue(5, 1, "Month");
            sheet.CellValue(5, 2, "Placeholder");
            sheet.CellValue(6, 1, "Jan");
            sheet.CellValue(6, 2, 1);
            sheet.CellValue(7, 1, "Feb");
            sheet.CellValue(7, 2, 2);
            sheet.CellValue(8, 1, "Mar");
            sheet.CellValue(8, 2, 3);
            sheet.AddChartFromRange("A5:B8", row: 1, column: 6, widthPixels: 260, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Row Chart");
            PointFirstChartAtHorizontalData(document);

            ExcelVisualChart visualChart = Assert.Single(sheet.Range("A1:J10").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false }).Charts);

            Assert.Equal(new[] { "Jan", "Feb", "Mar" }, visualChart.Snapshot.Data.Categories);
            ExcelChartSeries series = Assert.Single(visualChart.Snapshot.Data.Series);
            Assert.Equal("Actual", series.Name);
            Assert.Equal(new[] { 10D, 20D, 30D }, series.Values);

            ExcelChart chart = Assert.Single(sheet.Charts);
            chart.UpdateData(new ExcelChartData(
                new[] { "Apr", "May", "Jun" },
                new[] { new ExcelChartSeries("Forecast", new[] { 40D, 50D, 60D }) }));

            Assert.True(sheet.TryGetCellText(1, 2, out string? firstCategory));
            Assert.Equal("Apr", firstCategory);
            Assert.True(sheet.TryGetCellText(2, 1, out string? seriesName));
            Assert.Equal("Forecast", seriesName);
            Assert.True(sheet.TryGetCellText(2, 4, out string? lastValue));
            Assert.Equal("60", lastValue);
        }

        private static void ReplaceFirstChartCategoryWithNumberReference(ExcelDocument document, string formula, params double[] values) {
            ChartPart chartPart = GetFirstChartPart(document);
            CategoryAxisData categoryAxisData = chartPart.ChartSpace!.Descendants<BarChartSeries>().First().GetFirstChild<CategoryAxisData>()!;
            categoryAxisData.RemoveAllChildren<StringReference>();
            categoryAxisData.RemoveAllChildren<NumberReference>();
            categoryAxisData.Append(new NumberReference(new Formula(formula), CreateNumberingCache(values)));
            chartPart.ChartSpace.Save();
        }

        private static void PointFirstChartAtHorizontalData(ExcelDocument document) {
            ChartPart chartPart = GetFirstChartPart(document);
            BarChartSeries series = chartPart.ChartSpace!.Descendants<BarChartSeries>().First();

            SeriesText seriesText = series.GetFirstChild<SeriesText>()!;
            seriesText.RemoveAllChildren<StringReference>();
            seriesText.Append(new StringReference(new Formula("RowChart!$A$2")));

            CategoryAxisData categoryAxisData = series.GetFirstChild<CategoryAxisData>()!;
            StringReference categoryReference = categoryAxisData.GetFirstChild<StringReference>()!;
            categoryReference.Formula = new Formula("RowChart!$B$1:$D$1");

            NumberReference valueReference = series.GetFirstChild<Values>()!.GetFirstChild<NumberReference>()!;
            valueReference.Formula = new Formula("RowChart!$B$2:$D$2");
            chartPart.ChartSpace.Save();
        }

        private static void SetScatterChartSeriesIndexes(ExcelDocument document, params uint[] indexes) {
            ChartPart chartPart = GetFirstChartPart(document);
            ScatterChartSeries[] series = chartPart.ChartSpace!.Descendants<ScatterChartSeries>().ToArray();
            for (int i = 0; i < series.Length && i < indexes.Length; i++) {
                DocumentFormat.OpenXml.Drawing.Charts.Index index = series[i].GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Index>() ?? new DocumentFormat.OpenXml.Drawing.Charts.Index();
                index.Val = indexes[i];
                if (index.Parent == null) {
                    series[i].InsertAt(index, 0);
                }

                Order order = series[i].GetFirstChild<Order>() ?? new Order();
                order.Val = indexes[i];
                if (order.Parent == null) {
                    series[i].InsertAfter(order, index);
                }
            }

            chartPart.ChartSpace.Save();
        }

        private static void SetChartSeriesIndexes(ExcelDocument document, params uint[] indexes) {
            ChartPart chartPart = GetFirstChartPart(document);
            OpenXmlCompositeElement[] series = chartPart.ChartSpace!
                .Descendants<OpenXmlCompositeElement>()
                .Where(element => element is BarChartSeries || element is LineChartSeries || element is ScatterChartSeries || element is AreaChartSeries)
                .ToArray();
            for (int i = 0; i < series.Length && i < indexes.Length; i++) {
                DocumentFormat.OpenXml.Drawing.Charts.Index index = series[i].GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Index>() ?? new DocumentFormat.OpenXml.Drawing.Charts.Index();
                index.Val = indexes[i];
                if (index.Parent == null) {
                    series[i].InsertAt(index, 0);
                }

                Order order = series[i].GetFirstChild<Order>() ?? new Order();
                order.Val = indexes[i];
                if (order.Parent == null) {
                    series[i].InsertAfter(order, index);
                }
            }

            chartPart.ChartSpace.Save();
        }

        private static void SetFirstBarSeriesValueFormula(ExcelDocument document, string formula) {
            ChartPart chartPart = GetFirstChartPart(document);
            BarChartSeries series = chartPart.ChartSpace!.Descendants<BarChartSeries>().First();
            NumberReference reference = series.GetFirstChild<Values>()?.GetFirstChild<NumberReference>() ?? new NumberReference();
            reference.Formula = new Formula(formula);
            Values values = series.GetFirstChild<Values>() ?? new Values();
            if (reference.Parent == null) {
                values.Append(reference);
            }

            if (values.Parent == null) {
                series.Append(values);
            }

            chartPart.ChartSpace.Save();
        }

        private static void SetFirstScatterStyle(ExcelDocument document, ScatterStyleValues styleValue) {
            ChartPart chartPart = GetFirstChartPart(document);
            ScatterChart scatterChart = chartPart.ChartSpace!.Descendants<ScatterChart>().First();
            ScatterStyle style = scatterChart.GetFirstChild<ScatterStyle>() ?? new ScatterStyle();
            style.Val = styleValue;
            if (style.Parent == null) {
                scatterChart.InsertAt(style, 0);
            }

            chartPart.ChartSpace.Save();
        }

        private static void SetFirstChartAnchorOffset(ExcelDocument document, int offsetXPixels, int offsetYPixels) {
            DrawingsPart drawingsPart = document.WorkbookPartRoot.WorksheetParts.Select(part => part.DrawingsPart).First(part => part != null)!;
            Xdr.OneCellAnchor anchor = drawingsPart.WorksheetDrawing!.Descendants<Xdr.OneCellAnchor>().First(item => item.GetFirstChild<Xdr.GraphicFrame>() != null);
            anchor.FromMarker!.ColumnOffset = new Xdr.ColumnOffset((offsetXPixels * 9525L).ToString(System.Globalization.CultureInfo.InvariantCulture));
            anchor.FromMarker.RowOffset = new Xdr.RowOffset((offsetYPixels * 9525L).ToString(System.Globalization.CultureInfo.InvariantCulture));
            drawingsPart.WorksheetDrawing.Save();
        }

        private static void ConvertSecondBarSeriesToScatter(ExcelDocument document) {
            ChartPart chartPart = GetFirstChartPart(document);
            PlotArea plotArea = chartPart.ChartSpace!.Descendants<PlotArea>().First();
            BarChart barChart = plotArea.GetFirstChild<BarChart>()!;
            BarChartSeries secondSeries = barChart.Elements<BarChartSeries>().ElementAt(1);
            secondSeries.Remove();

            var scatterChart = new ScatterChart(new ScatterStyle { Val = ScatterStyleValues.LineMarker });
            scatterChart.Append(new ScatterChartSeries(
                new DocumentFormat.OpenXml.Drawing.Charts.Index { Val = 1U },
                new Order { Val = 1U },
                new SeriesText(new NumericValue("Trend")),
                new XValues(new NumberReference(
                    new Formula("ComboScatter!$D$2:$D$3"),
                    CreateNumberingCache(100D, 200D))),
                new YValues(new NumberReference(
                    new Formula("ComboScatter!$C$2:$C$3"),
                    CreateNumberingCache(30D, 40D)))));

            foreach (AxisId axisId in barChart.Elements<AxisId>()) {
                scatterChart.Append((AxisId)axisId.CloneNode(true));
            }

            plotArea.InsertAfter(scatterChart, barChart);
            chartPart.ChartSpace.Save();
        }

        private static NumberingCache CreateNumberingCache(params double[] values) {
            var cache = new NumberingCache(new FormatCode("General"), new PointCount { Val = (uint)values.Length });
            for (int i = 0; i < values.Length; i++) {
                cache.Append(new NumericPoint(new NumericValue(values[i].ToString(System.Globalization.CultureInfo.InvariantCulture))) { Index = (uint)i });
            }

            return cache;
        }

        [Fact]
        public void ExcelRange_ImageExportIncludesChartsThatOverlapSelectedRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartOverlap");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 10);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 20);
            sheet.AddChartFromRange("A1:B3", row: 1, column: 1, widthPixels: 260, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Overlap");

            ExcelRangeVisualSnapshot snapshot = sheet.Range("B1:D8").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);

            Assert.True(visualChart.X < 0D);
            Assert.True(visualChart.X + visualChart.Width > 0D);
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsChartAnchorOffsets() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartOffset");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 10);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 20);
            sheet.AddChartFromRange("A1:B3", row: 1, column: 2, widthPixels: 160, heightPixels: 90, type: ExcelChartType.ColumnClustered, title: "Offset");
            SetFirstChartAnchorOffset(document, offsetXPixels: 23, offsetYPixels: 17);

            ExcelVisualChart visualChart = Assert.Single(sheet.Range("A1:D8").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false }).Charts);

            Assert.Equal(23, visualChart.Snapshot.OffsetXPixels);
            Assert.Equal(17, visualChart.Snapshot.OffsetYPixels);
            Assert.True(visualChart.X > 80D, $"Expected the chart X coordinate to include the from-marker offset. X={visualChart.X}");
            Assert.Equal(17D, visualChart.Y, precision: 0);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartBodyTextColorsIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartBodyText");
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
        public void ExcelRange_ImageExportCarriesChartDataLabelShapeStyleIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartLabelBox");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Label Boxes");
            chart.SetDataLabels(
                showLegendKey: false,
                showValue: true,
                showCategoryName: false,
                showSeriesName: false,
                showPercent: false,
                position: DataLabelPositionValues.OutsideEnd,
                numberFormat: "0");
            chart.SetDataLabelShapeStyle(fillColor: "FDE68A", lineColor: "B45309", lineWidthPoints: 1.5D);
            chart.SetDataLabelTextStyle(color: "7C2D12", bold: true);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 3D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Style);
            Assert.Equal(OfficeColor.FromRgb(253, 230, 138), visualChart.Snapshot.Style!.DataLabelFillColor);
            Assert.Equal(OfficeColor.FromRgb(180, 83, 9), visualChart.Snapshot.Style.DataLabelBorderColor);
            Assert.Equal(1.5D, visualChart.Snapshot.Style.DataLabelBorderWidth!.Value, 3);
            Assert.Equal(OfficeColor.FromRgb(124, 45, 18), visualChart.Snapshot.Style.DataLabelTextColor);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartSeriesStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartTextStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("#FDE68A", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#B45309", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#7C2D12", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(253, 230, 138),
                    tolerance: 28),
                "Expected the exported chart to include the authored data-label fill color.");
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(180, 83, 9),
                    tolerance: 42),
                "Expected the exported chart to include the authored data-label border color.");
        }

        [Fact]
        public void ExcelRange_ImageExportResolvesChartSeriesThemeColorsThroughSharedResolver() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartThemeColor");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Theme Series");
            SetFirstChartSeriesSchemeFill(document, A.SchemeColorValues.Accent2);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 3D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.Equal("C0504D", visualChart.Snapshot.Data.Series[0].SeriesColorArgb);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartSeriesStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("#C0504D", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(192, 80, 77),
                    tolerance: 8),
                "Expected the exported chart to include the workbook theme accent2 series color.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartAxisTextColorIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartAxisText");
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
            ExcelSheet sheet = document.AddWorksheet("ChartTextSizes");
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
        public void ExcelRange_ImageExportReportsUnresolvedChartFontFamilyFallback() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartFont");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Missing Font");
            chart.SetTitleTextStyle(fontName: "OfficeIMO Missing Chart Font", color: "BE123C");

            OfficeImageExportResult png = sheet.Range("A1:H9").ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D });

            OfficeImageExportDiagnostic diagnostic = Assert.Single(
                png.Diagnostics,
                item => item.Code ==
                        OfficeImageExportDiagnosticCodes.FontSubstituted &&
                        item.Source == "ChartFont!" + chart.Name &&
                        item.Message.Contains(
                            "OfficeIMO Missing Chart Font",
                            StringComparison.Ordinal));
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("ChartFont!" + chart.Name, diagnostic.Source);
            Assert.Contains("OfficeIMO Missing Chart Font", diagnostic.Message);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartTextStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        [Fact]
        public void ExcelRange_ImageExportUsesCallerScopedChartFont() {
            OfficeTrueTypeFont? font = OfficeTrueTypeFont.TryLoadDefault(out string? fontPath);
            if (font == null || string.IsNullOrWhiteSpace(fontPath)) return;

            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartFont");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            ExcelChart chart = sheet.AddChartFromRange(
                "A1:B2",
                row: 1,
                column: 4,
                widthPixels: 265,
                heightPixels: 170,
                type: ExcelChartType.ColumnClustered,
                title: "Scoped Font");
            chart.SetTitleTextStyle(
                fontName: "OfficeIMO Scoped Chart Font",
                color: "BE123C");
            var options = new ExcelImageExportOptions {
                ShowGridlines = false
            };
            options.Fonts.Add(
                "OfficeIMO Scoped Chart Font",
                File.ReadAllBytes(fontPath));
            options.Fonts.Add(
                "OfficeIMO Scoped Chart Font",
                File.ReadAllBytes(fontPath),
                OfficeFontStyle.Bold);
            options.Fonts.Add(
                "OfficeIMO Scoped Chart Font",
                File.ReadAllBytes(fontPath),
                OfficeFontStyle.Italic);
            options.Fonts.Add(
                "OfficeIMO Scoped Chart Font",
                File.ReadAllBytes(fontPath),
                OfficeFontStyle.Bold | OfficeFontStyle.Italic);

            OfficeImageExportResult svg = sheet.Range("A1:H9")
                .ExportImage(OfficeImageExportFormat.Svg, options);

            Assert.DoesNotContain(
                svg.Diagnostics,
                item => item.Code ==
                        OfficeImageExportDiagnosticCodes.FontSubstituted &&
                        item.Source == "ChartFont!" + chart.Name &&
                        item.Message.Contains(
                            "OfficeIMO Scoped Chart Font",
                            StringComparison.Ordinal));
            Assert.Contains(
                "OfficeIMO Scoped Chart Font",
                Encoding.UTF8.GetString(svg.Bytes));
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartVerticalValueAxisNumberFormatIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartAxisFormat");
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
            Assert.Contains("900.0", svg, StringComparison.Ordinal);
            Assert.Contains("0.0", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartHorizontalBarValueAxisNumberFormatIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartBarAxisFormat");
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
        public void ExcelRange_ImageExportRendersSecondaryAxisChartSeriesWithoutApproximation() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("SecondaryAxis");
            ExcelChartData data = new(
                new[] { "Q1", "Q2", "Q3" },
                new[] {
                    new ExcelChartSeries("Sales", new[] { 120D, 180D, 160D }, ExcelChartType.ColumnClustered, ExcelChartAxisGroup.Primary),
                    new ExcelChartSeries("Margin", new[] { 0.12D, 0.18D, 0.16D }, ExcelChartType.Line, ExcelChartAxisGroup.Secondary)
                });
            sheet.AddChart(data, row: 1, column: 4, widthPixels: 265, heightPixels: 170,
                type: ExcelChartType.ColumnClustered, title: "Combo");

            ExcelRange range = sheet.Range("A1:H9");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(
                new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D });
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png,
                new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D });

            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            ExcelChartSeries margin = Assert.Single(visualChart.Snapshot.Data.Series,
                series => series.Name == "Margin");
            Assert.Equal(ExcelChartAxisGroup.Secondary, margin.AxisGroup);
            Assert.DoesNotContain(png.Diagnostics,
                item => item.Code == ExcelImageExportDiagnosticCodes.ChartSecondaryAxisUnsupported);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        [Fact]
        public void ExcelRange_ImageExportReportsUnsupportedChartAxisNumberFormat() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartAxisFormatDiag");
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
        public void ExcelRange_ImageExportCarriesSimpleChartCategoryAxisNumberFormatIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartCategoryFormat");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "1");
            sheet.CellValue(2, 2, 1200);
            sheet.CellValue(3, 1, "2");
            sheet.CellValue(3, 2, 1800);
            sheet.CellValue(4, 1, "3");
            sheet.CellValue(4, 2, 1600);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Category Axis Format");
            chart.SetCategoryAxisNumberFormat("0.0");

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.Equal("0.0", visualChart.Snapshot.Layout!.CategoryAxisNumberFormat);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartCategoryAxisNumberFormatUnsupported);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartAxisNumberFormatApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("1.0", svg, StringComparison.Ordinal);
            Assert.Contains("2.0", svg, StringComparison.Ordinal);
            Assert.Contains("3.0", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportReportsUnsupportedChartCategoryAxisDateNumberFormat() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartCategoryFormatDiag");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "1");
            sheet.CellValue(2, 2, 1200);
            sheet.CellValue(3, 1, "2");
            sheet.CellValue(3, 2, 1800);
            sheet.CellValue(4, 1, "3");
            sheet.CellValue(4, 2, 1600);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Category Axis Format Diagnostic");
            chart.SetCategoryAxisNumberFormat("yyyy-mm-dd");

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
            ExcelSheet sheet = document.AddWorksheet("NoCategoryLabels");
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
            ExcelSheet sheet = document.AddWorksheet("NoValueLabels");
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
        public void ExcelRange_ImageExportCarriesHiddenLegendIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("NoLegend");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "No Legend")
                .HideLegend();

            ExcelRange range = sheet.Range("D1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.False(visualChart.Snapshot.Layout!.ShowLegend);
            Assert.DoesNotContain(">Actual<", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsDeletedChartAxes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("DeletedAxes");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Deleted Axes");
            SetFirstChartAxisDeleted<CategoryAxis>(document);
            SetFirstChartAxisDeleted<ValueAxis>(document);

            ExcelRange range = sheet.Range("D1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.False(visualChart.Snapshot.Layout!.ShowCategoryAxis);
            Assert.False(visualChart.Snapshot.Layout.ShowValueAxis);
            Assert.False(visualChart.Snapshot.Layout.ShowCategoryAxisLine);
            Assert.False(visualChart.Snapshot.Layout.ShowValueAxisLine);
            Assert.False(visualChart.Snapshot.Layout.ShowCategoryAxisLabels);
            Assert.False(visualChart.Snapshot.Layout.ShowValueAxisLabels);
            Assert.DoesNotContain("Jan", svg, StringComparison.Ordinal);
            Assert.DoesNotContain("Feb", svg, StringComparison.Ordinal);
            Assert.DoesNotContain("Mar", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesHighChartAxisTickLabelPositionIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartAxisTicks");
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
            ExcelSheet sheet = document.AddWorksheet("ChartTickMarks");
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
                    Math.Abs(shape.Shape.Height) < 0.001D &&
                    shape.Shape.Points.Count == 2 &&
                    Math.Abs(Math.Abs(shape.Shape.Points[1].X - shape.Shape.Points[0].X) - 4D) < 0.001D &&
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
            ExcelSheet sheet = document.AddWorksheet("ChartMinorTicks");
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
                    Math.Abs(shape.Shape.Height) < 0.001D &&
                    shape.Shape.Points.Count == 2 &&
                    Math.Abs(Math.Abs(shape.Shape.Points[1].X - shape.Shape.Points[0].X) - 4D) < 0.001D &&
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
            ExcelSheet sheet = document.AddWorksheet("ChartAxisCrossing");
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
            ExcelSheet sheet = document.AddWorksheet("ChartCategoryCrossing");
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
            ExcelSheet sheet = document.AddWorksheet("ChartAxisScale");
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
            ExcelSheet sheet = document.AddWorksheet("ChartDisplayUnits");
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
            ExcelSheet sheet = document.AddWorksheet("ChartAxisReverse");
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
        public void ExcelRange_ImageExportPreservesHorizontalBarCategoryOrder() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("BarOrder");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 265, heightPixels: 170, type: ExcelChartType.BarClustered, title: "Bar Order");
            chart.SetCategoryAxisReverseOrder(false);

            ExcelRange range = sheet.Range("A1:H9");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D });
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeChartData chartData = new OfficeChartData(
                visualChart.Snapshot.Data.Categories,
                visualChart.Snapshot.Data.Series.Select(series => new OfficeChartSeries(series.Name, series.Values)));
            OfficeChartSnapshot officeChartSnapshot = new OfficeChartSnapshot(
                visualChart.Snapshot.Name,
                visualChart.Snapshot.Title,
                OfficeChartKind.BarClustered,
                chartData,
                visualChart.Width,
                visualChart.Height,
                visualChart.Snapshot.Style,
                visualChart.Snapshot.Layout);
            OfficeDrawing chartDrawing = OfficeChartDrawingRenderer.Render(officeChartSnapshot);

            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.False(visualChart.Snapshot.Layout!.ReverseCategoryAxis);
            OfficeDrawingText janLabel = Assert.Single(chartDrawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "Jan");
            OfficeDrawingText febLabel = Assert.Single(chartDrawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "Feb");
            OfficeDrawingText marLabel = Assert.Single(chartDrawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "Mar");
            Assert.True(janLabel.Y < febLabel.Y && febLabel.Y < marLabel.Y, "Expected horizontal bar categories to follow the authored axis order from top to bottom.");
        }

        [Fact]
        public void ExcelRange_ImageExportKeepsHigherPriorityDataBarsBeforeStopIfTrueFill() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ConditionalPriority");
            sheet.CellValue(1, 1, 10);
            sheet.CellValue(2, 1, 20);
            sheet.CellValue(3, 1, 30);
            sheet.AddConditionalDataBar("A1:A3", OfficeColor.Blue);
            sheet.AddConditionalRule(
                "A1:A3",
                DocumentFormat.OpenXml.Spreadsheet.ConditionalFormattingOperatorValues.GreaterThan,
                "0",
                null,
                "FFFF0000",
                stopIfTrue: true);

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:A3").CreateVisualSnapshot(new ExcelImageExportOptions {
                ShowGridlines = false,
                IncludeConditionalFormatting = true
            });

            Assert.Equal(3, snapshot.ConditionalDataBars.Count);
            Assert.All(snapshot.Cells, cell => Assert.Equal("FFFF0000", cell.Style.FillColorArgb));
        }

        [Fact]
        public void ExcelRange_ImageExportAllowsDistinctChartBodyTextColorBuckets() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartTextConflict");
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

        private static void SetFirstChartAxisDeleted<TAxis>(ExcelDocument document) where TAxis : OpenXmlCompositeElement {
            var chartPart = GetFirstChartPart(document);
            TAxis axis = chartPart.ChartSpace.Descendants<TAxis>().First();
            Delete delete = axis.GetFirstChild<Delete>() ?? new Delete();
            delete.Val = true;
            if (delete.Parent == null) {
                axis.InsertAt(delete, 0);
            }

            chartPart.ChartSpace.Save();
        }

        private static void SetFirstChartSeriesSchemeFill(ExcelDocument document, A.SchemeColorValues schemeColor) {
            var chartPart = GetFirstChartPart(document);
            BarChartSeries series = chartPart.ChartSpace.Descendants<BarChartSeries>().First();
            ChartShapeProperties properties = series.GetFirstChild<ChartShapeProperties>() ?? new ChartShapeProperties();
            properties.RemoveAllChildren<A.SolidFill>();
            properties.PrependChild(new A.SolidFill(new A.SchemeColor { Val = schemeColor }));
            if (properties.Parent == null) {
                series.Append(properties);
            }

            chartPart.ChartSpace.Save();
        }
    }
}
