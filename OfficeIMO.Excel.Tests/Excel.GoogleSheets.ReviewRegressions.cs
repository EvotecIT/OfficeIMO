using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_GoogleSheetsBatch_SizesGridForDimensionsAndProtectionExceptions() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsMetadataGridSizing.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                sheet.SetRowHeight(2000, 24);
                sheet.SetColumnWidth(28, 16);
                sheet.Protect();
                var options = new GoogleSheetsSaveOptions();
                options.Protection.UnprotectedRangesBySheet["Data"] = new List<string> { "A1:A2100" };

                GoogleSheetsBatch batch = new GoogleSheetsExporter().BuildBatch(document, options);

                GoogleSheetsAddSheetRequest addSheet = Assert.Single(
                    batch.Requests.OfType<GoogleSheetsAddSheetRequest>(),
                    request => request.SheetName == "Data");
                Assert.Equal(2100, addSheet.RowCount);
                Assert.Equal(28, addSheet.ColumnCount);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatch_SizesGridsForFiltersConditionalFormatsAndChartAnchors() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsAdvancedMetadataGridSizing.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                ExcelSheet filters = document.AddWorksheet("Filters");
                filters.AddAutoFilter("A1:AA2001");

                ExcelSheet conditional = document.AddWorksheet("Conditional");
                conditional.AddConditionalRule("B2:AB2002", ConditionalFormattingOperatorValues.GreaterThan, "10");

                ExcelSheet charts = document.AddWorksheet("Charts");
                charts.AddChart(
                    new ExcelChartData(new[] { "A" }, new[] { new ExcelChartSeries("Sales", new[] { 1d }) }),
                    row: 2003,
                    column: 29,
                    type: ExcelChartType.ColumnClustered);

                GoogleSheetsBatch batch = new GoogleSheetsExporter().BuildBatch(document);

                GoogleSheetsAddSheetRequest filterSheet = Assert.Single(batch.Requests.OfType<GoogleSheetsAddSheetRequest>(), request => request.SheetName == "Filters");
                Assert.Equal(2001, filterSheet.RowCount);
                Assert.Equal(27, filterSheet.ColumnCount);

                GoogleSheetsAddSheetRequest conditionalSheet = Assert.Single(batch.Requests.OfType<GoogleSheetsAddSheetRequest>(), request => request.SheetName == "Conditional");
                Assert.Equal(2002, conditionalSheet.RowCount);
                Assert.Equal(28, conditionalSheet.ColumnCount);

                GoogleSheetsAddSheetRequest chartSheet = Assert.Single(batch.Requests.OfType<GoogleSheetsAddSheetRequest>(), request => request.SheetName == "Charts");
                Assert.Equal(2003, chartSheet.RowCount);
                Assert.Equal(29, chartSheet.ColumnCount);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatch_AnchorsDuplicateAndUniqueCountRanges() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsCountIfRanges.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.AddConditionalDuplicateValuesRule("A2:A100");
                sheet.AddConditionalUniqueValuesRule("B2:B100");

                GoogleSheetsBatch batch = new GoogleSheetsExporter().BuildBatch(document);
                GoogleSheetsAddConditionalFormatRuleRequest[] rules = batch.Requests
                    .OfType<GoogleSheetsAddConditionalFormatRuleRequest>()
                    .ToArray();

                Assert.Contains(rules, rule => Assert.Single(rule.Values) == "=COUNTIF($A$2:$A$100,A2)>1");
                Assert.Contains(rules, rule => Assert.Single(rule.Values) == "=COUNTIF($B$2:$B$100,B2)=1");
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatch_HonorsWorkbookDateSystemForValidationDates() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheets1904Validation.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                document.DateSystem = ExcelDateSystem.NineteenFour;
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.ValidationDate(
                    "A1:A2",
                    DataValidationOperatorValues.Between,
                    new DateTime(2024, 1, 1),
                    new DateTime(2024, 12, 31));

                GoogleSheetsBatch batch = new GoogleSheetsExporter().BuildBatch(document);

                GoogleSheetsSetDataValidationRequest validation = Assert.Single(
                    batch.Requests.OfType<GoogleSheetsSetDataValidationRequest>());
                Assert.Equal(new[] { "2024-01-01", "2024-12-31" }, validation.Rule.Values);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Test_GoogleSheetsCheckpoint_SeparatesSameNamedSheetScopedRanges() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsScopedNameHashes.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                ExcelSheet north = document.AddWorksheet("North");
                ExcelSheet south = document.AddWorksheet("South");
                north.SetNamedRange("LocalData", "A1:A2", save: false);
                south.SetNamedRange("LocalData", "B1:B2", save: false);

                GoogleSheetsSyncCheckpoint checkpoint = GoogleSheetsDiffPlanner.CreateCheckpoint(document);

                Assert.Contains("name/sheet/North/LocalData", checkpoint.ContentHashes.Keys);
                Assert.Contains("name/sheet/South/LocalData", checkpoint.ContentHashes.Keys);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Test_GoogleSheetsCheckpoint_HashesSupportedCellFormattingAndHyperlinks() {
            string CellHash(Action<ExcelSheet>? configure = null) {
                string path = Path.Combine(_directoryWithFiles, $"GoogleSheetsCellHash-{Guid.NewGuid():N}.xlsx");
                try {
                    using var document = ExcelDocument.Create(path);
                    ExcelSheet sheet = document.AddWorksheet("Data");
                    sheet.CellValue(1, 1, "Value");
                    configure?.Invoke(sheet);
                    return GoogleSheetsDiffPlanner.CreateCheckpoint(document).ContentHashes["sheet/Data/cell/1:1"];
                } finally {
                    if (File.Exists(path)) File.Delete(path);
                }
            }

            string baseline = CellHash();

            Assert.NotEqual(baseline, CellHash(sheet => sheet.CellBold(1, 1)));
            Assert.NotEqual(baseline, CellHash(sheet => sheet.CellItalic(1, 1)));
            Assert.NotEqual(baseline, CellHash(sheet => sheet.CellFontSize(1, 1, 18)));
            Assert.NotEqual(baseline, CellHash(sheet => sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center)));
            Assert.NotEqual(baseline, CellHash(sheet => sheet.CellBorder(1, 1, BorderStyleValues.Medium, "#FF0000")));
            Assert.NotEqual(baseline, CellHash(sheet => sheet.WrapCells(1, 1, 1)));
            Assert.NotEqual(baseline, CellHash(sheet => sheet.SetHyperlink(1, 1, "https://example.test/", display: "Value", style: false)));
        }

        [Fact]
        public void Test_GoogleSheetsCheckpoint_HashesWorksheetTabColor() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsTabColorHash.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Value");
                GoogleSheetsSyncCheckpoint baseline = GoogleSheetsDiffPlanner.CreateCheckpoint(document);

                sheet.SetTabColor("#336699");
                GoogleSheetsSyncCheckpoint colored = GoogleSheetsDiffPlanner.CreateCheckpoint(document);

                Assert.NotEqual(baseline.ContentHashes["sheet/Data"], colored.ContentHashes["sheet/Data"]);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Test_GoogleSheetsCheckpoint_HashesSupportedDimensionsAndValidationRules() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsMetadataHashes.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.SetRowHeight(2, 20);
                sheet.SetColumnWidth(3, 12);
                sheet.ValidationDate(
                    "A1:A2",
                    DataValidationOperatorValues.Between,
                    new DateTime(2024, 1, 1),
                    new DateTime(2024, 12, 31));

                GoogleSheetsSyncCheckpoint baseline = GoogleSheetsDiffPlanner.CreateCheckpoint(document);
                string rowPath = "sheet/Data/row/2";
                string columnPath = "sheet/Data/column/3:3";
                string validationPath = "sheet/Data/validation/0";

                sheet.SetRowHeight(2, 24);
                GoogleSheetsSyncCheckpoint resizedRow = GoogleSheetsDiffPlanner.CreateCheckpoint(document);
                Assert.NotEqual(baseline.ContentHashes[rowPath], resizedRow.ContentHashes[rowPath]);

                sheet.SetRowHidden(2, true);
                GoogleSheetsSyncCheckpoint hiddenRow = GoogleSheetsDiffPlanner.CreateCheckpoint(document);
                Assert.NotEqual(resizedRow.ContentHashes[rowPath], hiddenRow.ContentHashes[rowPath]);

                sheet.SetRowOutline(2, 1);
                GoogleSheetsSyncCheckpoint outlinedRow = GoogleSheetsDiffPlanner.CreateCheckpoint(document);
                Assert.NotEqual(hiddenRow.ContentHashes[rowPath], outlinedRow.ContentHashes[rowPath]);

                sheet.SetColumnWidth(3, 18);
                GoogleSheetsSyncCheckpoint resizedColumn = GoogleSheetsDiffPlanner.CreateCheckpoint(document);
                Assert.NotEqual(baseline.ContentHashes[columnPath], resizedColumn.ContentHashes[columnPath]);

                sheet.SetColumnHidden(3, true);
                GoogleSheetsSyncCheckpoint hiddenColumn = GoogleSheetsDiffPlanner.CreateCheckpoint(document);
                Assert.NotEqual(resizedColumn.ContentHashes[columnPath], hiddenColumn.ContentHashes[columnPath]);

                sheet.SetColumnOutline(3, 1);
                GoogleSheetsSyncCheckpoint outlinedColumn = GoogleSheetsDiffPlanner.CreateCheckpoint(document);
                Assert.NotEqual(hiddenColumn.ContentHashes[columnPath], outlinedColumn.ContentHashes[columnPath]);

                string alternatePath = Path.Combine(_directoryWithFiles, "GoogleSheetsAlternateValidationHash.xlsx");
                try {
                    using var alternate = ExcelDocument.Create(alternatePath);
                    ExcelSheet alternateSheet = alternate.AddWorksheet("Data");
                    alternateSheet.SetRowHeight(2, 20);
                    alternateSheet.SetColumnWidth(3, 12);
                    alternateSheet.ValidationDate(
                        "A1:A2",
                        DataValidationOperatorValues.Between,
                        new DateTime(2024, 1, 1),
                        new DateTime(2025, 12, 31));
                    GoogleSheetsSyncCheckpoint alternateValidation = GoogleSheetsDiffPlanner.CreateCheckpoint(alternate);
                    Assert.NotEqual(baseline.ContentHashes[validationPath], alternateValidation.ContentHashes[validationPath]);
                } finally {
                    if (File.Exists(alternatePath)) File.Delete(alternatePath);
                }
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Test_GoogleSheetsPivot_UsesOnlyHeadersInsideSourceRange() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsPivotSourceHeaders.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 3, "Region");
                sheet.CellValue(1, 4, "Value");
                sheet.CellValue(2, 3, "North");
                sheet.CellValue(2, 4, 10d);
                sheet.CellValue(3, 3, "South");
                sheet.CellValue(3, 4, 20d);
                sheet.AddPivotTable(
                    "C1:D3",
                    "AA2000",
                    name: "ScopedHeaders",
                    rowFields: new[] { "Region" },
                    dataFields: new[] { new ExcelPivotDataField("Value", DataConsolidateFunctionValues.Sum, "Total") });

                GoogleSheetsBatch batch = document.BuildGoogleSheetsBatch();

                GoogleSheetsAddPivotTableRequest pivot = Assert.Single(
                    batch.Requests.OfType<GoogleSheetsAddPivotTableRequest>());
                Assert.Equal(0, Assert.Single(pivot.Rows).SourceColumnOffset);
                Assert.Equal(1, Assert.Single(pivot.Values).SourceColumnOffset);
                GoogleSheetsAddSheetRequest addSheet = Assert.Single(
                    batch.Requests.OfType<GoogleSheetsAddSheetRequest>(),
                    request => request.SheetName == "Data");
                Assert.True(addSheet.RowCount >= 2000);
                Assert.True(addSheet.ColumnCount >= 27);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }
    }
}
