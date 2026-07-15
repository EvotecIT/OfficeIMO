using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_GoogleSheetsBatch_CompilesRichTextConditionalFormattingAndOutlines() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsAdvancedCells.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.SetRichText(1, 1, new[] {
                    new ExcelRichTextRun("Bold") { Bold = true, FontColor = "FF112233" },
                    new ExcelRichTextRun(" plain"),
                });
                sheet.CellValue(2, 1, 12);
                sheet.AddConditionalRule("A2:A5", ConditionalFormattingOperatorValues.GreaterThan, "10", fillColor: "FFC6EFCE");
                sheet.GroupRows(2, 5, outlineLevel: 1);

                GoogleSheetsBatch batch = new GoogleSheetsExporter().BuildBatch(document);

                GoogleSheetsUpdateCellsRequest cells = Assert.Single(batch.Requests.OfType<GoogleSheetsUpdateCellsRequest>(), request => request.SheetName == "Data");
                Assert.Equal(2, Assert.Single(cells.Cells, cell => cell.RowIndex == 0).TextFormatRuns.Count);
                GoogleSheetsAddConditionalFormatRuleRequest conditional = Assert.Single(batch.Requests.OfType<GoogleSheetsAddConditionalFormatRuleRequest>());
                Assert.Equal("NUMBER_GREATER", conditional.ConditionType);
                Assert.Contains(batch.Requests.OfType<GoogleSheetsAddDimensionGroupRequest>(), request => request.StartIndex == 1 && request.EndIndexExclusive == 5);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatch_CompilesSupportedChartThroughHiddenDataRange() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsChart.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                ExcelSheet sheet = document.AddWorksheet("Dashboard");
                sheet.AddChart(
                    new ExcelChartData(new[] { "A", "B" }, new[] { new ExcelChartSeries("Sales", new[] { 2d, 5d }) }),
                    row: 2,
                    column: 4,
                    type: ExcelChartType.ColumnClustered,
                    title: "Sales");

                GoogleSheetsBatch batch = new GoogleSheetsExporter().BuildBatch(document);

                GoogleSheetsAddChartRequest chart = Assert.Single(batch.Requests.OfType<GoogleSheetsAddChartRequest>());
                Assert.Equal("COLUMN", chart.ChartType);
                Assert.Equal("_OfficeIMO_ChartData", chart.DataSheetName);
                Assert.Contains(batch.Requests.OfType<GoogleSheetsAddSheetRequest>(), request => request.SheetName == chart.DataSheetName && request.Hidden);
                Assert.DoesNotContain(batch.Report.Notices, notice => notice.Code == "SHEETS.CHART.UNSUPPORTED");
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatch_CompilesSettingsProtectionAndIdentityMetadata() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsSettings.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                ExcelSheet sheet = document.AddWorksheet("Protected");
                sheet.CellValue(1, 1, "Value");
                sheet.Protect();
                var options = new GoogleSheetsSaveOptions {
                    Spreadsheet = new GoogleSheetsSpreadsheetOptions { Locale = "pl_PL", TimeZone = "Europe/Warsaw", RecalculationInterval = GoogleSheetsRecalculationInterval.Minute },
                    Protection = new GoogleSheetsProtectionOptions { WarningOnly = true, DomainUsersCanEdit = true },
                    Identity = new GoogleSheetsIdentityOptions { WriteDeveloperMetadata = true },
                };
                options.Protection.EditorEmailAddresses.Add("editor@example.com");
                options.Protection.UnprotectedRangesBySheet["Protected"] = new List<string> { "A1:A2" };

                GoogleSheetsBatch batch = new GoogleSheetsExporter().BuildBatch(document, options);

                GoogleSheetsUpdateSpreadsheetPropertiesRequest properties = Assert.Single(batch.Requests.OfType<GoogleSheetsUpdateSpreadsheetPropertiesRequest>());
                Assert.Equal("Europe/Warsaw", properties.TimeZone);
                GoogleSheetsAddProtectedRangeRequest protection = Assert.Single(batch.Requests.OfType<GoogleSheetsAddProtectedRangeRequest>());
                Assert.True(protection.WarningOnly);
                Assert.Equal("editor@example.com", Assert.Single(protection.EditorEmailAddresses));
                Assert.Equal(2, batch.Requests.OfType<GoogleSheetsAddDeveloperMetadataRequest>().Count());
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Test_GoogleSheetsFeatureMatrix_IsCodeOwnedAndExplicit() {
            Assert.Contains(GoogleSheetsFeatureSupportCatalog.Features, feature => feature.Feature == "Charts" && feature.Export == GoogleSheetsFeatureSupportLevel.Partial);
            Assert.Contains(GoogleSheetsFeatureSupportCatalog.Features, feature => feature.Feature == "Embedded drawings and images" && feature.Export == GoogleSheetsFeatureSupportLevel.Unsupported);
        }
    }
}
