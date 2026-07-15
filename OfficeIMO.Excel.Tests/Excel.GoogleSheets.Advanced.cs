using DocumentFormat.OpenXml.Packaging;
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
        public void Test_GoogleSheetsBatch_UsesUniqueGeneratedChartDataSheetName() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsChartDataCollision.xlsx");
            try {
                using (var document = ExcelDocument.Create(path)) {
                    ExcelSheet dashboard = document.AddWorksheet("ChartData");
                    dashboard.CellValue(1, 1, "User content");
                    dashboard.AddChart(
                        new ExcelChartData(new[] { "A", "B" }, new[] { new ExcelChartSeries("Sales", new[] { 2d, 5d }) }),
                        row: 2,
                        column: 4,
                        type: ExcelChartType.ColumnClustered,
                        title: "Sales");
                    document.Save();
                }
                using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                    package.WorkbookPart!.Workbook.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name?.Value == "ChartData").Name = "_OfficeIMO_ChartData";
                    package.WorkbookPart.Workbook.Save();
                }
                using var reloaded = ExcelDocument.Load(path);
                GoogleSheetsBatch batch = new GoogleSheetsExporter().BuildBatch(reloaded);

                GoogleSheetsAddChartRequest chart = Assert.Single(batch.Requests.OfType<GoogleSheetsAddChartRequest>());
                Assert.Equal("_OfficeIMO_ChartData_2", chart.DataSheetName);
                Assert.Contains(batch.Requests.OfType<GoogleSheetsAddSheetRequest>(), request => request.SheetName == "_OfficeIMO_ChartData");
                Assert.Contains(batch.Requests.OfType<GoogleSheetsAddSheetRequest>(), request => request.SheetName == chart.DataSheetName && request.Hidden);
                Assert.Contains(batch.Requests.OfType<GoogleSheetsUpdateCellsRequest>(), request => request.SheetName == "_OfficeIMO_ChartData" && request.Cells.Any(cell => Equals(cell.Value.Value, "User content")));
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Test_GoogleSheetsValuesBatch_EscapesLiteralStringsButKeepsFormulas() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsLiteralValues.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "=1+1");
                sheet.CellValue(1, 2, "00123");
                sheet.CellFormula(1, 3, "1+1");

                GoogleSheetsBatch batch = new GoogleSheetsExporter().BuildBatch(document);
                var ids = GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch);
                GoogleSheetsApiBatchUpdateValuesPayload payload = GoogleSheetsApiPayloadBuilder.BuildValuesBatchUpdatePayload(batch, ids, "spreadsheet-1");
                List<object?> values = Assert.Single(payload.Data).Values.Single();

                Assert.Equal("'=1+1", values[0]);
                Assert.Equal("'00123", values[1]);
                Assert.Equal("=1+1", values[2]);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatch_PreservesSharedStringRichTextRuns() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsSharedStringRichText.xlsx");
            try {
                using (var document = ExcelDocument.Create(path)) {
                    ExcelSheet sheet = document.AddWorksheet("Data");
                    sheet.SetRichText(1, 1, new[] {
                        new ExcelRichTextRun("Bold") { Bold = true },
                        new ExcelRichTextRun(" plain") { Italic = true },
                    });
                    document.Save();
                }

                using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                    WorkbookPart workbookPart = package.WorkbookPart!;
                    WorksheetPart worksheetPart = workbookPart.WorksheetParts.Single();
                    Cell cell = worksheetPart.Worksheet.Descendants<Cell>().Single();
                    InlineString inline = Assert.IsType<InlineString>(cell.InlineString);
                    SharedStringTablePart sharedPart = workbookPart.SharedStringTablePart ?? workbookPart.AddNewPart<SharedStringTablePart>();
                    sharedPart.SharedStringTable ??= new SharedStringTable();
                    int sharedIndex = sharedPart.SharedStringTable.Elements<SharedStringItem>().Count();
                    var item = new SharedStringItem();
                    foreach (var child in inline.ChildElements) item.Append(child.CloneNode(true));
                    sharedPart.SharedStringTable.Append(item);
                    cell.InlineString = null;
                    cell.DataType = CellValues.SharedString;
                    cell.CellValue = new CellValue(sharedIndex.ToString(System.Globalization.CultureInfo.InvariantCulture));
                    sharedPart.SharedStringTable.Save();
                    worksheetPart.Worksheet.Save();
                }

                using var reloaded = ExcelDocument.Load(path);
                GoogleSheetsBatch batch = new GoogleSheetsExporter().BuildBatch(reloaded);
                GoogleSheetsCellData richCell = Assert.Single(batch.Requests.OfType<GoogleSheetsUpdateCellsRequest>().Single(request => request.SheetName == "Data").Cells);

                Assert.Equal(2, richCell.TextFormatRuns.Count);
                Assert.True(richCell.TextFormatRuns[0].Format.Bold);
                Assert.True(richCell.TextFormatRuns[1].Format.Italic);
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
