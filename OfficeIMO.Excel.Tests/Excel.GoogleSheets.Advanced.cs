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
        public void Test_GoogleSheetsBatch_NormalizesExpressionRulesAndDropsRejectedStyles() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsConditionalExpression.xlsx");
            try {
                using (var document = ExcelDocument.Create(path)) {
                    ExcelSheet sheet = document.AddWorksheet("Data");
                    sheet.CellValue(2, 1, 12);
                    sheet.AddConditionalRule("A2:A5", ConditionalFormattingOperatorValues.GreaterThan, "10", fillColor: "FFC6EFCE");
                    document.Save();
                }

                using (SpreadsheetDocument package = SpreadsheetDocument.Open(path, true)) {
                    WorksheetPart worksheetPart = package.WorkbookPart!.WorksheetParts.Single();
                    ConditionalFormattingRule rule = worksheetPart.Worksheet.Descendants<ConditionalFormattingRule>().Single();
                    rule.Type = ConditionalFormatValues.Expression;
                    rule.Operator = null;
                    rule.Elements<Formula>().Single().Text = "A2>10";

                    Stylesheet stylesheet = package.WorkbookPart.WorkbookStylesPart!.Stylesheet!;
                    DifferentialFormat differential = stylesheet.DifferentialFormats!.Elements<DifferentialFormat>()
                        .ElementAt((int)(rule.FormatId?.Value ?? 0));
                    differential.Font = new Font(
                        new FontName { Val = "Aptos" },
                        new Bold(),
                        new Italic(),
                        new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = "FF112233" },
                        new FontSize { Val = 14D },
                        new Underline());
                    differential.Border = new Border(new LeftBorder { Style = BorderStyleValues.Thin });
                    stylesheet.Save();
                    worksheetPart.Worksheet.Save();
                }

                using var reloaded = ExcelDocument.Load(path);
                GoogleSheetsBatch batch = new GoogleSheetsExporter().BuildBatch(reloaded);
                GoogleSheetsAddConditionalFormatRuleRequest conditional = Assert.Single(batch.Requests.OfType<GoogleSheetsAddConditionalFormatRuleRequest>());

                Assert.Equal("CUSTOM_FORMULA", conditional.ConditionType);
                Assert.Equal("=A2>10", Assert.Single(conditional.Values));
                Assert.True(conditional.Format!.Bold);
                Assert.True(conditional.Format.Italic);
                Assert.False(conditional.Format.Underline);
                Assert.Null(conditional.Format.FontName);
                Assert.Null(conditional.Format.FontSize);
                Assert.Null(conditional.Format.Borders);
                Assert.Contains(batch.Report.Notices, notice => notice.Code == "SHEETS.CONDITIONAL_FORMAT.STYLE_REDUCED");

                GoogleSheetsApiBatchUpdatePayload payload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(
                    batch,
                    GoogleSheetsApiPayloadBuilder.BuildSheetIdMap(batch));
                GoogleSheetsApiCellFormatPayload format = Assert.Single(payload.Requests, request => request.AddConditionalFormatRule != null)
                    .AddConditionalFormatRule!.Rule.BooleanRule.Format;
                Assert.Null(format.Borders);
                Assert.Null(format.TextFormat!.Underline);
                Assert.Null(format.TextFormat.FontFamily);
                Assert.Null(format.TextFormat.FontSize);
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
        public void Test_GoogleSheetsBatch_SizesHiddenChartDataGridFromStagedCells() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsLargeChart.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                ExcelSheet sheet = document.AddWorksheet("Dashboard");
                string[] categories = Enumerable.Range(1, 1001).Select(index => $"Category {index}").ToArray();
                ExcelChartSeries[] series = Enumerable.Range(1, 27)
                    .Select(seriesIndex => new ExcelChartSeries(
                        $"Series {seriesIndex}",
                        Enumerable.Range(1, categories.Length).Select(value => (double)(value + seriesIndex)).ToArray()))
                    .ToArray();
                sheet.AddChart(
                    new ExcelChartData(categories, series),
                    row: 2,
                    column: 4,
                    type: ExcelChartType.ColumnClustered,
                    title: "Large chart");

                GoogleSheetsBatch batch = new GoogleSheetsExporter().BuildBatch(document);
                GoogleSheetsAddChartRequest chart = Assert.Single(batch.Requests.OfType<GoogleSheetsAddChartRequest>());
                GoogleSheetsAddSheetRequest dataSheet = Assert.Single(
                    batch.Requests.OfType<GoogleSheetsAddSheetRequest>(),
                    request => request.SheetName == chart.DataSheetName);

                Assert.Equal(1002, dataSheet.RowCount);
                Assert.Equal(28, dataSheet.ColumnCount);

                GoogleSheetsApiCreateSpreadsheetPayload payload = GoogleSheetsApiPayloadBuilder.BuildCreateSpreadsheetPayload(batch);
                GoogleSheetsApiGridPropertiesPayload grid = Assert.Single(
                    payload.Sheets,
                    candidate => candidate.Properties.Title == chart.DataSheetName)
                    .Properties.GridProperties;
                Assert.Equal(1002, grid.RowCount);
                Assert.Equal(28, grid.ColumnCount);
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatch_UsesExplicitScatterXValues() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsScatter.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                ExcelSheet sheet = document.AddWorksheet("Dashboard");
                sheet.AddChart(
                    new ExcelChartData(
                        Array.Empty<string>(),
                        new[] { new ExcelChartSeries("Points", new[] { 10D, 20D }, new[] { 1D, 4D }, ExcelChartType.Scatter) }),
                    row: 2,
                    column: 4,
                    type: ExcelChartType.Scatter,
                    title: "Scatter");

                GoogleSheetsBatch batch = new GoogleSheetsExporter().BuildBatch(document);

                GoogleSheetsAddChartRequest chart = Assert.Single(batch.Requests.OfType<GoogleSheetsAddChartRequest>());
                Assert.Equal("SCATTER", chart.ChartType);
                Assert.Equal(3, chart.DataRowCount);
                GoogleSheetsUpdateCellsRequest chartData = Assert.Single(
                    batch.Requests.OfType<GoogleSheetsUpdateCellsRequest>(),
                    request => request.SheetName == chart.DataSheetName);
                Assert.Equal(new[] { 1D, 4D }, chartData.Cells
                    .Where(cell => cell.ColumnIndex == 0 && cell.RowIndex > chart.DataStartRowIndex)
                    .OrderBy(cell => cell.RowIndex)
                    .Select(cell => Assert.IsType<double>(cell.Value.Value))
                    .ToArray());
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatch_SkipsScatterSeriesWithDifferentXDomains() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsScatterDifferentDomains.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                ExcelSheet sheet = document.AddWorksheet("Dashboard");
                sheet.AddChart(
                    new ExcelChartData(
                        Array.Empty<string>(),
                        new[] {
                            new ExcelChartSeries("First", new[] { 10D, 20D }, new[] { 1D, 4D }, ExcelChartType.Scatter),
                            new ExcelChartSeries("Second", new[] { 30D, 40D }, new[] { 2D, 8D }, ExcelChartType.Scatter),
                        }),
                    row: 2,
                    column: 4,
                    type: ExcelChartType.Scatter,
                    title: "Different domains");

                GoogleSheetsBatch batch = new GoogleSheetsExporter().BuildBatch(document);

                Assert.Empty(batch.Requests.OfType<GoogleSheetsAddChartRequest>());
                Assert.Contains(batch.Report.Notices, notice => notice.Code == "SHEETS.CHART.SCATTER_X_VALUES_UNSUPPORTED");
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatch_DetectsMismatchedChartSeriesLengths() {
            var series = new ExcelChartSeries("Sales", new[] { 2d, 5d });
            var data = new ExcelChartData(new[] { "A", "B" }, new[] { series });
            Assert.IsType<List<double>>(series.Values).RemoveAt(1);

            Assert.False(GoogleSheetsBatchCompiler.HasAlignedChartSeries(data, dataRowCount: 2));
        }

        [Fact]
        public void Test_GoogleSheetsBatch_QualifiesSheetScopedNamedRanges() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsNamedRangeScopes.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                ExcelSheet north = document.AddWorksheet("North");
                ExcelSheet south = document.AddWorksheet("South");
                north.SetNamedRange("LocalData", "A1:A2", save: false);
                south.SetNamedRange("LocalData", "B1:B2", save: false);
                document.SetNamedRange("North_LocalData", "'North'!C1:C2", save: false);

                GoogleSheetsBatch batch = new GoogleSheetsExporter().BuildBatch(document);
                GoogleSheetsAddNamedRangeRequest[] names = batch.Requests.OfType<GoogleSheetsAddNamedRangeRequest>().ToArray();

                Assert.Equal(3, names.Select(name => name.Name).Distinct(StringComparer.OrdinalIgnoreCase).Count());
                Assert.Contains(names, name => name.SheetName == null && name.Name == "North_LocalData");
                Assert.Contains(names, name => name.SheetName == "North" && name.SourceName == "LocalData" && name.Name == "North_LocalData_2");
                Assert.Contains(names, name => name.SheetName == "South" && name.SourceName == "LocalData" && name.Name == "South_LocalData");
                Assert.Equal(2, batch.Report.Notices.Count(notice => notice.Code == "SHEETS.NAMED_RANGE.QUALIFIED"));
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

        [Fact]
        public void Test_GoogleSheetsPivot_MapsPopulationAggregatesExactly() {
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsPopulationPivot.xlsx");
            try {
                using var document = ExcelDocument.Create(path);
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Region");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "North");
                sheet.CellValue(2, 2, 10d);
                sheet.CellValue(3, 1, "South");
                sheet.CellValue(3, 2, 20d);
                sheet.CellValue(4, 1, "North");
                sheet.CellValue(4, 2, 30d);
                sheet.AddPivotTable(
                    "A1:B4",
                    "D1",
                    name: "PopulationPivot",
                    rowFields: new[] { "Region" },
                    dataFields: new[] {
                        new ExcelPivotDataField("Value", DataConsolidateFunctionValues.Count, "Non-empty count"),
                        new ExcelPivotDataField("Value", DataConsolidateFunctionValues.CountNumbers, "Numeric count"),
                        new ExcelPivotDataField("Value", DataConsolidateFunctionValues.StandardDeviationP, "Population SD"),
                        new ExcelPivotDataField("Value", DataConsolidateFunctionValues.VarianceP, "Population Variance"),
                    });

                GoogleSheetsBatch batch = document.BuildGoogleSheetsBatch();
                GoogleSheetsAddPivotTableRequest pivot = Assert.Single(batch.Requests.OfType<GoogleSheetsAddPivotTableRequest>());
                Assert.Equal(new[] { "COUNTA", "COUNT", "STDEVP", "VARP" }, pivot.Values.Select(value => value.SummarizeFunction).ToArray());
            } finally {
                if (File.Exists(path)) File.Delete(path);
            }
        }
    }
}
