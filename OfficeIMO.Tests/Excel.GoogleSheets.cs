using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using OfficeIMO.GoogleWorkspace;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ExcelInspectionSnapshot_ExposesOfficeIMOWorkbookModel() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelInspectionSnapshot.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var summary = document.AddWorkSheet("Summary");
                    var hidden = document.AddWorkSheet("Hidden");

                    summary.CellValue(1, 1, "Name");
                    summary.CellValue(2, 1, "Alpha");
                    summary.CellValue(2, 3, "Accent");
                    summary.CellFormula(2, 2, "SUM(1,2)");
                    summary.SetHyperlink(1, 1, "https://example.org", display: "Name");
                    summary.FormatCell(2, 1, "0.00%");
                    summary.CellBackground(2, 1, "#00FF00");
                    summary.CellBold(2, 1, true);
                    summary.CellAlign(2, 1, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
                    summary.CellFontColor(2, 3, "#112233");
                    summary.SetColumnWidth(1, 20);
                    summary.CellValue(3, 1, "First line\nSecond line");
                    summary.WrapCells(3, 3, 1, 20);
                    summary.AutoFitRow(3);
                    summary.CellValue(1, 4, "Status");
                    summary.CellValue(1, 5, "Region");
                    summary.CellValue(2, 4, "Open");
                    summary.CellValue(2, 5, "North");
                    summary.CellValue(3, 4, "Closed");
                    summary.CellValue(3, 5, "South");
                    summary.CellValue(4, 4, "Open");
                    summary.CellValue(4, 5, "East");
                    summary.AddTable("A1:B2", hasHeader: true, name: "SummaryTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);
                    summary.AddAutoFilter("D1:E4", new Dictionary<uint, IEnumerable<string>> {
                        { 0, new[] { "Open" } }
                    });
                    summary.Freeze(topRows: 1, leftCols: 1);
                    hidden.SetHidden(true);
                    document.SetNamedRange("GlobalData", "'Summary'!A1:B2", save: false);
                    document.Save();
                }

                ApplyBorderToCell(filePath, "Summary", "A2");

                using var reloadedDocument = ExcelDocument.Load(filePath);
                var snapshot = reloadedDocument.CreateInspectionSnapshot();

                Assert.Equal(2, snapshot.Worksheets.Count);

                var summarySheet = Assert.Single(snapshot.Worksheets, w => w.Name == "Summary");
                Assert.Equal(0, summarySheet.Index);
                Assert.False(summarySheet.Hidden);
                Assert.Equal(1, summarySheet.FrozenRowCount);
                Assert.Equal(1, summarySheet.FrozenColumnCount);
                Assert.NotNull(summarySheet.AutoFilter);
                Assert.Equal("D1:E4", summarySheet.AutoFilter!.A1Range);
                var worksheetFilterColumn = Assert.Single(summarySheet.AutoFilter.Columns);
                Assert.Equal(0, worksheetFilterColumn.ColumnId);
                Assert.Equal(new[] { "Open" }, worksheetFilterColumn.Values);
                var table = Assert.Single(summarySheet.Tables);
                Assert.Equal("SummaryTable", table.Name);
                Assert.Equal("A1:B2", table.A1Range);
                Assert.Equal("TableStyleMedium2", table.StyleName);
                Assert.True(table.HasHeaderRow);
                Assert.NotNull(table.AutoFilter);
                Assert.Equal("A1:B2", table.AutoFilter!.A1Range);
                Assert.Equal(new[] { "Name", "Column2" }, table.Columns.Select(c => c.Name).ToArray());
                Assert.Contains(summarySheet.Cells, c => c.Row == 2 && c.Column == 2 && c.Formula == "SUM(1,2)");
                var linkedCell = Assert.Single(summarySheet.Cells, c => c.Row == 1 && c.Column == 1);
                Assert.NotNull(linkedCell.Hyperlink);
                Assert.True(linkedCell.Hyperlink!.IsExternal);
                Assert.Equal("https://example.org", linkedCell.Hyperlink.Target);

                var styledCell = Assert.Single(summarySheet.Cells, c => c.Row == 2 && c.Column == 1);
                Assert.NotNull(styledCell.Style);
                var style = styledCell.Style!;
                Assert.Equal("0.00%", style.NumberFormatCode);
                Assert.True(style.Bold);
                Assert.Equal("FF00FF00", style.FillColorArgb);
                Assert.Equal("center", style.HorizontalAlignment);
                Assert.NotNull(style.Border);
                Assert.Equal("medium", style.Border!.Left!.Style);
                Assert.Equal("FFFF0000", style.Border.Left.ColorArgb);
                Assert.Equal("dashed", style.Border.Top!.Style);
                Assert.Equal("FF0000FF", style.Border.Top.ColorArgb);

                var fontColorCell = Assert.Single(summarySheet.Cells, c => c.Row == 2 && c.Column == 3);
                Assert.NotNull(fontColorCell.Style);
                Assert.Equal("FF112233", fontColorCell.Style!.FontColorArgb);

                var column = Assert.Single(summarySheet.Columns, c => c.StartIndex == 1 && c.EndIndex == 1);
                Assert.Equal(20, column.Width);

                var row = Assert.Single(summarySheet.Rows, r => r.Index == 3);
                Assert.True(row.Height > 15);

                var hiddenSheet = Assert.Single(snapshot.Worksheets, w => w.Name == "Hidden");
                Assert.True(hiddenSheet.Hidden);

                Assert.Contains(snapshot.NamedRanges, n => n.Name == "GlobalData" && n.ReferenceA1 == "'Summary'!$A$1:$B$2" && !n.IsBuiltIn);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatchCompiler_EmitsWorkbookStructureAndCellRequests() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsBatchCompiler.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var summary = document.AddWorkSheet("Summary");
                    var hidden = document.AddWorkSheet("Hidden");

                    summary.CellValue(1, 1, "Name");
                    summary.CellValue(1, 2, "Count");
                    summary.CellValue(2, 1, "Alpha");
                    summary.CellValue(2, 2, 12);
                    summary.CellValue(2, 3, true);
                    summary.CellValue(2, 4, new DateTime(2024, 12, 24, 10, 30, 0, DateTimeKind.Utc));
                    summary.CellValue(2, 6, "Accent");
                    summary.CellFormula(2, 5, "SUM(B2:B2)");
                    summary.SetHyperlink(2, 1, "https://alpha.example/", display: "Alpha");
                    summary.FormatCell(2, 2, "0.00%");
                    summary.CellBackground(2, 2, "#00FF00");
                    summary.CellBold(2, 2, true);
                    summary.CellAlign(2, 2, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
                    summary.CellFontColor(2, 6, "#112233");
                    summary.SetColumnWidth(2, 20);
                    summary.CellValue(3, 2, "Wrapped\nRow");
                    summary.WrapCells(3, 3, 2, 20);
                    summary.AutoFitRow(3);
                    summary.CellValue(1, 7, "Status");
                    summary.CellValue(1, 8, "Region");
                    summary.CellValue(2, 7, "Open");
                    summary.CellValue(2, 8, "North");
                    summary.CellValue(3, 7, "Closed");
                    summary.CellValue(3, 8, "South");
                    summary.CellValue(4, 7, "Open");
                    summary.CellValue(4, 8, "East");
                    summary.AddTable("A1:B3", hasHeader: true, name: "SummaryTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);
                    summary.AddAutoFilter("G1:H4", new Dictionary<uint, IEnumerable<string>> {
                        { 0, new[] { "Open" } }
                    });
                    summary.Freeze(topRows: 1, leftCols: 1);
                    hidden.SetHidden(true);

                    document.SetNamedRange("GlobalData", "'Summary'!A1:B2", save: false);
                    summary.SetNamedRange("LocalData", "A2:B2", save: false);
                    document.Save();
                }

                ApplyBorderToCell(filePath, "Summary", "B2");

                using var reloadedDocument = ExcelDocument.Load(filePath);
                var batch = reloadedDocument.CreateGoogleSheetsBatch(new GoogleSheetsSaveOptions {
                    Title = "OfficeIMO Export"
                });

                Assert.Equal("OfficeIMO Export", batch.Title);

                var addSheetRequests = batch.Requests.OfType<GoogleSheetsAddSheetRequest>().ToList();
                Assert.Equal(2, addSheetRequests.Count);

                var summaryRequest = Assert.Single(addSheetRequests, r => r.SheetName == "Summary");
                Assert.Equal(0, summaryRequest.SheetIndex);
                Assert.False(summaryRequest.Hidden);
                Assert.Equal(1, summaryRequest.FrozenRowCount);
                Assert.Equal(1, summaryRequest.FrozenColumnCount);

                var hiddenRequest = Assert.Single(addSheetRequests, r => r.SheetName == "Hidden");
                Assert.True(hiddenRequest.Hidden);

                var updateRequest = Assert.Single(batch.Requests.OfType<GoogleSheetsUpdateCellsRequest>(), r => r.SheetName == "Summary");
                Assert.Contains(updateRequest.Cells, c => c.RowIndex == 1 && c.ColumnIndex == 0 && Equals(c.Value.Value, "Alpha"));
                Assert.Contains(updateRequest.Cells, c => c.RowIndex == 1 && c.ColumnIndex == 1 && c.Value.Kind == GoogleSheetsCellValueKind.Number && Equals(c.Value.Value, 12d));
                Assert.Contains(updateRequest.Cells, c => c.RowIndex == 1 && c.ColumnIndex == 2 && c.Value.Kind == GoogleSheetsCellValueKind.Boolean && Equals(c.Value.Value, true));
                Assert.Contains(updateRequest.Cells, c => c.RowIndex == 1 && c.ColumnIndex == 3 && c.Value.Kind == GoogleSheetsCellValueKind.DateTime);
                Assert.Contains(updateRequest.Cells, c => c.RowIndex == 1 && c.ColumnIndex == 4 && c.Value.Kind == GoogleSheetsCellValueKind.Formula && Equals(c.Value.Value, "=SUM(B2:B2)"));
                var hyperlinkCell = Assert.Single(updateRequest.Cells, c => c.RowIndex == 1 && c.ColumnIndex == 0);
                Assert.NotNull(hyperlinkCell.Hyperlink);
                Assert.True(hyperlinkCell.Hyperlink!.IsExternal);
                Assert.Equal("https://alpha.example/", hyperlinkCell.Hyperlink.Target);

                var styledCell = Assert.Single(updateRequest.Cells, c => c.RowIndex == 1 && c.ColumnIndex == 1);
                Assert.NotNull(styledCell.Style);
                Assert.Equal("0.00%", styledCell.Style!.NumberFormatCode);
                Assert.True(styledCell.Style.Bold);
                Assert.Equal("FF00FF00", styledCell.Style.FillColorArgb);
                Assert.Equal("center", styledCell.Style.HorizontalAlignment);
                Assert.NotNull(styledCell.Style.Borders);
                Assert.Equal("medium", styledCell.Style.Borders!.Left!.Style);
                Assert.Equal("FFFF0000", styledCell.Style.Borders.Left.ColorArgb);
                Assert.Equal("dashed", styledCell.Style.Borders.Top!.Style);
                Assert.Equal("FF0000FF", styledCell.Style.Borders.Top.ColorArgb);

                var fontColorCell = Assert.Single(updateRequest.Cells, c => c.RowIndex == 1 && c.ColumnIndex == 5);
                Assert.NotNull(fontColorCell.Style);
                Assert.Equal("FF112233", fontColorCell.Style!.FontColorArgb);

                var dimensionRequests = batch.Requests.OfType<GoogleSheetsUpdateDimensionPropertiesRequest>().ToList();
                Assert.Contains(dimensionRequests, r => r.SheetName == "Summary" && r.DimensionKind == GoogleSheetsDimensionKind.Columns && r.StartIndex == 1 && r.EndIndexExclusive == 2 && r.PixelSize.HasValue && r.PixelSize.Value > 0);
                Assert.Contains(dimensionRequests, r => r.SheetName == "Summary" && r.DimensionKind == GoogleSheetsDimensionKind.Rows && r.StartIndex == 2 && r.EndIndexExclusive == 3 && r.PixelSize.HasValue && r.PixelSize.Value > 20);

                var tableRequest = Assert.Single(batch.Requests.OfType<GoogleSheetsAddTableRequest>(), r => r.SheetName == "Summary");
                Assert.Equal("SummaryTable", tableRequest.TableName);
                Assert.Equal("A1:B3", tableRequest.A1Range);
                Assert.Equal(new[] { "Name", "Count" }, tableRequest.Columns.Select(c => c.Name).ToArray());
                Assert.Equal("TEXT", tableRequest.Columns[0].ColumnType);
                Assert.Equal("TEXT", tableRequest.Columns[1].ColumnType);

                var basicFilter = Assert.Single(batch.Requests.OfType<GoogleSheetsSetBasicFilterRequest>(), r => r.SheetName == "Summary");
                Assert.Equal("G1:H4", basicFilter.A1Range);
                var basicFilterCriteria = Assert.Single(basicFilter.Criteria);
                Assert.Equal(0, basicFilterCriteria.ColumnId);
                Assert.Equal(new[] { "Closed" }, basicFilterCriteria.HiddenValues);

                var filterView = Assert.Single(batch.Requests.OfType<GoogleSheetsAddFilterViewRequest>(), r => r.SheetName == "Summary");
                Assert.Equal("SummaryTable Filter", filterView.Title);
                Assert.Equal("A1:B3", filterView.A1Range);

                var namedRanges = batch.Requests.OfType<GoogleSheetsAddNamedRangeRequest>().ToList();
                Assert.Equal(2, namedRanges.Count);
                Assert.Contains(namedRanges, r => r.Name == "GlobalData" && r.SheetName == null && r.A1Range == "'Summary'!$A$1:$B$2");
                Assert.Contains(namedRanges, r => r.Name == "LocalData" && r.SheetName == "Summary" && r.A1Range == "'Summary'!$A$2:$B$2");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleSheetsBatchCompiler_TreatsBuiltInNamesAsDiagnostics() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsBatchCompilerBuiltInNames.xlsx");

            try {
                using var document = ExcelDocument.Create(filePath);
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValue(1, 1, "Value");
                document.SetPrintArea(sheet, "A1:A5", save: false);

                var batch = document.CreateGoogleSheetsBatch();

                Assert.Empty(batch.Requests.OfType<GoogleSheetsAddNamedRangeRequest>());
                Assert.Contains(batch.Report.Notices, n => n.Feature == "BuiltInNames");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public void Test_GoogleSheetsApiPayloadBuilder_TranslatesNeutralBatchToSheetsPayloads() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsApiPayloadBuilder.xlsx");

            try {
                using (var document = ExcelDocument.Create(filePath)) {
                    var summary = document.AddWorkSheet("Summary");
                    var hidden = document.AddWorkSheet("Hidden");

                    summary.CellValue(1, 1, "Name");
                    summary.CellValue(2, 2, 12);
                    summary.SetHyperlink(2, 1, "https://alpha.example/", display: "Alpha");
                    summary.FormatCell(2, 2, "0.00%");
                    summary.CellBackground(2, 2, "#00FF00");
                    summary.CellBold(2, 2, true);
                    summary.CellAlign(2, 2, DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center);
                    summary.SetColumnWidth(2, 20);
                    summary.CellValue(3, 2, "Wrapped\nRow");
                    summary.WrapCells(3, 3, 2, 20);
                    summary.AutoFitRow(3);
                    summary.CellValue(1, 7, "Status");
                    summary.CellValue(1, 8, "Region");
                    summary.CellValue(2, 7, "Open");
                    summary.CellValue(2, 8, "North");
                    summary.CellValue(3, 7, "Closed");
                    summary.CellValue(3, 8, "South");
                    summary.CellValue(4, 7, "Open");
                    summary.CellValue(4, 8, "East");
                    summary.AddTable("A1:B3", hasHeader: true, name: "SummaryTable", style: OfficeIMO.Excel.TableStyle.TableStyleMedium2, includeAutoFilter: true);
                    summary.AddAutoFilter("G1:H4", new Dictionary<uint, IEnumerable<string>> {
                        { 0, new[] { "Open" } }
                    });
                    summary.Freeze(topRows: 1, leftCols: 1);
                    hidden.SetHidden(true);
                    document.SetNamedRange("GlobalData", "'Summary'!A1:B3", save: false);
                    document.Save();
                }

                ApplyBorderToCell(filePath, "Summary", "B2");

                using var reloadedDocument = ExcelDocument.Load(filePath);
                var batch = reloadedDocument.CreateGoogleSheetsBatch(new GoogleSheetsSaveOptions {
                    Title = "API Export"
                });

                var createPayload = GoogleSheetsApiPayloadBuilder.BuildCreateSpreadsheetPayload(batch);
                Assert.Equal("API Export", createPayload.Properties.Title);
                Assert.Equal(2, createPayload.Sheets.Count);
                Assert.Contains(createPayload.Sheets, s => s.Properties.SheetId == 1 && s.Properties.Title == "Summary" && s.Properties.GridProperties.FrozenRowCount == 1);
                Assert.Contains(createPayload.Sheets, s => s.Properties.SheetId == 2 && s.Properties.Title == "Hidden" && s.Properties.Hidden);

                var batchPayload = GoogleSheetsApiPayloadBuilder.BuildBatchUpdatePayload(batch);
                Assert.NotEmpty(batchPayload.Requests);

                Assert.Contains(batchPayload.Requests, r =>
                    r.UpdateDimensionProperties != null
                    && r.UpdateDimensionProperties.Range.Dimension == "COLUMNS"
                    && r.UpdateDimensionProperties.Range.SheetId == 1
                    && r.UpdateDimensionProperties.Range.StartIndex == 1);

                Assert.Contains(batchPayload.Requests, r =>
                    r.UpdateDimensionProperties != null
                    && r.UpdateDimensionProperties.Range.Dimension == "ROWS"
                    && r.UpdateDimensionProperties.Range.SheetId == 1
                    && r.UpdateDimensionProperties.Range.StartIndex == 2);

                var basicFilterRequest = Assert.Single(batchPayload.Requests, r => r.SetBasicFilter != null);
                Assert.Equal(1, basicFilterRequest.SetBasicFilter!.Filter.Range.SheetId);
                Assert.Equal(6, basicFilterRequest.SetBasicFilter.Filter.Range.StartColumnIndex);
                Assert.Contains("Closed", basicFilterRequest.SetBasicFilter.Filter.Criteria!["0"].HiddenValues!);

                var filterViewRequest = Assert.Single(batchPayload.Requests, r => r.AddFilterView != null);
                Assert.Equal("SummaryTable Filter", filterViewRequest.AddFilterView!.Filter.Title);
                Assert.Equal(0, filterViewRequest.AddFilterView.Filter.Range.StartColumnIndex);

                var tableRequest = Assert.Single(batchPayload.Requests, r => r.AddTable != null);
                Assert.Equal("SummaryTable", tableRequest.AddTable!.Table.Name);
                Assert.Equal(1, tableRequest.AddTable.Table.Range.SheetId);
                Assert.Equal("Name", tableRequest.AddTable.Table.ColumnProperties![0].Name);
                Assert.Equal("TEXT", tableRequest.AddTable.Table.ColumnProperties[1].ColumnType);

                var hyperlinkCell = batchPayload.Requests
                    .Where(r => r.UpdateCells != null)
                    .SelectMany(r => r.UpdateCells!.Rows)
                    .SelectMany(r => r.Values)
                    .First(c => c.UserEnteredValue?.FormulaValue != null && c.UserEnteredValue.FormulaValue.Contains("HYPERLINK", StringComparison.Ordinal));
                Assert.Contains("https://alpha.example/", hyperlinkCell.UserEnteredValue!.FormulaValue);

                var styledCell = batchPayload.Requests
                    .Where(r => r.UpdateCells != null)
                    .SelectMany(r => r.UpdateCells!.Rows)
                    .SelectMany(r => r.Values)
                    .First(c => c.UserEnteredFormat?.NumberFormat?.Pattern == "0.00%");
                Assert.Equal("PERCENT", styledCell.UserEnteredFormat!.NumberFormat!.Type);
                Assert.Equal("CENTER", styledCell.UserEnteredFormat.HorizontalAlignment);
                Assert.NotNull(styledCell.UserEnteredFormat.Borders);
                Assert.Equal("SOLID_MEDIUM", styledCell.UserEnteredFormat.Borders!.Left!.Style);
                Assert.Equal(1d, styledCell.UserEnteredFormat.Borders.Left.Color!.Red);
                Assert.Equal("DASHED", styledCell.UserEnteredFormat.Borders.Top!.Style);
                Assert.Equal(1d, styledCell.UserEnteredFormat.Borders.Top.Color!.Blue);
                Assert.Equal("WRAP", batchPayload.Requests
                    .Where(r => r.UpdateCells != null)
                    .SelectMany(r => r.UpdateCells!.Rows)
                    .SelectMany(r => r.Values)
                    .First(c => c.UserEnteredFormat?.WrapStrategy == "WRAP")
                    .UserEnteredFormat!.WrapStrategy);

                var namedRange = Assert.Single(batchPayload.Requests, r => r.AddNamedRange != null);
                Assert.Equal("GlobalData", namedRange.AddNamedRange!.NamedRange.Name);
                Assert.Equal(1, namedRange.AddNamedRange.NamedRange.Range.SheetId);
                Assert.Equal(0, namedRange.AddNamedRange.NamedRange.Range.StartRowIndex);
                Assert.Equal(2, namedRange.AddNamedRange.NamedRange.Range.EndColumnIndex);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleSheetsExporter_UsesConfiguredHttpPipeline_ForCreateFlow() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsExporterCreate.xlsx");

            try {
                using var document = ExcelDocument.Create(filePath);
                var summary = document.AddWorkSheet("Summary");
                summary.CellValue(1, 1, "Name");
                summary.SetHyperlink(2, 1, "https://alpha.example/", display: "Alpha");
                summary.CellValue(2, 2, 5);

                var recordedRequests = new List<(Uri Uri, string Method, string? Body, string? Authorization)>();
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body, request.Headers.Authorization?.ToString()));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets") {
                        return CreateJsonResponse("{\"spreadsheetId\":\"spread123\",\"spreadsheetUrl\":\"https://docs.google.com/spreadsheets/d/spread123/edit\",\"properties\":{\"title\":\"Create Export\"}}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets/spread123:batchUpdate") {
                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleSheetsAsync(session, new GoogleSheetsSaveOptions {
                    Title = "Create Export",
                });

                Assert.Equal("spread123", result.SpreadsheetId);
                Assert.Equal("https://docs.google.com/spreadsheets/d/spread123/edit", result.WebViewLink);
                Assert.Equal(2, recordedRequests.Count);
                Assert.All(recordedRequests, r => Assert.Equal("Bearer fake-access-token", r.Authorization));

                var createRequest = Assert.Single(recordedRequests, r => r.Uri.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets");
                Assert.Equal("POST", createRequest.Method);
                using (var json = JsonDocument.Parse(createRequest.Body!)) {
                    Assert.Equal("Create Export", json.RootElement.GetProperty("properties").GetProperty("title").GetString());
                }

                var updateRequest = Assert.Single(recordedRequests, r => r.Uri.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets/spread123:batchUpdate");
                Assert.Equal("POST", updateRequest.Method);
                Assert.Contains("HYPERLINK", updateRequest.Body!);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleSheetsExporter_MovesCreatedSpreadsheet_ToRequestedFolder() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsExporterCreateMove.xlsx");

            try {
                using var document = ExcelDocument.Create(filePath);
                var summary = document.AddWorkSheet("Summary");
                summary.CellValue(1, 1, "Name");
                summary.CellValue(2, 1, "Alpha");

                var recordedRequests = new List<(Uri Uri, string Method, string? Body)>();
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body));

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets") {
                        return CreateJsonResponse("{\"spreadsheetId\":\"spreadMove\",\"spreadsheetUrl\":\"https://docs.google.com/spreadsheets/d/spreadMove/edit\",\"properties\":{\"title\":\"Move Export\"}}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets/spreadMove:batchUpdate") {
                        return CreateJsonResponse("{}");
                    }

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri == "https://www.googleapis.com/drive/v3/files/spreadMove?fields=id,parents,webViewLink&supportsAllDrives=true") {
                        return CreateJsonResponse("{\"id\":\"spreadMove\",\"parents\":[\"oldParent\"],\"webViewLink\":\"https://docs.google.com/spreadsheets/d/spreadMove/edit\"}");
                    }

                    if (string.Equals(request.Method.Method, "PATCH", StringComparison.Ordinal) && request.RequestUri!.AbsoluteUri.Contains("https://www.googleapis.com/drive/v3/files/spreadMove?", StringComparison.Ordinal)) {
                        return CreateJsonResponse("{\"id\":\"spreadMove\",\"parents\":[\"folder123\"],\"webViewLink\":\"https://docs.google.com/spreadsheets/d/spreadMove/edit\"}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleSheetsAsync(session, new GoogleSheetsSaveOptions {
                    Title = "Move Export",
                    Location = new GoogleDriveFileLocation {
                        FolderId = "folder123",
                        SharedDriveAware = true,
                    }
                });

                Assert.Equal("spreadMove", result.SpreadsheetId);
                Assert.Equal(4, recordedRequests.Count);
                Assert.Contains(recordedRequests, r => r.Method == "GET" && r.Uri.AbsoluteUri.Contains("/drive/v3/files/spreadMove?", StringComparison.Ordinal));
                var patchRequest = Assert.Single(recordedRequests, r => r.Method == "PATCH");
                Assert.Contains("addParents=folder123", patchRequest.Uri.Query);
                Assert.Contains("removeParents=oldParent", patchRequest.Uri.Query);
                Assert.DoesNotContain(result.Report.Notices, n => n.Feature == "DrivePlacement" && n.Severity >= OfficeIMO.GoogleWorkspace.TranslationSeverity.Warning);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleSheetsExporter_CanReplaceExistingSpreadsheet() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsExporterUpdate.xlsx");

            try {
                using var document = ExcelDocument.Create(filePath);
                var summary = document.AddWorkSheet("Summary");
                summary.CellValue(1, 1, "Name");
                summary.CellValue(2, 1, "Alpha");
                summary.SetColumnWidth(1, 18);

                var recordedRequests = new List<(Uri Uri, string Method, string? Body)>();
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    string? body = request.Content == null ? null : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    recordedRequests.Add((request.RequestUri!, request.Method.Method, body));

                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri.Contains("/v4/spreadsheets/existing123?", StringComparison.Ordinal)) {
                        return CreateJsonResponse("{\"spreadsheetId\":\"existing123\",\"spreadsheetUrl\":\"https://docs.google.com/spreadsheets/d/existing123/edit\",\"properties\":{\"title\":\"Old Title\"},\"sheets\":[{\"properties\":{\"sheetId\":7}},{\"properties\":{\"sheetId\":8}}]}");
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets/existing123:batchUpdate") {
                        return CreateJsonResponse("{}");
                    }

                    return new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    };
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var result = await document.ExportToGoogleSheetsAsync(session, new GoogleSheetsSaveOptions {
                    Title = "Replacement Export",
                    Location = new GoogleDriveFileLocation {
                        ExistingFileId = "existing123",
                    }
                });

                Assert.Equal("existing123", result.SpreadsheetId);
                Assert.Equal("https://docs.google.com/spreadsheets/d/existing123/edit", result.WebViewLink);
                Assert.Equal(3, recordedRequests.Count);
                Assert.Equal("GET", recordedRequests[0].Method);
                Assert.Equal("POST", recordedRequests[1].Method);
                Assert.Equal("POST", recordedRequests[2].Method);
                Assert.DoesNotContain(recordedRequests, r => r.Uri.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets");
                Assert.Contains(result.Report.Notices, n => n.Feature == "ExistingSpreadsheet");

                using (var resetJson = JsonDocument.Parse(recordedRequests[1].Body!)) {
                    var requests = resetJson.RootElement.GetProperty("requests");
                    Assert.True(requests.GetArrayLength() >= 3);
                    var requestKinds = requests.EnumerateArray()
                        .SelectMany(r => r.EnumerateObject().Select(p => p.Name))
                        .ToList();
                    Assert.Contains("deleteSheet", requestKinds);
                    Assert.Contains("addSheet", requestKinds);
                    Assert.Contains("updateSpreadsheetProperties", requestKinds);
                }
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        private static HttpResponseMessage CreateJsonResponse(string json) {
            return new HttpResponseMessage(HttpStatusCode.OK) {
                Content = new StringContent(json, Encoding.UTF8, "application/json")
            };
        }

        private static void ApplyBorderToCell(string filePath, string sheetName, string cellReference) {
            using var document = SpreadsheetDocument.Open(filePath, true);
            var workbookPart = document.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is missing.");
            var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet ?? throw new InvalidOperationException("Stylesheet is missing.");
            var sheet = workbookPart.Workbook.Sheets?.Elements<Sheet>().FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.Ordinal))
                ?? throw new InvalidOperationException($"Sheet '{sheetName}' was not found.");
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
            var cell = worksheetPart.Worksheet.Descendants<Cell>().FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellReference, StringComparison.Ordinal))
                ?? throw new InvalidOperationException($"Cell '{cellReference}' was not found.");

            stylesheet.Borders ??= new Borders(new Border());
            stylesheet.Borders.Count = (uint)stylesheet.Borders.Count();
            stylesheet.CellFormats ??= new CellFormats(new CellFormat());
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();

            var border = new Border(
                new LeftBorder(new Color { Rgb = "FFFF0000" }) { Style = BorderStyleValues.Medium },
                new RightBorder(),
                new TopBorder(new Color { Rgb = "FF0000FF" }) { Style = BorderStyleValues.Dashed },
                new BottomBorder(),
                new DiagonalBorder());

            stylesheet.Borders.Append(border);
            stylesheet.Borders.Count = (uint)stylesheet.Borders.Count();
            var borderId = stylesheet.Borders.Count!.Value - 1;

            var existingStyleIndex = cell.StyleIndex?.Value ?? 0U;
            var existingFormat = stylesheet.CellFormats.Elements<CellFormat>().ElementAtOrDefault((int)existingStyleIndex) ?? new CellFormat();
            var clonedFormat = (CellFormat)existingFormat.CloneNode(true);
            clonedFormat.BorderId = borderId;
            clonedFormat.ApplyBorder = true;
            stylesheet.CellFormats.Append(clonedFormat);
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
            cell.StyleIndex = stylesheet.CellFormats.Count!.Value - 1;

            stylesheet.Save();
            worksheetPart.Worksheet.Save();
            workbookPart.Workbook.Save();
        }

        private sealed class FakeGoogleWorkspaceCredentialSource : IGoogleWorkspaceCredentialSource {
            public Task<GoogleWorkspaceAccessToken> AcquireAccessTokenAsync(IEnumerable<string> scopes, CancellationToken cancellationToken = default) {
                return Task.FromResult(new GoogleWorkspaceAccessToken(
                    "fake-access-token",
                    DateTimeOffset.UtcNow.AddHours(1),
                    scopes.ToList()));
            }
        }

        private sealed class FakeHttpMessageHandler : HttpMessageHandler {
            private readonly Func<HttpRequestMessage, Task<HttpResponseMessage>> _handler;

            public FakeHttpMessageHandler(Func<HttpRequestMessage, Task<HttpResponseMessage>> handler) {
                _handler = handler ?? throw new ArgumentNullException(nameof(handler));
            }

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
                return _handler(request);
            }
        }
    }
}
