using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using OfficeIMO.GoogleWorkspace;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public async Task Test_GoogleSheetsImporter_DriveExport_LoadsBroadXlsxFallback() {
            string sourcePath = Path.Combine(_directoryWithFiles, "GoogleSheetsDriveImportSource.xlsx");
            try {
                using (var source = ExcelDocument.Create(sourcePath)) {
                    var sheet = source.AddWorksheet("Summary");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(2, 1, "Alpha");
                    source.Save();
                }
                byte[] package = File.ReadAllBytes(sourcePath);

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsolutePath.EndsWith("/files/sheet123", StringComparison.Ordinal)) {
                        return Task.FromResult(CreateJsonResponse("{\"id\":\"sheet123\",\"name\":\"Imported\",\"mimeType\":\"application/vnd.google-apps.spreadsheet\",\"version\":17,\"modifiedTime\":\"2026-07-15T10:00:00Z\",\"capabilities\":{\"canDownload\":true}}"));
                    }
                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsolutePath.EndsWith("/files/sheet123/export", StringComparison.Ordinal)) {
                        return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                            Content = new ByteArrayContent(package),
                        });
                    }
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
                }));
                var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                GoogleSheetsImportResult result = await session.ImportGoogleSheetAsync("sheet123");
                using (result.Document) {
                    ExcelWorkbookSnapshot snapshot = result.Document.CreateInspectionSnapshot();
                    Assert.Contains(snapshot.Worksheets.Single().Cells, cell => cell.Row == 2 && cell.Column == 1 && Equals(cell.Value, "Alpha"));
                }
                Assert.Equal(17, result.Source.DriveVersion);
                Assert.Contains(result.Report.Notices, notice => notice.Code == "SHEETS.IMPORT.DRIVE_EXPORT");
            } finally {
                if (File.Exists(sourcePath)) File.Delete(sourcePath);
            }
        }

        [Fact]
        public async Task Test_GoogleSheetsImporter_Native_ProjectsContentWhenDriveExportIsDisabled() {
            Uri? sheetsRequest = null;
            const string nativeJson = """
                {
                  "spreadsheetId":"native123",
                  "spreadsheetUrl":"https://docs.google.com/spreadsheets/d/native123/edit",
                  "properties":{"title":"Native","locale":"en_US","timeZone":"Europe/Warsaw"},
                  "namedRanges":[{"name":"InputData","range":{"sheetId":42,"startRowIndex":0,"endRowIndex":2,"startColumnIndex":0,"endColumnIndex":2}}],
                  "sheets":[{
                    "properties":{"sheetId":42,"title":"Summary","index":0,"rightToLeft":true,"tabColor":{"red":0.2,"green":0.4,"blue":0.6},"gridProperties":{"frozenRowCount":1,"hideGridlines":true}},
                    "data":[{"startRow":0,"startColumn":0,"rowData":[
                      {"values":[{"userEnteredValue":{"stringValue":"Name"},"userEnteredFormat":{"textFormat":{"bold":true,"fontFamily":"Arial","fontSize":12},"backgroundColor":{"red":1,"green":1,"blue":0}}}]},
                      {"values":[{"userEnteredValue":{"stringValue":"Alpha"},"note":"Imported note"},{"userEnteredValue":{"formulaValue":"=SUM(1,2)"},"effectiveValue":{"numberValue":3},"userEnteredFormat":{"numberFormat":{"type":"NUMBER","pattern":"0.00"},"textRotation":{"angle":-45}}}]}
                    ]}],
                    "merges":[{"sheetId":42,"startRowIndex":0,"endRowIndex":1,"startColumnIndex":0,"endColumnIndex":2}],
                    "charts":[{}]
                  }]
                }
                """;
            using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                if (request.RequestUri!.Host == "www.googleapis.com") {
                    return Task.FromResult(CreateJsonResponse("{\"id\":\"native123\",\"name\":\"Native\",\"mimeType\":\"application/vnd.google-apps.spreadsheet\",\"version\":21,\"capabilities\":{\"canDownload\":false}}"));
                }
                if (request.RequestUri.Host == "sheets.googleapis.com") {
                    sheetsRequest = request.RequestUri;
                    return Task.FromResult(CreateJsonResponse(nativeJson));
                }
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }));
            var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

            GoogleSheetsImportResult result = await session.ImportGoogleSheetAsync("native123", new GoogleSheetsImportOptions {
                Mode = GoogleSheetsImportMode.Native,
                Ranges = new[] { "Summary!A1:B2" },
            });

            using (result.Document) {
                ExcelWorkbookSnapshot snapshot = result.Document.CreateInspectionSnapshot();
                ExcelWorksheetSnapshot sheet = Assert.Single(snapshot.Worksheets);
                Assert.True(sheet.RightToLeft);
                Assert.False(sheet.ShowGridlines);
                Assert.Equal(1, sheet.FrozenRowCount);
                Assert.Contains(sheet.MergedRanges, merge => merge.A1Range == "A1:B1");
                Assert.Contains(snapshot.NamedRanges, range => range.Name == "InputData");
                Assert.Contains(sheet.Cells, cell => cell.Row == 2 && cell.Column == 2 && cell.Formula == "SUM(1,2)");
                ExcelCellSnapshot header = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1);
                Assert.True(header.Style!.Bold);
                Assert.Equal("Arial", header.Style.FontName);
                Assert.Equal("FFFFFF00", header.Style.FillColorArgb);
                ExcelCellSnapshot formula = Assert.Single(sheet.Cells, cell => cell.Row == 2 && cell.Column == 2);
                Assert.Equal("0.00", formula.Style!.NumberFormatCode);
                Assert.Equal(135, formula.Style.TextRotation);
            }

            Assert.NotNull(sheetsRequest);
            Assert.Contains("ranges=Summary%21A1%3AB2", sheetsRequest!.Query);
            Assert.Equal(21, result.Source.DriveVersion);
            Assert.Contains(result.Report.Notices, notice => notice.Code == "SHEETS.IMPORT.CHARTS_FALLBACK");
        }

        [Fact]
        public async Task Test_GoogleSheetsDiffPlanner_ReportsRemoteDriveVersionChange() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsDriveVersionDiff.xlsx");
            try {
                using var source = ExcelDocument.Create(filePath);
                source.AddWorksheet("Data").CellValue(1, 1, "value");
                GoogleSheetsSyncCheckpoint checkpoint = GoogleSheetsDiffPlanner.CreateCheckpoint(source, driveVersion: 5);
                const string nativeJson = "{\"spreadsheetId\":\"diff-sheet\",\"properties\":{\"title\":\"Diff\"},\"sheets\":[{\"properties\":{\"sheetId\":0,\"title\":\"Data\",\"index\":0},\"data\":[{\"startRow\":0,\"startColumn\":0,\"rowData\":[{\"values\":[{\"userEnteredValue\":{\"stringValue\":\"value\"}}]}]}]}]}";
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                    if (request.RequestUri!.Host == "www.googleapis.com") {
                        return Task.FromResult(CreateJsonResponse("{\"id\":\"diff-sheet\",\"name\":\"Diff\",\"mimeType\":\"application/vnd.google-apps.spreadsheet\",\"version\":6,\"capabilities\":{\"canDownload\":false}}"));
                    }
                    if (request.RequestUri.Host == "sheets.googleapis.com") {
                        return Task.FromResult(CreateJsonResponse(nativeJson));
                    }
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
                }));
                var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                GoogleSheetsDiffPlan plan = await GoogleSheetsDiffPlanner.BuildAsync(source, "diff-sheet", session, checkpoint);

                GoogleSheetsDiffItem versionChange = Assert.Single(plan.Items, item => item.Path == "spreadsheet/driveVersion");
                Assert.Equal(GoogleSheetsDiffKind.RemoteChange, versionChange.Kind);
                Assert.Equal(6, plan.Remote.DriveVersion);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public async Task Test_GoogleSheetsExporter_ReplaceRequiresObservedDriveVersion() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsReplacePreflight.xlsx");
            try {
                using var document = ExcelDocument.Create(filePath);
                document.AddWorksheet("Data").CellValue(1, 1, "value");
                int requestCount = 0;
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                    requestCount++;
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.InternalServerError));
                }));
                var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                GoogleWorkspacePreflightException exception = await Assert.ThrowsAsync<GoogleWorkspacePreflightException>(() =>
                    document.ExportToGoogleSheetsAsync(session, new GoogleSheetsSaveOptions {
                        Location = new GoogleDriveFileLocation { ExistingFileId = "existing" },
                    }));

                Assert.Equal(0, requestCount);
                Assert.Contains(exception.BlockingNotices, notice => notice.Code == "SHEETS.REPLACE.EXPECTED_VERSION_REQUIRED");
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public async Task Test_GoogleSheetsExporter_ReplaceRejectsVersionConflictBeforeSheetsMutation() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsReplaceConflict.xlsx");
            try {
                using var document = ExcelDocument.Create(filePath);
                document.AddWorksheet("Data").CellValue(1, 1, "value");
                int sheetsRequests = 0;
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                    if (request.RequestUri!.Host == "www.googleapis.com") {
                        return Task.FromResult(CreateJsonResponse("{\"id\":\"existing\",\"mimeType\":\"application/vnd.google-apps.spreadsheet\",\"version\":6,\"capabilities\":{\"canEdit\":true}}"));
                    }
                    sheetsRequests++;
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.InternalServerError));
                }));
                var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                GoogleWorkspaceConflictException exception = await Assert.ThrowsAsync<GoogleWorkspaceConflictException>(() =>
                    document.ExportToGoogleSheetsAsync(session, new GoogleSheetsSaveOptions {
                        Location = new GoogleDriveFileLocation { ExistingFileId = "existing" },
                        Replace = new GoogleSheetsReplaceOptions { ExpectedDriveVersion = 5 },
                    }));

                Assert.Equal("5", exception.ExpectedVersion);
                Assert.Equal("6", exception.ActualVersion);
                Assert.Equal(0, sheetsRequests);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void Test_GoogleSheetsFormulaCatalog_RewritesKnownPrefixesAndDetectsUnsupportedFunctions() {
            GoogleSheetsFormulaTranslation translated = GoogleSheetsFormulaCatalog.Translate(
                "=_xlfn.XLOOKUP(A1,B:B,C:C)&\"WEBSERVICE(ignored)\"");
            GoogleSheetsFormulaTranslation unsupported = GoogleSheetsFormulaCatalog.Translate("=WEBSERVICE(A1)");

            Assert.Equal("=XLOOKUP(A1,B:B,C:C)&\"WEBSERVICE(ignored)\"", translated.Formula);
            Assert.True(translated.IsSupported);
            Assert.Equal(new[] { "XLOOKUP" }, translated.Functions);
            Assert.False(unsupported.IsSupported);
            Assert.Equal(new[] { "WEBSERVICE" }, unsupported.UnsupportedFunctions);
        }

        [Fact]
        public async Task Test_GoogleSheetsExporter_FormulaErrorPolicyBlocksBeforeMutation() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsFormulaPreflight.xlsx");
            try {
                using var document = ExcelDocument.Create(filePath);
                document.AddWorksheet("Data").CellFormula(1, 1, "WEBSERVICE(A2)");
                int requestCount = 0;
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                    requestCount++;
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.InternalServerError));
                }));
                var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                GoogleWorkspacePreflightException exception = await Assert.ThrowsAsync<GoogleWorkspacePreflightException>(() =>
                    document.ExportToGoogleSheetsAsync(session, new GoogleSheetsSaveOptions {
                        Formulas = new GoogleSheetsFormulaOptions { UnsupportedFormulaMode = GoogleSheetsUnsupportedFormulaMode.Error },
                    }));

                Assert.Equal(0, requestCount);
                Assert.Contains(exception.BlockingNotices, notice => notice.Code == "SHEETS.FORMULA.UNSUPPORTED");
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public async Task Test_GoogleSheetsExporter_ChunksSparseValueRangesAndReportsProgress() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleSheetsValueChunks.xlsx");
            try {
                using var document = ExcelDocument.Create(filePath);
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "A");
                sheet.CellValue(3, 1, "B");
                sheet.CellValue(5, 1, "C");
                int valueCalls = 0;
                var progress = new List<GoogleSheetsExportProgress>();
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                    if (request.RequestUri!.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets") {
                        return Task.FromResult(CreateJsonResponse("{\"spreadsheetId\":\"chunk123\"}"));
                    }
                    if (request.RequestUri.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets/chunk123/values:batchUpdate") {
                        valueCalls++;
                        return Task.FromResult(CreateJsonResponse("{}"));
                    }
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
                }));
                var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                GoogleSpreadsheetReference result = await document.ExportToGoogleSheetsAsync(session, new GoogleSheetsSaveOptions {
                    Execution = new GoogleSheetsExecutionOptions {
                        MaxValueRangesPerRequest = 2,
                        Progress = new ImmediateProgress<GoogleSheetsExportProgress>(item => progress.Add(item)),
                    },
                });

                Assert.Equal(2, valueCalls);
                Assert.Equal(3, progress.Last().Completed);
                Assert.Equal(3, progress.Last().Total);
                Assert.Contains(result.Report.Notices, notice => notice.Code == "SHEETS.VALUES.BATCH_UPDATE" && notice.Count == 3);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        private sealed class ImmediateProgress<T> : IProgress<T> {
            private readonly Action<T> _report;
            public ImmediateProgress(Action<T> report) => _report = report;
            public void Report(T value) => _report(value);
        }
    }
}
