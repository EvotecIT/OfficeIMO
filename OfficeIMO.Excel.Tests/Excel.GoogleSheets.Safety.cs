using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using OfficeIMO.GoogleWorkspace;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task Test_GoogleSheetsExporter_ValidatesTargetFolderBeforeCreateOrReplace(bool replaceExisting) {
            string filePath = Path.Combine(_directoryWithFiles, $"GoogleSheetsFolderPreflight-{replaceExisting}.xlsx");
            try {
                using var document = ExcelDocument.Create(filePath);
                document.AddWorksheet("Data").CellValue(1, 1, "Content");
                int sheetsMutationCount = 0;
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                    if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "www.googleapis.com") {
                        return Task.FromResult(CreateJsonResponse("{\"id\":\"not-a-folder\",\"name\":\"File\",\"mimeType\":\"text/plain\"}"));
                    }
                    if (request.Method == HttpMethod.Post && request.RequestUri!.Host == "sheets.googleapis.com") {
                        sheetsMutationCount++;
                    }
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
                }));
                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                var options = new GoogleSheetsSaveOptions {
                    Location = new GoogleDriveFileLocation {
                        FolderId = "not-a-folder",
                        ExistingFileId = replaceExisting ? "existing-sheet" : null,
                    },
                    Replace = new GoogleSheetsReplaceOptions { ConflictMode = GoogleSheetsReplaceConflictMode.Overwrite },
                };
                await Assert.ThrowsAsync<GoogleWorkspaceExportException>(() => document.ExportToGoogleSheetsAsync(session, options));

                Assert.Equal(0, sheetsMutationCount);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }
    }
}
