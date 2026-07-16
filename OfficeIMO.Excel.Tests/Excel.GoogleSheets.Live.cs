using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [GoogleWorkspaceLiveFact]
        [Trait("Category", "GoogleWorkspaceLive")]
        public async Task Test_GoogleSheetsLive_CreateReadExportAndDelete() {
            string token = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_ACCESS_TOKEN")!;
            string folderId = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_FOLDER_ID")!;
            string? driveId = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_DRIVE_ID");
            var options = new GoogleWorkspaceSessionOptions {
                ApplicationName = "OfficeIMO.Tests",
                DefaultFolderId = folderId,
                DefaultDriveId = driveId,
            };
            var session = new GoogleWorkspaceSession(new StaticAccessTokenCredentialSource(token), options);
            string path = Path.Combine(_directoryWithFiles, "GoogleSheetsLive.xlsx");
            string? fileId = null;
            try {
                using (var source = ExcelDocument.Create(path)) {
                    ExcelSheet sheet = source.AddWorksheet("RoundTrip");
                    sheet.CellValue(1, 1, "Name");
                    sheet.CellValue(2, 1, "OfficeIMO");
                    sheet.CellFormula(2, 2, "SUM(1,2)");
                    GoogleSpreadsheetReference created = await source.ExportToGoogleSheetsAsync(session, new GoogleSheetsSaveOptions {
                        Title = "OfficeIMO disposable live test " + Guid.NewGuid().ToString("N"),
                        Location = new GoogleDriveFileLocation { FolderId = folderId, DriveId = driveId, SharedDriveAware = true },
                        Identity = new GoogleSheetsIdentityOptions { WriteDeveloperMetadata = true },
                    });
                    fileId = created.FileId;
                    Assert.False(string.IsNullOrWhiteSpace(fileId));
                }

                var importer = new GoogleSheetsImporter();
                GoogleSheetsImportResult native = await importer.ImportAsync(fileId!, session, new GoogleSheetsImportOptions { Mode = GoogleSheetsImportMode.Native });
                using (native.Document) {
                    Assert.Equal("OfficeIMO", native.Document.Sheets[0].CellAt(2, 1).GetValue<string>());
                }
                GoogleSheetsImportResult broad = await importer.ImportAsync(fileId!, session, new GoogleSheetsImportOptions { Mode = GoogleSheetsImportMode.DriveExport });
                using (broad.Document) {
                    Assert.Equal("OfficeIMO", broad.Document.Sheets[0].CellAt(2, 1).GetValue<string>());
                }
            } finally {
                if (!string.IsNullOrWhiteSpace(fileId)) {
                    using var drive = new GoogleDriveClient(session);
                    await drive.DeleteFileAsync(fileId!);
                }
                if (File.Exists(path)) File.Delete(path);
            }
        }
    }
}
