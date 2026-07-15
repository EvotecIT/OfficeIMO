using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;
using OfficeIMO.Word;
using OfficeIMO.Word.GoogleDocs;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [GoogleWorkspaceLiveFact]
        [Trait("Category", "GoogleWorkspaceLive")]
        public async Task Test_GoogleDocsLive_CreateReadExportAndDelete() {
            string token = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_ACCESS_TOKEN")!;
            string folderId = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_FOLDER_ID")!;
            string? driveId = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_DRIVE_ID");
            var session = new GoogleWorkspaceSession(
                new StaticAccessTokenCredentialSource(token),
                new GoogleWorkspaceSessionOptions {
                    ApplicationName = "OfficeIMO.Tests",
                    DefaultFolderId = folderId,
                    DefaultDriveId = driveId,
                });
            string path = Path.Combine(_directoryWithFiles, "GoogleDocsLive.docx");
            string? fileId = null;
            try {
                using (var source = WordDocument.Create(path)) {
                    source.AddParagraph("OfficeIMO Google Docs live round trip").SetBold();
                    source.AddTable(2, 2, WordTableStyle.TableGrid).Rows[0].Cells[0].AddParagraph("Native table", removeExistingParagraphs: true);
                    GoogleDocumentReference created = await source.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                        Title = "OfficeIMO disposable Docs test " + Guid.NewGuid().ToString("N"),
                        Location = new GoogleDriveFileLocation { FolderId = folderId, DriveId = driveId, SharedDriveAware = true },
                    });
                    fileId = created.FileId;
                    Assert.False(string.IsNullOrWhiteSpace(fileId));
                }

                GoogleDocsImportResult native = await new GoogleDocsImporter().ImportAsync(fileId!, session, new GoogleDocsImportOptions { Mode = GoogleDocsImportMode.Native });
                using (native.Document) {
                    Assert.Contains(native.Document.Paragraphs, paragraph => paragraph.Text.Contains("OfficeIMO Google Docs live round trip", StringComparison.Ordinal));
                }
                GoogleDocsImportResult broad = await new GoogleDocsImporter().ImportAsync(fileId!, session, new GoogleDocsImportOptions { Mode = GoogleDocsImportMode.DriveExport });
                using (broad.Document) {
                    Assert.Contains(broad.Document.Paragraphs, paragraph => paragraph.Text.Contains("OfficeIMO Google Docs live round trip", StringComparison.Ordinal));
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
