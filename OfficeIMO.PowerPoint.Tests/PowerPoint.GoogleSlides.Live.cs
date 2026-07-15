using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.GoogleSlides;
using System;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class GoogleSlidesLiveTests {
        [GoogleWorkspaceLiveFact]
        [Trait("Category", "GoogleWorkspaceLive")]
        public async Task CreateReadExportAndDelete() {
            string token = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_ACCESS_TOKEN")!;
            string folderId = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_FOLDER_ID")!;
            string? driveId = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_DRIVE_ID");
            var session = new GoogleWorkspaceSession(new StaticAccessTokenCredentialSource(token), new GoogleWorkspaceSessionOptions {
                ApplicationName = "OfficeIMO.Tests", DefaultFolderId = folderId, DefaultDriveId = driveId,
            });
            string? id = null;
            try {
                using (PowerPointPresentation source = PowerPointPresentation.Create()) {
                    PowerPointSlide slide = source.AddSlide();
                    slide.AddTextBoxPoints("OfficeIMO Google Slides live round trip", 30, 40, 500, 80);
                    slide.AddTablePoints(1, 1, 30, 150, 300, 100).RowItems[0].Cells[0].Text = "Native table";
                    slide.Notes.Text = "Live speaker notes";
                    GooglePresentationReference created = await source.ExportToGoogleSlidesAsync(session, new GoogleSlidesSaveOptions {
                        Title = "OfficeIMO disposable Slides test " + Guid.NewGuid().ToString("N"),
                        Location = new GoogleDriveFileLocation { FolderId = folderId, DriveId = driveId, SharedDriveAware = true },
                    });
                    id = created.PresentationId;
                }
                GoogleSlidesImportResult native = await new GoogleSlidesImporter().ImportAsync(id!, session, new GoogleSlidesImportOptions { Mode = GoogleSlidesImportMode.Native });
                using (native.Presentation) Assert.Contains(native.Presentation.Slides[0].TextBoxes, box => box.Text.Contains("OfficeIMO Google Slides", StringComparison.Ordinal));
                GoogleSlidesImportResult broad = await new GoogleSlidesImporter().ImportAsync(id!, session, new GoogleSlidesImportOptions { Mode = GoogleSlidesImportMode.DriveExport });
                using (broad.Presentation) Assert.NotEmpty(broad.Presentation.Slides);
            } finally {
                if (!string.IsNullOrWhiteSpace(id)) { using var drive = new GoogleDriveClient(session); await drive.DeleteFileAsync(id!); }
            }
        }
    }
}
