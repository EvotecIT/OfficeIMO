using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;
using OfficeIMO.GoogleWorkspace.Sync;
using System;
using System.Linq;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class GoogleWorkspaceSyncLiveTests {
        [GoogleWorkspaceLiveFact]
        [Trait("Category", "GoogleWorkspaceLive")]
        public async Task ChangeTracker_ObservesAndCleansDisposableFile() {
            string token = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_ACCESS_TOKEN")!;
            string folderId = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_FOLDER_ID")!;
            string? driveId = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_DRIVE_ID");
            var session = new GoogleWorkspaceSession(new StaticAccessTokenCredentialSource(token), new GoogleWorkspaceSessionOptions {
                ApplicationName = "OfficeIMO.Tests", DefaultFolderId = folderId, DefaultDriveId = driveId,
            });
            using var tracker = new GoogleWorkspaceChangeTracker(session);
            GoogleWorkspaceSyncCheckpoint checkpoint = await tracker.InitializeAsync(string.IsNullOrWhiteSpace(driveId) ? null : new[] { driveId! });
            string? fileId = null;
            try {
                using (var drive = new GoogleDriveClient(session)) {
                    GoogleDriveFile created = await drive.CreateFolderAsync("OfficeIMO disposable sync test " + Guid.NewGuid().ToString("N"), folderId);
                    fileId = created.Id;
                }
                Assert.False(string.IsNullOrWhiteSpace(fileId));
                var options = new GoogleWorkspaceChangeReadOptions();
                if (!string.IsNullOrWhiteSpace(driveId)) options.SharedDriveIds.Add(driveId!);
                bool observed = false;
                for (int attempt = 0; attempt < 5 && !observed; attempt++) {
                    GoogleWorkspaceChangeReadResult changes = await tracker.ReadAsync(checkpoint, options);
                    checkpoint = changes.NextCheckpoint;
                    observed = changes.Changes.Any(change => string.Equals(change.Change.FileId, fileId, StringComparison.Ordinal));
                    if (!observed) await Task.Delay(TimeSpan.FromSeconds(2));
                }
                Assert.True(observed, "The Drive change feed did not expose the disposable folder within the live-test window.");
            } finally {
                if (!string.IsNullOrWhiteSpace(fileId)) { using var drive = new GoogleDriveClient(session); await drive.DeleteFileAsync(fileId!); }
            }
        }
    }
}
