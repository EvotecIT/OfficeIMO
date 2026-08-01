using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;
using OfficeIMO.GoogleWorkspace.Sync;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class GoogleWorkspaceSyncLiveTests {
        [GoogleWorkspaceLiveFact("GOOGLE_WORKSPACE_ACCOUNT", "GOOGLE_WORKSPACE_MY_DRIVE_FOLDER_ID")]
        [Trait("Category", "GoogleWorkspaceLive")]
        public Task ChangeTracker_ObservesAndCleansDisposableFileInMyDrive() =>
            ExerciseDisposableChangeAsync(Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_MY_DRIVE_FOLDER_ID")!, null);

        [GoogleWorkspaceLiveFact("GOOGLE_WORKSPACE_ACCOUNT", "GOOGLE_WORKSPACE_SHARED_DRIVE_ID", "GOOGLE_WORKSPACE_SHARED_DRIVE_FOLDER_ID")]
        [Trait("Category", "GoogleWorkspaceLive")]
        public Task ChangeTracker_ObservesAndCleansDisposableFileInSharedDrive() =>
            ExerciseDisposableChangeAsync(Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_SHARED_DRIVE_FOLDER_ID")!,
                Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_SHARED_DRIVE_ID")!);

        private static async Task ExerciseDisposableChangeAsync(string folderId, string? driveId) {
            string token = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_ACCESS_TOKEN")!;
            string account = Environment.GetEnvironmentVariable("GOOGLE_WORKSPACE_ACCOUNT")!;
            var receipts = new List<GoogleWorkspaceOperationReceipt>();
            string expectedRevision = GoogleWorkspaceOperationPolicy.ResourceAbsentForCreateRevision;
            var sessionOptions = new GoogleWorkspaceSessionOptions {
                ApplicationName = "OfficeIMO.Tests", DefaultFolderId = folderId, DefaultDriveId = driveId,
                ExpectedAccount = account, OperationReceiptSink = receipts.Add,
            };
            sessionOptions.OperationPolicyProvider = context => {
                bool isDelete = string.Equals(context.Method, "DELETE", StringComparison.OrdinalIgnoreCase);
                return new GoogleWorkspaceOperationPolicy(
                    account, context.RequiredScopes, context.Target,
                    expectedRevision, context.MaxRetryCount,
                    context.MaxRetryElapsedTime, context.RateLimitPolicy,
                    isDelete ? GoogleWorkspaceDataLossDecision.AcceptSpecifiedLoss : GoogleWorkspaceDataLossDecision.RejectPotentialLoss,
                    isDelete ? "the disposable live-test folder" : null);
            };
            var session = new GoogleWorkspaceSession(new StaticAccessTokenCredentialSource(token), sessionOptions);
            using var tracker = new GoogleWorkspaceChangeTracker(session);
            GoogleWorkspaceSyncCheckpoint checkpoint = await tracker.InitializeAsync(string.IsNullOrWhiteSpace(driveId) ? null : new[] { driveId! });
            string? fileId = null;
            try {
                using (var drive = new GoogleDriveClient(session)) {
                    GoogleDriveFile created = await drive.CreateFolderAsync("OfficeIMO disposable sync test " + Guid.NewGuid().ToString("N"), folderId);
                    fileId = created.Id;
                    expectedRevision = GoogleWorkspaceOperationPolicy.ExplicitlyUnversionedRevision(
                        "Drive API v3 exposes file version " + (created.Version?.ToString() ?? "not-returned")
                        + " but no conditional mutation precondition for files.delete");
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
                Assert.Contains(receipts, receipt => receipt.Succeeded &&
                    receipt.Policy.ExpectedRevision == GoogleWorkspaceOperationPolicy.ResourceAbsentForCreateRevision);
            } finally {
                if (!string.IsNullOrWhiteSpace(fileId)) { using var drive = new GoogleDriveClient(session); await drive.DeleteFileAsync(fileId!); }
            }
            Assert.Contains(receipts, receipt => receipt.Succeeded && receipt.Method == "DELETE" &&
                receipt.Policy.ExpectedRevision == expectedRevision &&
                receipt.Policy.DataLossDecision == GoogleWorkspaceDataLossDecision.AcceptSpecifiedLoss);
        }
    }
}
