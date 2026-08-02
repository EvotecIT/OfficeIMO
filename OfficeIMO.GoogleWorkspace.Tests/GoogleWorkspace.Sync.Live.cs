using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;
using OfficeIMO.GoogleWorkspace.Sync;
using System;
using System.Linq;
using System.Collections.Generic;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
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
            var session = new GoogleWorkspaceSession(new DelegateGoogleWorkspaceCredentialSource(
                (scopes, cancellationToken) => VerifyLiveTokenAsync(token, account, scopes, cancellationToken)),
                sessionOptions);
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

        private static async Task<GoogleWorkspaceAccessToken> VerifyLiveTokenAsync(string token, string expectedAccount,
            IReadOnlyList<string> requiredScopes, CancellationToken cancellationToken) {
            using var http = new HttpClient();
            using HttpResponseMessage response = await http.GetAsync(
                "https://oauth2.googleapis.com/tokeninfo?access_token=" + Uri.EscapeDataString(token),
                cancellationToken).ConfigureAwait(false);
            if (!response.IsSuccessStatusCode) {
                throw new InvalidOperationException(
                    $"Google token verification failed with status code {(int)response.StatusCode}.");
            }
            using JsonDocument document = JsonDocument.Parse(await response.Content.ReadAsStringAsync().ConfigureAwait(false));
            string? account = document.RootElement.TryGetProperty("email", out JsonElement email)
                ? email.GetString()
                : null;
            string? granted = document.RootElement.TryGetProperty("scope", out JsonElement scope)
                ? scope.GetString()
                : null;
            if (!StringComparer.OrdinalIgnoreCase.Equals(account, expectedAccount)) {
                throw new InvalidOperationException("The live Google token identity does not match GOOGLE_WORKSPACE_ACCOUNT.");
            }
            var grantedScopes = new HashSet<string>((granted ?? string.Empty)
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries), StringComparer.Ordinal);
            if (requiredScopes.Any(required => !grantedScopes.Contains(required))) {
                throw new InvalidOperationException("The live Google token does not contain every scope required by the operation.");
            }
            return GoogleWorkspaceAccessToken.FromVerifiedCredential(token, DateTimeOffset.UtcNow.AddMinutes(5),
                new GoogleWorkspaceCredentialBinding(account!, requiredScopes));
        }
    }
}
