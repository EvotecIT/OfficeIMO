using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Sync;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class GoogleWorkspaceSyncTests {
        [Fact]
        public async Task ChangeTracker_ConsumesEveryPageForUserAndSharedDrive() {
            var changeUris = new List<string>();
            using var httpClient = new HttpClient(new FakeHandler(request => {
                string uri = request.RequestUri!.AbsoluteUri;
                if (uri.Contains("/changes?", StringComparison.Ordinal)) changeUris.Add(uri);
                if (uri.Contains("startPageToken", StringComparison.Ordinal) && uri.Contains("driveId=drive-a", StringComparison.Ordinal)) return Task.FromResult(Json("{\"startPageToken\":\"drive-start\"}"));
                if (uri.Contains("startPageToken", StringComparison.Ordinal)) return Task.FromResult(Json("{\"startPageToken\":\"user-start\"}"));
                if (uri.Contains("pageToken=user-start", StringComparison.Ordinal)) return Task.FromResult(Json("{\"changes\":[{\"fileId\":\"user-1\"}],\"nextPageToken\":\"user-page-2\"}"));
                if (uri.Contains("pageToken=user-page-2", StringComparison.Ordinal)) return Task.FromResult(Json("{\"changes\":[{\"fileId\":\"user-2\",\"removed\":true}],\"newStartPageToken\":\"user-next\"}"));
                if (uri.Contains("pageToken=drive-start", StringComparison.Ordinal) && uri.Contains("driveId=drive-a", StringComparison.Ordinal)) return Task.FromResult(Json("{\"changes\":[{\"fileId\":\"drive-1\",\"driveId\":\"drive-a\"}],\"newStartPageToken\":\"drive-next\"}"));
                return Task.FromResult(NotFound(uri));
            }));
            using var tracker = new GoogleWorkspaceChangeTracker(Session(httpClient));
            GoogleWorkspaceSyncCheckpoint checkpoint = await tracker.InitializeAsync(new[] { "drive-a" });
            var options = new GoogleWorkspaceChangeReadOptions();
            options.SharedDriveIds.Add("drive-a");

            GoogleWorkspaceChangeReadResult result = await tracker.ReadAsync(checkpoint, options);

            Assert.False(result.HasFailures);
            Assert.Equal(new[] { "user-1", "user-2", "drive-1" }, result.Changes.Select(change => change.Change.FileId));
            Assert.Equal("user-next", result.NextCheckpoint.UserChangeToken);
            Assert.Equal("drive-next", result.NextCheckpoint.SharedDriveChangeTokens["drive-a"]);
            Assert.Equal("user-start", checkpoint.UserChangeToken);
            Assert.Equal("drive-start", checkpoint.SharedDriveChangeTokens["drive-a"]);
            Assert.All(changeUris.Where(uri => !uri.Contains("driveId=drive-a", StringComparison.Ordinal)),
                uri => Assert.Contains("includeItemsFromAllDrives=false", uri, StringComparison.Ordinal));
            Assert.All(changeUris.Where(uri => uri.Contains("driveId=drive-a", StringComparison.Ordinal)),
                uri => Assert.Contains("includeItemsFromAllDrives=true", uri, StringComparison.Ordinal));
        }

        [Fact]
        public async Task ChangeTracker_RejectsPagesThatExceedConfiguredChangeBudget() {
            using var httpClient = new HttpClient(new FakeHandler(request =>
                Task.FromResult(request.RequestUri!.AbsoluteUri.Contains("pageToken=user-old", StringComparison.Ordinal)
                    ? Json("{\"changes\":[{\"fileId\":\"one\"},{\"fileId\":\"two\"}],\"newStartPageToken\":\"user-next\"}")
                    : NotFound(request.RequestUri.AbsoluteUri))));
            var checkpoint = new GoogleWorkspaceSyncCheckpoint { UserChangeToken = "user-old" };
            using var tracker = new GoogleWorkspaceChangeTracker(Session(httpClient));

            GoogleWorkspaceChangeReadResult result = await tracker.ReadAsync(checkpoint, new GoogleWorkspaceChangeReadOptions {
                MaxChangesPerSource = 1,
                MaxTotalChanges = 1,
            });

            GoogleWorkspaceChangeSourceResult source = Assert.Single(result.Sources);
            Assert.Equal(GoogleWorkspaceChangeReadStatus.Failed, source.Status);
            Assert.IsType<InvalidDataException>(source.Exception);
            Assert.Empty(result.Changes);
            Assert.Equal("user-old", result.NextCheckpoint.UserChangeToken);
        }

        [Fact]
        public async Task ChangeTracker_DoesNotAdvanceFailedSourceOrReturnItsPartialPages() {
            using var httpClient = new HttpClient(new FakeHandler(request => {
                string uri = request.RequestUri!.AbsoluteUri;
                if (uri.Contains("pageToken=user-old", StringComparison.Ordinal)) return Task.FromResult(Json("{\"changes\":[{\"fileId\":\"user-1\"}],\"newStartPageToken\":\"user-new\"}"));
                if (uri.Contains("pageToken=drive-old", StringComparison.Ordinal)) return Task.FromResult(new HttpResponseMessage(HttpStatusCode.InternalServerError) { Content = new StringContent("failed") });
                return Task.FromResult(NotFound(uri));
            }));
            var checkpoint = new GoogleWorkspaceSyncCheckpoint { UserChangeToken = "user-old" };
            checkpoint.SharedDriveChangeTokens["drive-a"] = "drive-old";
            using var tracker = new GoogleWorkspaceChangeTracker(Session(httpClient));

            GoogleWorkspaceChangeReadResult result = await tracker.ReadAsync(checkpoint);

            Assert.True(result.HasFailures);
            Assert.Equal("user-new", result.NextCheckpoint.UserChangeToken);
            Assert.Equal("drive-old", result.NextCheckpoint.SharedDriveChangeTokens["drive-a"]);
            Assert.Equal("user-1", Assert.Single(result.Changes).Change.FileId);
            Assert.Contains(result.Report.Notices, notice => notice.Code == "SYNC.CHANGES.SOURCE_FAILED" && notice.TargetId == "drive:drive-a");
        }

        [Fact]
        public async Task ChangeTracker_UsesUserFeedToCoverNewSharedDriveHistoryWithoutDuplicatingPartitionedDrives() {
            var changeUris = new List<string>();
            using var httpClient = new HttpClient(new FakeHandler(request => {
                string uri = request.RequestUri!.AbsoluteUri;
                if (uri.Contains("/changes?", StringComparison.Ordinal)) changeUris.Add(uri);
                if (uri.Contains("pageToken=user-old", StringComparison.Ordinal)) {
                    return Task.FromResult(Json("{\"changes\":[{\"fileId\":\"existing-drive-change\",\"driveId\":\"drive-a\"},{\"fileId\":\"new-drive-change\",\"driveId\":\"drive-b\"}],\"newStartPageToken\":\"user-next\"}"));
                }
                if (uri.Contains("pageToken=drive-a-old", StringComparison.Ordinal)) {
                    return Task.FromResult(Json("{\"changes\":[{\"fileId\":\"existing-drive-change\",\"driveId\":\"drive-a\"}],\"newStartPageToken\":\"drive-a-next\"}"));
                }
                if (uri.Contains("startPageToken", StringComparison.Ordinal) && uri.Contains("driveId=drive-b", StringComparison.Ordinal)) {
                    return Task.FromResult(Json("{\"startPageToken\":\"drive-b-start\"}"));
                }
                return Task.FromResult(NotFound(uri));
            }));
            var checkpoint = new GoogleWorkspaceSyncCheckpoint { UserChangeToken = "user-old" };
            checkpoint.SharedDriveChangeTokens["drive-a"] = "drive-a-old";
            var options = new GoogleWorkspaceChangeReadOptions();
            options.SharedDriveIds.Add("drive-b");
            using var tracker = new GoogleWorkspaceChangeTracker(Session(httpClient));

            GoogleWorkspaceChangeReadResult result = await tracker.ReadAsync(checkpoint, options);

            Assert.False(result.HasFailures);
            Assert.Equal(new[] { "new-drive-change", "existing-drive-change" }, result.Changes.Select(change => change.Change.FileId));
            Assert.Equal("drive-b-start", result.NextCheckpoint.SharedDriveChangeTokens["drive-b"]);
            Assert.Contains(changeUris, uri =>
                uri.Contains("pageToken=user-old", StringComparison.Ordinal)
                && uri.Contains("includeItemsFromAllDrives=true", StringComparison.Ordinal));
        }

        [Fact]
        public async Task Executor_DryRunBlocksConflictsAndUnapprovedLossyItems() {
            GoogleWorkspaceSyncPlan plan = Plan();
            int calls = 0;

            GoogleWorkspaceSyncApplyResult result = await GoogleWorkspaceSyncExecutor.ApplyAsync(plan, (item, token) => { calls++; return Task.CompletedTask; });

            Assert.Equal(0, calls);
            Assert.Equal(new[] {
                GoogleWorkspaceSyncApplyStatus.Planned,
                GoogleWorkspaceSyncApplyStatus.Conflict,
                GoogleWorkspaceSyncApplyStatus.ApprovalRequired,
                GoogleWorkspaceSyncApplyStatus.Planned,
            }, result.Items.Select(item => item.Status));
            Assert.True(result.HasConflicts);
            Assert.True(result.NeedsApproval);
        }

        [Theory]
        [InlineData(GoogleWorkspaceSyncItemKind.Conflict, GoogleWorkspaceSyncApplyStatus.Conflict)]
        [InlineData(GoogleWorkspaceSyncItemKind.LossyAction, GoogleWorkspaceSyncApplyStatus.ApprovalRequired)]
        public async Task Executor_MarksAppliedResultPartialWhenAnotherItemNeedsIntervention(
            GoogleWorkspaceSyncItemKind blockedKind,
            GoogleWorkspaceSyncApplyStatus expectedStatus) {
            GoogleWorkspaceSyncPlan plan = GoogleWorkspaceSyncPlan.Create(new[] {
                new GoogleWorkspaceSyncItem("source", GoogleWorkspaceSyncItemKind.SourceChange, "source", "local changed"),
                new GoogleWorkspaceSyncItem("blocked", blockedKind, "blocked", "review required"),
            });
            var options = new GoogleWorkspaceSyncApplyOptions { DryRun = false };

            GoogleWorkspaceSyncApplyResult result = await GoogleWorkspaceSyncExecutor.ApplyAsync(
                plan,
                (item, token) => Task.CompletedTask,
                options);

            Assert.Equal(GoogleWorkspaceSyncApplyStatus.Applied, result.Items[0].Status);
            Assert.Equal(expectedStatus, result.Items[1].Status);
            Assert.True(result.IsPartial);
        }

        [Fact]
        public async Task Executor_ReturnsAppliedAndFailedOutcomesWithoutLosingProgress() {
            GoogleWorkspaceSyncPlan plan = Plan();
            var options = new GoogleWorkspaceSyncApplyOptions { DryRun = false, ContinueOnError = true };
            options.ApprovedLossyItemIds.Add("lossy");

            GoogleWorkspaceSyncApplyResult result = await GoogleWorkspaceSyncExecutor.ApplyAsync(plan, (item, token) => {
                if (item.Id == "remote") throw new InvalidOperationException("target unavailable");
                return Task.CompletedTask;
            }, options);

            Assert.Equal(GoogleWorkspaceSyncApplyStatus.Applied, result.Items.Single(item => item.Item.Id == "source").Status);
            Assert.Equal(GoogleWorkspaceSyncApplyStatus.Conflict, result.Items.Single(item => item.Item.Id == "conflict").Status);
            Assert.Equal(GoogleWorkspaceSyncApplyStatus.Applied, result.Items.Single(item => item.Item.Id == "lossy").Status);
            Assert.Equal(GoogleWorkspaceSyncApplyStatus.Failed, result.Items.Single(item => item.Item.Id == "remote").Status);
            Assert.True(result.HasFailures);
            Assert.True(result.IsPartial);
        }

        [Fact]
        public async Task Executor_CancellationReturnsPartialOutcome() {
            GoogleWorkspaceSyncPlan plan = GoogleWorkspaceSyncPlan.Create(new[] {
                new GoogleWorkspaceSyncItem("first", GoogleWorkspaceSyncItemKind.SourceChange, "first", "first"),
                new GoogleWorkspaceSyncItem("second", GoogleWorkspaceSyncItemKind.SourceChange, "second", "second"),
            });
            using var cancellation = new CancellationTokenSource();
            var options = new GoogleWorkspaceSyncApplyOptions { DryRun = false };

            GoogleWorkspaceSyncApplyResult result = await GoogleWorkspaceSyncExecutor.ApplyAsync(plan, (item, token) => {
                if (item.Id == "first") { cancellation.Cancel(); return Task.CompletedTask; }
                return Task.CompletedTask;
            }, options, cancellation.Token);

            Assert.True(result.WasCanceled);
            Assert.True(result.IsPartial);
            Assert.Equal(GoogleWorkspaceSyncApplyStatus.Applied, result.Items[0].Status);
            Assert.Equal(GoogleWorkspaceSyncApplyStatus.Canceled, result.Items[1].Status);
        }

        private static GoogleWorkspaceSyncPlan Plan() => GoogleWorkspaceSyncPlan.Create(new[] {
            new GoogleWorkspaceSyncItem("source", GoogleWorkspaceSyncItemKind.SourceChange, "source", "local changed"),
            new GoogleWorkspaceSyncItem("conflict", GoogleWorkspaceSyncItemKind.Conflict, "conflict", "both changed"),
            new GoogleWorkspaceSyncItem("lossy", GoogleWorkspaceSyncItemKind.LossyAction, "lossy", "render required"),
            new GoogleWorkspaceSyncItem("remote", GoogleWorkspaceSyncItemKind.RemoteChange, "remote", "remote changed"),
        });

        private static GoogleWorkspaceSession Session(HttpClient client) => new GoogleWorkspaceSession(new StaticAccessTokenCredentialSource("token"), new GoogleWorkspaceSessionOptions { HttpClient = client, MaxRetryCount = 1 });
        private static HttpResponseMessage Json(string value) => new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(value, Encoding.UTF8, "application/json") };
        private static HttpResponseMessage NotFound(string uri) => new HttpResponseMessage(HttpStatusCode.NotFound) { Content = new StringContent("unexpected " + uri) };

        private sealed class FakeHandler : HttpMessageHandler {
            private readonly Func<HttpRequestMessage, Task<HttpResponseMessage>> _handler;
            public FakeHandler(Func<HttpRequestMessage, Task<HttpResponseMessage>> handler) { _handler = handler; }
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => _handler(request);
        }
    }
}
