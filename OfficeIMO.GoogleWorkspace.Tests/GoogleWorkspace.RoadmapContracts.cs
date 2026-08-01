using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class GoogleWorkspaceRoadmapContractTests {
        [Fact]
        public async Task MutationRequiresExplicitPolicyAndEmitsReceipt() {
            using var http = new HttpClient(new Handler(_ => Json("{\"id\":\"folder-1\",\"version\":\"1\"}")));
            using (var unguarded = new GoogleDriveClient(new GoogleWorkspaceSession(
                new StaticAccessTokenCredentialSource("token"), new GoogleWorkspaceSessionOptions { HttpClient = http }))) {
                await Assert.ThrowsAsync<InvalidOperationException>(() => unguarded.CreateFolderAsync("blocked"));
            }

            var receipts = new List<GoogleWorkspaceOperationReceipt>();
            GoogleWorkspaceSession session = Session(http, receipts);
            using var guarded = new GoogleDriveClient(session);
            GoogleDriveFile file = await guarded.CreateFolderAsync("allowed");

            Assert.Equal("folder-1", file.Id);
            GoogleWorkspaceOperationReceipt receipt = Assert.Single(receipts);
            Assert.True(receipt.Succeeded);
            Assert.Equal("test@example.com", receipt.Policy.Account);
            Assert.NotEmpty(receipt.Policy.Scopes);
            Assert.Equal("explicit-test-revision", receipt.Policy.ExpectedRevision);
            Assert.Equal(GoogleWorkspaceDataLossDecision.RejectPotentialLoss, receipt.Policy.DataLossDecision);
        }

        [Fact]
        public async Task DeleteRequiresExplicitAcceptanceOfDataLoss() {
            int requests = 0;
            using var http = new HttpClient(new Handler(_ => { requests++; return Json("{}"); }));
            using (var refusing = new GoogleDriveClient(Session(http, new List<GoogleWorkspaceOperationReceipt>()))) {
                await Assert.ThrowsAsync<InvalidOperationException>(() => refusing.DeleteFileAsync("file-1"));
            }
            Assert.Equal(0, requests);

            using var accepting = new GoogleDriveClient(Session(http,
                new List<GoogleWorkspaceOperationReceipt>(), GoogleWorkspaceDataLossDecision.AcceptSpecifiedLoss,
                "delete file-1"));
            await accepting.DeleteFileAsync("file-1");
            Assert.Equal(1, requests);
        }

        [Fact]
        public async Task ResumableUploadCheckpointSurvivesRestartAndReconcilesServerOffset() {
            const int total = 300 * 1024;
            bool firstChunkCommitted = false;
            using var http = new HttpClient(new Handler(async request => {
                if (request.Method == HttpMethod.Post) { var response = Json("{}"); response.Headers.Location = new Uri("https://upload.googleapis.com/durable"); return response; }
                string range = request.Content!.Headers.GetValues("Content-Range").Single();
                if (range == $"bytes */{total}") {
                    var status = new HttpResponseMessage((HttpStatusCode)308) { Content = new StringContent(string.Empty) };
                    if (firstChunkCommitted) status.Headers.TryAddWithoutValidation("Range", "bytes=0-262143");
                    return status;
                }
                _ = await request.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                if (range.StartsWith("bytes 0-", StringComparison.Ordinal)) {
                    firstChunkCommitted = true; var accepted = new HttpResponseMessage((HttpStatusCode)308) { Content = new StringContent(string.Empty) };
                    accepted.Headers.TryAddWithoutValidation("Range", "bytes=0-262143"); return accepted;
                }
                return new HttpResponseMessage(HttpStatusCode.Created) { Content = new StringContent("{\"id\":\"durable-1\",\"version\":\"2\"}", Encoding.UTF8, "application/json") };
            }));
            byte[] payload = new byte[total]; GoogleDriveResumableUploadCheckpoint? saved = null;
            var receipts = new List<GoogleWorkspaceOperationReceipt>();
            using (var first = new GoogleDriveClient(Session(http, receipts))) {
                await Assert.ThrowsAsync<StopAfterCheckpointException>(() => first.UploadResumableStreamAsync(
                    new MemoryStream(payload, false), payload.Length,
                    new GoogleDriveUploadOptions { Name = "durable.bin", ResumableChunkSize = 256 * 1024 },
                    checkpointSink: (checkpoint, _) => {
                        saved = checkpoint;
                        if (checkpoint.ConfirmedBytes > 0) throw new StopAfterCheckpointException();
                        return Task.CompletedTask;
                    }));
            }
            GoogleDriveResumableUploadCheckpoint restored = GoogleDriveResumableUploadCheckpoint.Parse(Assert.IsType<GoogleDriveResumableUploadCheckpoint>(saved).Value);
            using var second = new GoogleDriveClient(Session(http, receipts));
            GoogleDriveResumableUploadResult result = await second.UploadResumableStreamAsync(
                new MemoryStream(payload, false), payload.Length,
                new GoogleDriveUploadOptions { Name = "durable.bin", ResumableChunkSize = 256 * 1024 }, restored);
            Assert.Equal("durable-1", result.File.Id);
            Assert.Equal(total, result.Checkpoint.ConfirmedBytes);
            Assert.DoesNotContain("upload.googleapis.com", result.Checkpoint.ToString(), StringComparison.Ordinal);
            Assert.DoesNotContain(receipts, receipt => receipt.Target.Contains("upload.googleapis.com", StringComparison.Ordinal));
        }

        [Fact]
        public async Task ResumableUploadRejectsSourceMutationBeforeReportingSuccess() {
            const int total = 16;
            using var http = new HttpClient(new Handler(request => {
                if (request.Method == HttpMethod.Post) {
                    var initiated = Json("{}");
                    initiated.Headers.Location = new Uri("https://upload.googleapis.com/mutable");
                    return initiated;
                }
                string range = request.Content!.Headers.GetValues("Content-Range").Single();
                return range == $"bytes */{total}"
                    ? new HttpResponseMessage((HttpStatusCode)308) { Content = new StringContent(string.Empty) }
                    : new HttpResponseMessage(HttpStatusCode.Created) { Content = new StringContent("{\"id\":\"mutable-1\",\"version\":\"2\"}", Encoding.UTF8, "application/json") };
            }));
            byte[] payload = new byte[total];
            using var client = new GoogleDriveClient(Session(http, new List<GoogleWorkspaceOperationReceipt>()));

            await Assert.ThrowsAsync<InvalidOperationException>(() => client.UploadResumableStreamAsync(
                new MemoryStream(payload), payload.Length, new GoogleDriveUploadOptions { Name = "mutable.bin" },
                checkpointSink: (checkpoint, _) => {
                    if (checkpoint.ConfirmedBytes == 0) payload[0] = 1;
                    return Task.CompletedTask;
                }));
        }

        [Fact]
        public async Task RangedDownloadCheckpointSurvivesRestartAndGuardsRevision() {
            byte[] payload = Enumerable.Range(0, 700_000).Select(value => unchecked((byte)value)).ToArray();
            using var http = new HttpClient(new Handler(request => {
                if (!request.RequestUri!.Query.Contains("alt=media", StringComparison.Ordinal)) return Json($"{{\"id\":\"file-1\",\"version\":\"7\",\"size\":\"{payload.Length}\"}}");
                RangeItemHeaderValue range = request.Headers.Range!.Ranges.Single(); int start = checked((int)range.From!.Value); int end = checked((int)range.To!.Value);
                return new HttpResponseMessage(HttpStatusCode.PartialContent) { Content = new ByteArrayContent(payload.Skip(start).Take(end - start + 1).ToArray()) };
            }));
            string path = Path.Combine(Path.GetTempPath(), "OfficeIMO-download-" + Guid.NewGuid().ToString("N") + ".bin");
            GoogleDriveDownloadCheckpoint? saved = null;
            try {
                using (var first = new GoogleDriveClient(Session(http, new List<GoogleWorkspaceOperationReceipt>()))) {
                    await Assert.ThrowsAsync<StopAfterCheckpointException>(() => first.DownloadToFileAsync("file-1", path,
                        checkpointSink: (checkpoint, _) => {
                            if (checkpoint.ConfirmedBytes > 0) throw new StopAfterCheckpointException();
                            saved = checkpoint;
                            return Task.CompletedTask;
                        },
                        chunkSize: 256 * 1024));
                }
                Assert.Equal(256 * 1024, new FileInfo(path).Length);
                GoogleDriveDownloadCheckpoint restored = GoogleDriveDownloadCheckpoint.Parse(Assert.IsType<GoogleDriveDownloadCheckpoint>(saved).Value);
                using var second = new GoogleDriveClient(Session(http, new List<GoogleWorkspaceOperationReceipt>()));
                GoogleDriveDownloadCheckpoint completed = await second.DownloadToFileAsync("file-1", path, restored, chunkSize: 256 * 1024);
                Assert.Equal(payload.Length, completed.ConfirmedBytes);
                Assert.Equal(payload, File.ReadAllBytes(path));
            } finally { if (File.Exists(path)) File.Delete(path); }
        }

        private static GoogleWorkspaceSession Session(HttpClient http, IList<GoogleWorkspaceOperationReceipt> receipts,
            GoogleWorkspaceDataLossDecision dataLossDecision = GoogleWorkspaceDataLossDecision.RejectPotentialLoss,
            string? acceptedLoss = null) {
            var options = new GoogleWorkspaceSessionOptions { HttpClient = http, ExpectedAccount = "test@example.com", OperationReceiptSink = receipts.Add };
            options.OperationPolicyProvider = context => new GoogleWorkspaceOperationPolicy(options.ExpectedAccount!,
                new[] { GoogleWorkspaceScopeCatalog.DriveFile }, context.Target, "explicit-test-revision",
                options.MaxRetryCount, options.MaxRetryElapsedTime, options.RateLimitPolicy,
                dataLossDecision, acceptedLoss);
            return new GoogleWorkspaceSession(new StaticAccessTokenCredentialSource("token"), options);
        }
        private static HttpResponseMessage Json(string value) => new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(value, Encoding.UTF8, "application/json") };
        private sealed class StopAfterCheckpointException : Exception { }
        private sealed class Handler : HttpMessageHandler {
            private readonly Func<HttpRequestMessage, Task<HttpResponseMessage>> _handler;
            internal Handler(Func<HttpRequestMessage, HttpResponseMessage> handler) { _handler = request => Task.FromResult(handler(request)); }
            internal Handler(Func<HttpRequestMessage, Task<HttpResponseMessage>> handler) { _handler = handler; }
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => _handler(request);
        }
    }
}
