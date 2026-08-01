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
            Assert.Equal(GoogleWorkspaceOperationPolicy.ResourceAbsentForCreateRevision, receipt.Policy.ExpectedRevision);
            Assert.Equal(GoogleWorkspaceDataLossDecision.RejectPotentialLoss, receipt.Policy.DataLossDecision);
            Assert.Equal(GoogleWorkspaceRevisionPreconditionKind.ResourceAbsentCreate, receipt.RevisionPreconditionKind);
            Assert.Null(receipt.EnforcedRevision);
        }

        [Fact]
        public async Task MutationAppliesExpectedEntityTagAsIfMatch() {
            string? ifMatch = null;
            using var http = new HttpClient(new Handler(request => {
                ifMatch = request.Headers.IfMatch.Single().Tag;
                return Json("{}");
            }));
            var receipts = new List<GoogleWorkspaceOperationReceipt>();
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                ExpectedAccount = "test@example.com",
                OperationReceiptSink = receipts.Add,
            };
            options.OperationPolicyProvider = context => new GoogleWorkspaceOperationPolicy(
                options.ExpectedAccount!, new[] { GoogleWorkspaceScopeCatalog.DriveFile }, context.Target,
                "\"revision-7\"", options.MaxRetryCount, options.MaxRetryElapsedTime,
                options.RateLimitPolicy, GoogleWorkspaceDataLossDecision.RejectPotentialLoss);
            using var transport = new GoogleWorkspaceHttpTransport(options);

            await transport.SendJsonAsync<object>("token", new HttpMethod("PATCH"),
                "https://www.googleapis.com/drive/v3/files/file-1", new { name = "updated" },
                GoogleWorkspaceRequestSafety.Idempotent, "Google Drive API", new TranslationReport());

            Assert.Equal("\"revision-7\"", ifMatch);
            GoogleWorkspaceOperationReceipt receipt = Assert.Single(receipts);
            Assert.True(receipt.Succeeded);
            Assert.Equal(GoogleWorkspaceRevisionPreconditionKind.HttpEntityTag, receipt.RevisionPreconditionKind);
            Assert.Equal("\"revision-7\"", receipt.EnforcedRevision);
        }

        [Theory]
        [InlineData("drive-version:7")]
        [InlineData("*")]
        [InlineData("W/\"revision-7\"")]
        public async Task MutationRejectsObservedRevisionThatCannotBeEnforced(string expectedRevision) {
            int requests = 0;
            using var http = new HttpClient(new Handler(_ => { requests++; return Json("{}"); }));
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                ExpectedAccount = "test@example.com",
                OperationReceiptSink = _ => { },
            };
            options.OperationPolicyProvider = context => new GoogleWorkspaceOperationPolicy(
                options.ExpectedAccount!, new[] { GoogleWorkspaceScopeCatalog.DriveFile }, context.Target,
                expectedRevision, options.MaxRetryCount, options.MaxRetryElapsedTime,
                options.RateLimitPolicy, GoogleWorkspaceDataLossDecision.RejectPotentialLoss);
            using var transport = new GoogleWorkspaceHttpTransport(options);

            await Assert.ThrowsAsync<InvalidOperationException>(() => transport.SendJsonAsync<object>(
                "token", new HttpMethod("PATCH"), "https://www.googleapis.com/drive/v3/files/file-1",
                new { name = "updated" }, GoogleWorkspaceRequestSafety.Idempotent,
                "Google Drive API", new TranslationReport()));

            Assert.Equal(0, requests);
        }

        [Fact]
        public async Task ResourceAbsentRevisionCannotAuthorizePostUpdate() {
            int requests = 0;
            using var http = new HttpClient(new Handler(_ => { requests++; return Json("{}"); }));
            var options = MutationOptions(http, _ => { });
            using var transport = new GoogleWorkspaceHttpTransport(options);

            await Assert.ThrowsAsync<InvalidOperationException>(() => transport.SendJsonAsync<object>(
                "token", HttpMethod.Post, "https://docs.googleapis.com/v1/documents/doc-1:batchUpdate",
                new { requests = Array.Empty<object>() }, GoogleWorkspaceRequestSafety.NonIdempotent,
                "Google Docs API", new TranslationReport(),
                mutationKind: GoogleWorkspaceMutationKind.Update,
                revisionPrecondition: GoogleWorkspaceRevisionPrecondition.PayloadRevision("docs-revision-7")));

            Assert.Equal(0, requests);
        }

        [Fact]
        public async Task PayloadRevisionMustMatchPolicyAndIsRecordedAsEnforced() {
            bool hadIfMatch = true;
            using var http = new HttpClient(new Handler(request => {
                hadIfMatch = request.Headers.IfMatch.Any();
                return Json("{}");
            }));
            var receipts = new List<GoogleWorkspaceOperationReceipt>();
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                ExpectedAccount = "test@example.com",
                OperationReceiptSink = receipts.Add,
            };
            options.OperationPolicyProvider = context => new GoogleWorkspaceOperationPolicy(
                options.ExpectedAccount!, new[] { GoogleWorkspaceScopeCatalog.Documents }, context.Target,
                context.AdapterExpectedRevision!, options.MaxRetryCount, options.MaxRetryElapsedTime,
                options.RateLimitPolicy, GoogleWorkspaceDataLossDecision.RejectPotentialLoss);
            using var transport = new GoogleWorkspaceHttpTransport(options);

            await transport.SendJsonAsync<object>("token", HttpMethod.Post,
                "https://docs.googleapis.com/v1/documents/doc-1:batchUpdate",
                new { requests = Array.Empty<object>() }, GoogleWorkspaceRequestSafety.NonIdempotent,
                "Google Docs API", new TranslationReport(), mutationKind: GoogleWorkspaceMutationKind.Update,
                revisionPrecondition: GoogleWorkspaceRevisionPrecondition.PayloadRevision("docs-revision-7"));

            Assert.False(hadIfMatch);
            GoogleWorkspaceOperationReceipt receipt = Assert.Single(receipts);
            Assert.Equal(GoogleWorkspaceRevisionPreconditionKind.PayloadRevision, receipt.RevisionPreconditionKind);
            Assert.Equal("docs-revision-7", receipt.EnforcedRevision);
            Assert.Equal(receipt.Policy.ExpectedRevision, receipt.EnforcedRevision);
        }

        [Fact]
        public async Task ReceiptSinkFailureKeepsRemoteSuccessDistinguishable() {
            int requests = 0;
            using var http = new HttpClient(new Handler(_ => { requests++; return Json("{}"); }));
            var options = MutationOptions(http, _ => throw new IOException("receipt store unavailable"));
            using var transport = new GoogleWorkspaceHttpTransport(options);

            GoogleWorkspaceReceiptPersistenceException exception =
                await Assert.ThrowsAsync<GoogleWorkspaceReceiptPersistenceException>(() =>
                    transport.SendJsonAsync<object>("token", HttpMethod.Post,
                        "https://www.googleapis.com/drive/v3/files", new { name = "created" },
                        GoogleWorkspaceRequestSafety.NonIdempotent, "Google Drive API", new TranslationReport(),
                        mutationKind: GoogleWorkspaceMutationKind.Create));

            Assert.Equal(1, requests);
            Assert.True(exception.RemoteOperationSucceeded);
            Assert.True(exception.Receipt.Succeeded);
        }

        [Fact]
        public async Task ReceiptSinkFailureDoesNotReplaceRemoteFailure() {
            int requests = 0;
            using var http = new HttpClient(new Handler(_ => {
                requests++;
                return new HttpResponseMessage(HttpStatusCode.BadRequest) {
                    Content = new StringContent("{\"error\":{\"message\":\"bad request\"}}", Encoding.UTF8, "application/json")
                };
            }));
            var options = MutationOptions(http, _ => throw new IOException("receipt store unavailable"));
            using var transport = new GoogleWorkspaceHttpTransport(options);

            GoogleWorkspaceApiException exception = await Assert.ThrowsAsync<GoogleWorkspaceApiException>(() =>
                transport.SendJsonAsync<object>("token", HttpMethod.Post,
                    "https://www.googleapis.com/drive/v3/files", new { name = "created" },
                    GoogleWorkspaceRequestSafety.NonIdempotent, "Google Drive API", new TranslationReport(),
                    mutationKind: GoogleWorkspaceMutationKind.Create));

            Assert.Equal(1, requests);
            var receiptFailure = Assert.IsType<GoogleWorkspaceReceiptPersistenceException>(
                exception.Data[GoogleWorkspaceReceiptPersistenceException.ExceptionDataKey]);
            Assert.False(receiptFailure.RemoteOperationSucceeded);
            Assert.False(receiptFailure.Receipt.Succeeded);
        }

        [Fact]
        public async Task Google403RateLimitReasonUsesConfiguredRetryPolicy() {
            int attempts = 0;
            using var http = new HttpClient(new Handler(_ => {
                attempts++;
                return attempts == 1
                    ? new HttpResponseMessage(HttpStatusCode.Forbidden) {
                        Content = new StringContent("{\"error\":{\"errors\":[{\"reason\":\"userRateLimitExceeded\"}]}}", Encoding.UTF8, "application/json")
                    }
                    : Json("{\"id\":\"file-1\"}");
            }));
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                MaxRetryCount = 1,
                RetryBaseDelay = TimeSpan.FromMilliseconds(1),
                RetryMaxDelay = TimeSpan.FromMilliseconds(1),
                RateLimitPolicy = GoogleWorkspaceRateLimitPolicy.HonorRetryAfter,
            };
            using var transport = new GoogleWorkspaceHttpTransport(options);

            await transport.SendJsonAsync<object>("token", HttpMethod.Get,
                "https://www.googleapis.com/drive/v3/files/file-1", null,
                GoogleWorkspaceRequestSafety.Safe, "Google Drive API", new TranslationReport());

            Assert.Equal(2, attempts);
        }

        [Theory]
        [InlineData("insufficientFilePermissions", GoogleWorkspaceRateLimitPolicy.HonorRetryAfter)]
        [InlineData("userRateLimitExceeded", GoogleWorkspaceRateLimitPolicy.FailFast)]
        public async Task Google403DoesNotRetryAuthorizationErrorsOrFailFastQuota(string reason,
            GoogleWorkspaceRateLimitPolicy rateLimitPolicy) {
            int attempts = 0;
            using var http = new HttpClient(new Handler(_ => {
                attempts++;
                return new HttpResponseMessage(HttpStatusCode.Forbidden) {
                    Content = new StringContent("{\"error\":{\"errors\":[{\"reason\":\"" + reason + "\"}]}}",
                        Encoding.UTF8, "application/json")
                };
            }));
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                MaxRetryCount = 1,
                RetryBaseDelay = TimeSpan.FromMilliseconds(1),
                RetryMaxDelay = TimeSpan.FromMilliseconds(1),
                RateLimitPolicy = rateLimitPolicy,
            };
            using var transport = new GoogleWorkspaceHttpTransport(options);

            await Assert.ThrowsAsync<GoogleWorkspaceApiException>(() => transport.SendJsonAsync<object>(
                "token", HttpMethod.Get, "https://www.googleapis.com/drive/v3/files/file-1", null,
                GoogleWorkspaceRequestSafety.Safe, "Google Drive API", new TranslationReport()));

            Assert.Equal(1, attempts);
        }

        [Fact]
        public async Task RetryResponseIsDisposedBeforeElapsedBudgetRejection() {
            var response = new TrackingResponseMessage(HttpStatusCode.ServiceUnavailable) {
                Content = new StringContent("retry later")
            };
            response.Headers.RetryAfter = new RetryConditionHeaderValue(TimeSpan.FromSeconds(5));
            using var http = new HttpClient(new Handler(_ => response));
            var retryOptions = new GoogleWorkspaceRetryOptions(1, TimeSpan.FromSeconds(1),
                TimeSpan.FromSeconds(5), new GoogleWorkspaceSessionOptions {
                    MaxRetryCount = 1,
                    MaxRetryElapsedTime = TimeSpan.FromMilliseconds(10),
                    RetryBaseDelay = TimeSpan.FromSeconds(1),
                    RetryMaxDelay = TimeSpan.FromSeconds(5),
                });

            await Assert.ThrowsAsync<TimeoutException>(() => GoogleWorkspaceRetryPolicy.SendAsync(
                http, () => new HttpRequestMessage(HttpMethod.Get, "https://www.googleapis.com/drive/v3/files/file-1"),
                retryOptions, GoogleWorkspaceRequestSafety.Safe, CancellationToken.None));

            Assert.True(response.WasDisposed);
        }

        [Fact]
        public async Task Google403BodyReadFailureDisposesAndRetriesSafeRequest() {
            int attempts = 0;
            var firstResponse = new TrackingResponseMessage(HttpStatusCode.Forbidden) {
                Content = new ThrowingReadContent()
            };
            using var http = new HttpClient(new Handler(_ => {
                attempts++;
                return attempts == 1 ? firstResponse : Json("{\"id\":\"file-1\"}");
            }));
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                MaxRetryCount = 1,
                RetryBaseDelay = TimeSpan.FromMilliseconds(1),
                RetryMaxDelay = TimeSpan.FromMilliseconds(1),
            };
            using var transport = new GoogleWorkspaceHttpTransport(options);

            await transport.SendJsonAsync<object>("token", HttpMethod.Get,
                "https://www.googleapis.com/drive/v3/files/file-1", null,
                GoogleWorkspaceRequestSafety.Safe, "Google Drive API", new TranslationReport());

            Assert.Equal(2, attempts);
            Assert.True(firstResponse.WasDisposed);
        }

        [Fact]
        public async Task FinalResponseIsDisposedWhenElapsedBudgetExpiresBeforeReturn() {
            var response = new TrackingResponseMessage(HttpStatusCode.OK) {
                Content = new StringContent("{}", Encoding.UTF8, "application/json")
            };
            using var http = new HttpClient(new Handler(async _ => {
                await Task.Delay(TimeSpan.FromMilliseconds(30)).ConfigureAwait(false);
                return response;
            }));
            var retryOptions = new GoogleWorkspaceRetryOptions(0, TimeSpan.FromMilliseconds(1),
                TimeSpan.FromMilliseconds(1), new GoogleWorkspaceSessionOptions {
                    MaxRetryCount = 0,
                    MaxRetryElapsedTime = TimeSpan.FromMilliseconds(5),
                });

            await Assert.ThrowsAsync<TimeoutException>(() => GoogleWorkspaceRetryPolicy.SendAsync(
                http, () => new HttpRequestMessage(HttpMethod.Get, "https://www.googleapis.com/drive/v3/files/file-1"),
                retryOptions, GoogleWorkspaceRequestSafety.Safe, CancellationToken.None));

            Assert.True(response.WasDisposed);
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
            Assert.DoesNotContain(receipts, receipt =>
                receipt.Succeeded && receipt.MutationKind == GoogleWorkspaceMutationKind.Create);
            GoogleDriveResumableUploadCheckpoint restored = GoogleDriveResumableUploadCheckpoint.Parse(Assert.IsType<GoogleDriveResumableUploadCheckpoint>(saved).Value);
            using var second = new GoogleDriveClient(Session(http, receipts));
            GoogleDriveResumableUploadResult result = await second.UploadResumableStreamAsync(
                new MemoryStream(payload, false), payload.Length,
                new GoogleDriveUploadOptions { Name = "durable.bin", ResumableChunkSize = 256 * 1024 }, restored);
            Assert.Equal("durable-1", result.File.Id);
            Assert.Equal(total, result.Checkpoint.ConfirmedBytes);
            Assert.DoesNotContain("upload.googleapis.com", result.Checkpoint.ToString(), StringComparison.Ordinal);
            Assert.DoesNotContain(receipts, receipt => receipt.Target.Contains("upload.googleapis.com", StringComparison.Ordinal));
            GoogleWorkspaceOperationReceipt createReceipt = Assert.Single(receipts, receipt =>
                receipt.Succeeded && receipt.MutationKind == GoogleWorkspaceMutationKind.Create);
            Assert.Equal(GoogleWorkspaceRevisionPreconditionKind.ResumableSessionState,
                createReceipt.RevisionPreconditionKind);
            Assert.StartsWith("resumable-session:", createReceipt.EnforcedRevision, StringComparison.Ordinal);
            Assert.Contains(receipts, receipt => receipt.Succeeded &&
                receipt.MutationKind == GoogleWorkspaceMutationKind.Action &&
                receipt.EnforcedRevision != null &&
                receipt.EnforcedRevision.StartsWith("content-range:bytes ", StringComparison.Ordinal));
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
                            saved = checkpoint;
                            if (checkpoint.ConfirmedBytes > 0) throw new StopAfterCheckpointException();
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

        [Fact]
        public async Task GuardedDownloadNeverTruncatesAnExistingDestination() {
            using var http = new HttpClient(new Handler(_ =>
                Json("{\"id\":\"file-1\",\"version\":\"7\",\"size\":\"3\"}")));
            string path = Path.Combine(Path.GetTempPath(),
                "OfficeIMO-download-existing-" + Guid.NewGuid().ToString("N") + ".bin");
            byte[] original = Encoding.UTF8.GetBytes("do not replace");
            File.WriteAllBytes(path, original);
            try {
                using var client = new GoogleDriveClient(Session(http, new List<GoogleWorkspaceOperationReceipt>()));

                await Assert.ThrowsAsync<IOException>(() => client.DownloadToFileAsync("file-1", path));

                Assert.Equal(original, File.ReadAllBytes(path));
            } finally { if (File.Exists(path)) File.Delete(path); }
        }

        [Fact]
        public async Task ZeroByteDownloadCheckpointRejectsAReplacementDestination() {
            using var http = new HttpClient(new Handler(_ =>
                Json("{\"id\":\"file-1\",\"version\":\"7\",\"size\":\"3\"}")));
            string path = Path.Combine(Path.GetTempPath(),
                "OfficeIMO-download-replaced-" + Guid.NewGuid().ToString("N") + ".bin");
            GoogleDriveDownloadCheckpoint? saved = null;
            byte[] replacement = Encoding.UTF8.GetBytes("unrelated content");
            try {
                using (var first = new GoogleDriveClient(Session(http,
                           new List<GoogleWorkspaceOperationReceipt>()))) {
                    await Assert.ThrowsAsync<StopAfterCheckpointException>(() =>
                        first.DownloadToFileAsync("file-1", path,
                            checkpointSink: (checkpoint, _) => {
                                saved = checkpoint;
                                throw new StopAfterCheckpointException();
                            },
                            chunkSize: 256 * 1024));
                }

                File.Delete(path);
                File.WriteAllBytes(path, replacement);
                GoogleDriveDownloadCheckpoint restored = GoogleDriveDownloadCheckpoint.Parse(
                    Assert.IsType<GoogleDriveDownloadCheckpoint>(saved).Value);
                using var second = new GoogleDriveClient(Session(http,
                    new List<GoogleWorkspaceOperationReceipt>()));

                await Assert.ThrowsAsync<InvalidOperationException>(() =>
                    second.DownloadToFileAsync("file-1", path, restored,
                        chunkSize: 256 * 1024));

                Assert.Equal(replacement, File.ReadAllBytes(path));
            } finally { if (File.Exists(path)) File.Delete(path); }
        }

        private static GoogleWorkspaceSession Session(HttpClient http, IList<GoogleWorkspaceOperationReceipt> receipts,
            GoogleWorkspaceDataLossDecision dataLossDecision = GoogleWorkspaceDataLossDecision.RejectPotentialLoss,
            string? acceptedLoss = null) {
            var options = new GoogleWorkspaceSessionOptions { HttpClient = http, ExpectedAccount = "test@example.com", OperationReceiptSink = receipts.Add };
            options.OperationPolicyProvider = context => new GoogleWorkspaceOperationPolicy(options.ExpectedAccount!,
                new[] { GoogleWorkspaceScopeCatalog.DriveFile }, context.Target,
                context.RevisionPreconditionKind switch {
                    GoogleWorkspaceRevisionPreconditionKind.ResourceAbsentCreate => GoogleWorkspaceOperationPolicy.ResourceAbsentForCreateRevision,
                    GoogleWorkspaceRevisionPreconditionKind.PayloadRevision => context.AdapterExpectedRevision!,
                    GoogleWorkspaceRevisionPreconditionKind.ResumableSessionState => context.AdapterExpectedRevision!,
                    GoogleWorkspaceRevisionPreconditionKind.Unavailable => GoogleWorkspaceOperationPolicy.ExplicitlyUnversionedRevision("test mutation"),
                    _ => "\"test-etag\"",
                },
                options.MaxRetryCount, options.MaxRetryElapsedTime, options.RateLimitPolicy,
                dataLossDecision, acceptedLoss);
            return new GoogleWorkspaceSession(new StaticAccessTokenCredentialSource("token"), options);
        }
        private static GoogleWorkspaceSessionOptions MutationOptions(HttpClient http,
            Action<GoogleWorkspaceOperationReceipt> receiptSink) {
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = http,
                ExpectedAccount = "test@example.com",
                OperationReceiptSink = receiptSink,
            };
            options.OperationPolicyProvider = context => new GoogleWorkspaceOperationPolicy(
                options.ExpectedAccount!, new[] { GoogleWorkspaceScopeCatalog.DriveFile }, context.Target,
                GoogleWorkspaceOperationPolicy.ResourceAbsentForCreateRevision,
                options.MaxRetryCount, options.MaxRetryElapsedTime, options.RateLimitPolicy,
                GoogleWorkspaceDataLossDecision.RejectPotentialLoss);
            return options;
        }
        private static HttpResponseMessage Json(string value) => new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(value, Encoding.UTF8, "application/json") };
        private sealed class StopAfterCheckpointException : Exception { }
        private sealed class TrackingResponseMessage : HttpResponseMessage {
            internal TrackingResponseMessage(HttpStatusCode statusCode) : base(statusCode) { }
            internal bool WasDisposed { get; private set; }
            protected override void Dispose(bool disposing) {
                WasDisposed = true;
                base.Dispose(disposing);
            }
        }
        private sealed class ThrowingReadContent : HttpContent {
            protected override Task SerializeToStreamAsync(Stream stream, TransportContext? context) =>
                Task.FromException(new IOException("Transient response read failure."));
            protected override bool TryComputeLength(out long length) { length = 0; return false; }
            protected override Task<Stream> CreateContentReadStreamAsync() =>
                Task.FromResult<Stream>(new ThrowingReadStream());
        }
        private sealed class ThrowingReadStream : MemoryStream {
            public override Task<int> ReadAsync(byte[] buffer, int offset, int count,
                CancellationToken cancellationToken) =>
                Task.FromException<int>(new IOException("Transient response read failure."));
        }
        private sealed class Handler : HttpMessageHandler {
            private readonly Func<HttpRequestMessage, Task<HttpResponseMessage>> _handler;
            internal Handler(Func<HttpRequestMessage, HttpResponseMessage> handler) { _handler = request => Task.FromResult(handler(request)); }
            internal Handler(Func<HttpRequestMessage, Task<HttpResponseMessage>> handler) { _handler = handler; }
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => _handler(request);
        }
    }
}
