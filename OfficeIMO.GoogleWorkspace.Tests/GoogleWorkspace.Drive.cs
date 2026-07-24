using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public class GoogleWorkspaceDriveTests {
        [Fact]
        public void Test_DriveClientOptions_FileAuthoringUsesDriveFileForReadsAndWrites() {
            GoogleDriveClientOptions options = GoogleDriveClientOptions.ForFileAuthoring();

            Assert.Equal(GoogleWorkspaceScopeCatalog.DriveFile, Assert.Single(options.ReadScopes));
            Assert.Equal(GoogleWorkspaceScopeCatalog.DriveFile, Assert.Single(options.WriteScopes));
        }

        [Fact]
        public async Task Test_DriveClient_DiscoversImportAndExportFormats() {
            Uri? requestedUri = null;
            using var httpClient = new HttpClient(new FakeHandler(request => {
                requestedUri = request.RequestUri;
                return Task.FromResult(Json("{\"importFormats\":{\"application/vnd.openxmlformats-officedocument.wordprocessingml.document\":[\"application/vnd.google-apps.document\"]},\"exportFormats\":{\"application/vnd.google-apps.document\":[\"application/vnd.openxmlformats-officedocument.wordprocessingml.document\"]}}"));
            }));
            var credential = new RecordingCredentialSource();
            using var client = CreateClient(httpClient, credential);

            var formats = await client.GetFormatsAsync();

            Assert.Contains(GoogleDriveMimeTypes.Document, formats.ImportFormats[GoogleDriveMimeTypes.MicrosoftWord]);
            Assert.Contains(GoogleDriveMimeTypes.MicrosoftWord, formats.ExportFormats[GoogleDriveMimeTypes.Document]);
            Assert.Equal("https://www.googleapis.com/drive/v3/about?fields=importFormats,exportFormats", requestedUri!.AbsoluteUri);
            Assert.Equal(GoogleWorkspaceScopeCatalog.DriveReadonly, Assert.Single(credential.LastScopes));
        }

        [Fact]
        public async Task Test_DriveClient_DownloadAndExportEnforceConfiguredResponseLimit() {
            using var httpClient = new HttpClient(new FakeHandler(_ => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                Content = new ByteArrayContent(new byte[5]),
            })));
            using var client = new GoogleDriveClient(
                new GoogleWorkspaceSession(
                    new RecordingCredentialSource(),
                    new GoogleWorkspaceSessionOptions { HttpClient = httpClient, MaxRetryCount = 1 }),
                new GoogleDriveClientOptions { MaxDownloadBytes = 4 });

            await Assert.ThrowsAsync<InvalidDataException>(() => client.DownloadAsync("file-1"));
            await Assert.ThrowsAsync<InvalidDataException>(() => client.ExportAsync("file-1", GoogleDriveMimeTypes.MicrosoftWord));
        }

        [Fact]
        public async Task Test_DriveClient_TemporaryPublicLease_AlwaysDeletesOwnedFile() {
            var methods = new List<string>();
            using var httpClient = new HttpClient(new FakeHandler(request => {
                methods.Add(request.Method.Method + " " + request.RequestUri!.AbsoluteUri);
                if (request.Method == HttpMethod.Post && request.RequestUri.AbsoluteUri.Contains("uploadType=multipart", StringComparison.Ordinal)) {
                    return Task.FromResult(Json("{\"id\":\"temporary-1\",\"name\":\"image.png\",\"mimeType\":\"image/png\"}"));
                }

                if (request.Method == HttpMethod.Post && request.RequestUri.AbsoluteUri.Contains("/temporary-1/permissions?", StringComparison.Ordinal)) {
                    return Task.FromResult(Json("{\"id\":\"permission-1\",\"type\":\"anyone\",\"role\":\"reader\"}"));
                }

                if (request.Method == HttpMethod.Delete && request.RequestUri.AbsoluteUri.Contains("/files/temporary-1?", StringComparison.Ordinal)) {
                    return Task.FromResult(Json("{}"));
                }

                return Task.FromResult(NotFound());
            }));
            using var client = CreateClient(httpClient);
            var report = new TranslationReport();

            var lease = await GoogleDriveTemporaryContentLease.CreatePublicReadLeaseAsync(
                client,
                Encoding.UTF8.GetBytes("image"),
                new GoogleDriveUploadOptions { Name = "image.png", ContentType = "image/png" },
                report);
            var cleanup = await lease.CleanupAsync();

            Assert.Equal("https://drive.google.com/uc?export=download&id=temporary-1", lease.PublicUri);
            Assert.Equal(GoogleDriveCleanupStatus.Deleted, Assert.Single(cleanup.Entries).Status);
            Assert.Contains(methods, request => request.StartsWith("DELETE ", StringComparison.Ordinal));
            Assert.Contains(report.Notices, notice => notice.Code == "DRIVE.TEMPORARY_CONTENT.CLEANED");
        }

        [Fact]
        public async Task Test_DriveClient_TemporaryPublicLease_RetriesCleanupAfterCancellation() {
            int deleteAttempts = 0;
            using var httpClient = new HttpClient(new FakeHandler((request, cancellationToken) => {
                if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri.Contains("uploadType=multipart", StringComparison.Ordinal)) {
                    return Task.FromResult(Json("{\"id\":\"temporary-retry\",\"name\":\"image.png\",\"mimeType\":\"image/png\"}"));
                }

                if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri.Contains("/temporary-retry/permissions?", StringComparison.Ordinal)) {
                    return Task.FromResult(Json("{\"id\":\"permission-retry\",\"type\":\"anyone\",\"role\":\"reader\"}"));
                }

                if (request.Method == HttpMethod.Delete && request.RequestUri!.AbsoluteUri.Contains("/files/temporary-retry?", StringComparison.Ordinal)) {
                    deleteAttempts++;
                    if (cancellationToken.IsCancellationRequested) {
                        return Task.FromCanceled<HttpResponseMessage>(cancellationToken);
                    }

                    return Task.FromResult(Json("{}"));
                }

                return Task.FromResult(NotFound());
            }));
            using var client = CreateClient(httpClient);
            GoogleDriveTemporaryContentLease lease = await GoogleDriveTemporaryContentLease.CreatePublicReadLeaseAsync(
                client,
                Encoding.UTF8.GetBytes("image"),
                new GoogleDriveUploadOptions { Name = "image.png", ContentType = "image/png" });

            using (var cancellation = new CancellationTokenSource()) {
                cancellation.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => lease.CleanupAsync(cancellation.Token));
            }
            int attemptsAfterCancellation = deleteAttempts;

            GoogleDriveCleanupReport cleanup = await lease.CleanupAsync();

            Assert.Equal(attemptsAfterCancellation + 1, deleteAttempts);
            Assert.Equal(GoogleDriveCleanupStatus.Deleted, Assert.Single(cleanup.Entries).Status);
        }

        [Fact]
        public async Task Test_DriveClient_TemporaryPublicLease_UsesResumableUploadAboveMultipartLimit() {
            var requests = new List<string>();
            using var httpClient = new HttpClient(new FakeHandler(request => {
                requests.Add(request.Method.Method + " " + request.RequestUri!.AbsoluteUri);
                if (request.Method == HttpMethod.Post && request.RequestUri.AbsoluteUri.Contains("uploadType=resumable", StringComparison.Ordinal)) {
                    HttpResponseMessage response = Json("{}");
                    response.Headers.Location = new Uri("https://upload.example.test/temporary-large");
                    return Task.FromResult(response);
                }
                if (request.Method == HttpMethod.Put && request.RequestUri.AbsoluteUri == "https://upload.example.test/temporary-large") {
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.Created) {
                        Content = new StringContent("{\"id\":\"temporary-large\",\"name\":\"large.png\",\"mimeType\":\"image/png\"}", Encoding.UTF8, "application/json")
                    });
                }
                if (request.Method == HttpMethod.Post && request.RequestUri.AbsoluteUri.Contains("/temporary-large/permissions?", StringComparison.Ordinal)) {
                    return Task.FromResult(Json("{\"id\":\"permission-large\",\"type\":\"anyone\",\"role\":\"reader\"}"));
                }
                if (request.Method == HttpMethod.Delete && request.RequestUri.AbsoluteUri.Contains("/files/temporary-large?", StringComparison.Ordinal)) {
                    return Task.FromResult(Json("{}"));
                }
                return Task.FromResult(NotFound());
            }));
            using var client = CreateClient(httpClient);
            byte[] content = new byte[(5 * 1024 * 1024) + 1];

            GoogleDriveTemporaryContentLease lease = await GoogleDriveTemporaryContentLease.CreatePublicReadLeaseAsync(
                client,
                content,
                new GoogleDriveUploadOptions { Name = "large.png", ContentType = "image/png" });
            await lease.CleanupAsync();

            Assert.Contains(requests, request => request.Contains("uploadType=resumable", StringComparison.Ordinal));
            Assert.DoesNotContain(requests, request => request.Contains("uploadType=multipart", StringComparison.Ordinal));
            Assert.Contains(requests, request => request.StartsWith("DELETE ", StringComparison.Ordinal));
        }

        [Fact]
        public async Task Test_DriveClient_ResumableUpload_UsesAlignedChunksAndProgress() {
            var ranges = new List<string>();
            using var httpClient = new HttpClient(new FakeHandler(async request => {
                if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri.Contains("uploadType=resumable", StringComparison.Ordinal)) {
                    var response = Json("{}");
                    response.Headers.Location = new Uri("https://upload.example.test/session-1");
                    return response;
                }

                if (request.Method == HttpMethod.Put && request.RequestUri!.AbsoluteUri == "https://upload.example.test/session-1") {
                    string range = request.Content!.Headers.GetValues("Content-Range").Single();
                    ranges.Add(range);
                    _ = await request.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                    if (ranges.Count < 3) {
                        var response = new HttpResponseMessage((HttpStatusCode)308) {
                            Content = new StringContent(string.Empty),
                        };
                        long end = ranges.Count == 1 ? 262143 : 524287;
                        response.Headers.TryAddWithoutValidation("Range", "bytes=0-" + end.ToString(System.Globalization.CultureInfo.InvariantCulture));
                        return response;
                    }

                    return new HttpResponseMessage(HttpStatusCode.Created) {
                        Content = new StringContent("{\"id\":\"resumable-1\",\"name\":\"large.bin\"}", Encoding.UTF8, "application/json")
                    };
                }

                return NotFound();
            }));
            using var client = CreateClient(httpClient);
            byte[] payload = new byte[600 * 1024];
            var progress = new List<long>();

            var file = await client.UploadResumableAsync(
                payload,
                new GoogleDriveUploadOptions {
                    Name = "large.bin",
                    ContentType = "application/octet-stream",
                    ResumableChunkSize = 256 * 1024,
                    Progress = new InlineProgress(value => progress.Add(value.BytesTransferred)),
                });

            Assert.Equal("resumable-1", file.Id);
            Assert.Equal(new[] {
                "bytes 0-262143/614400",
                "bytes 262144-524287/614400",
                "bytes 524288-614399/614400",
            }, ranges);
            Assert.Equal(614400, progress.Last());
        }

        [Fact]
        public async Task Test_DriveClient_ResumableUpload_CompletesZeroByteFiles() {
            string? contentRange = null;
            int bodyLength = -1;
            using var httpClient = new HttpClient(new FakeHandler(async request => {
                if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri.Contains("uploadType=resumable", StringComparison.Ordinal)) {
                    var response = Json("{}");
                    response.Headers.Location = new Uri("https://upload.example.test/session-empty");
                    return response;
                }

                if (request.Method == HttpMethod.Put && request.RequestUri!.AbsoluteUri == "https://upload.example.test/session-empty") {
                    contentRange = request.Content!.Headers.GetValues("Content-Range").Single();
                    bodyLength = (await request.Content.ReadAsByteArrayAsync().ConfigureAwait(false)).Length;
                    return new HttpResponseMessage(HttpStatusCode.Created) {
                        Content = new StringContent("{\"id\":\"resumable-empty\",\"name\":\"empty.bin\"}", Encoding.UTF8, "application/json")
                    };
                }

                return NotFound();
            }));
            using var client = CreateClient(httpClient, quotaUser: "tenant-user");

            GoogleDriveFile file = await client.UploadResumableAsync(
                Array.Empty<byte>(),
                new GoogleDriveUploadOptions { Name = "empty.bin", ContentType = "application/octet-stream" });

            Assert.Equal("resumable-empty", file.Id);
            Assert.Equal("bytes */0", contentRange);
            Assert.Equal(0, bodyLength);
        }

        [Fact]
        public async Task Test_DriveClient_ResumableUpload_RepeatsChunkWhenRangeIsMissing() {
            var ranges = new List<string>();
            using var httpClient = new HttpClient(new FakeHandler(async request => {
                if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri.Contains("uploadType=resumable", StringComparison.Ordinal)) {
                    var response = Json("{}");
                    response.Headers.Location = new Uri("https://upload.example.test/session-missing-range");
                    return response;
                }

                if (request.Method == HttpMethod.Put && request.RequestUri!.AbsoluteUri == "https://upload.example.test/session-missing-range") {
                    string range = request.Content!.Headers.GetValues("Content-Range").Single();
                    ranges.Add(range);
                    _ = await request.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                    if (ranges.Count == 1) {
                        return new HttpResponseMessage((HttpStatusCode)308) { Content = new StringContent(string.Empty) };
                    }

                    if (ranges.Count == 2) {
                        var acknowledged = new HttpResponseMessage((HttpStatusCode)308) { Content = new StringContent(string.Empty) };
                        acknowledged.Headers.TryAddWithoutValidation("Range", "bytes=0-262143");
                        return acknowledged;
                    }

                    return new HttpResponseMessage(HttpStatusCode.Created) {
                        Content = new StringContent("{\"id\":\"resumable-missing-range\",\"name\":\"large.bin\"}", Encoding.UTF8, "application/json")
                    };
                }

                return NotFound();
            }));
            using var client = CreateClient(httpClient);

            GoogleDriveFile file = await client.UploadResumableAsync(
                new byte[300 * 1024],
                new GoogleDriveUploadOptions {
                    Name = "large.bin",
                    ContentType = "application/octet-stream",
                    ResumableChunkSize = 256 * 1024,
                });

            Assert.Equal("resumable-missing-range", file.Id);
            Assert.Equal(new[] {
                "bytes 0-262143/307200",
                "bytes 0-262143/307200",
                "bytes 262144-307199/307200",
            }, ranges);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task Test_DriveClient_ResumableUpload_QueriesCommittedOffsetAfterAmbiguousTransportFailure(bool perRequestTimeout) {
            var ranges = new List<string>();
            int chunkAttempts = 0;
            using var httpClient = new HttpClient(new FakeHandler(async request => {
                if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri.Contains("uploadType=resumable", StringComparison.Ordinal)) {
                    var response = Json("{}");
                    response.Headers.Location = new Uri("https://upload.example.test/session-ambiguous");
                    return response;
                }

                if (request.Method == HttpMethod.Put && request.RequestUri!.AbsoluteUri == "https://upload.example.test/session-ambiguous") {
                    string range = request.Content!.Headers.GetValues("Content-Range").Single();
                    ranges.Add(range);
                    _ = await request.Content.ReadAsByteArrayAsync().ConfigureAwait(false);
                    if (range == "bytes 0-262143/307200" && ++chunkAttempts == 1) {
                        throw perRequestTimeout
                            ? new TaskCanceledException("per-request timeout after upload")
                            : new HttpRequestException("connection closed after upload");
                    }

                    if (range == "bytes */307200") {
                        var status = new HttpResponseMessage((HttpStatusCode)308) { Content = new StringContent(string.Empty) };
                        status.Headers.TryAddWithoutValidation("Range", "bytes=0-262143");
                        return status;
                    }

                    return new HttpResponseMessage(HttpStatusCode.Created) {
                        Content = new StringContent("{\"id\":\"resumable-recovered\",\"name\":\"large.bin\"}", Encoding.UTF8, "application/json")
                    };
                }

                return NotFound();
            }));
            using var client = CreateClient(httpClient, quotaUser: "tenant-user");

            GoogleDriveFile file = await client.UploadResumableAsync(
                new byte[300 * 1024],
                new GoogleDriveUploadOptions {
                    Name = "large.bin",
                    ContentType = "application/octet-stream",
                    ResumableChunkSize = 256 * 1024,
                });

            Assert.Equal("resumable-recovered", file.Id);
            Assert.Equal(new[] {
                "bytes 0-262143/307200",
                "bytes */307200",
                "bytes 262144-307199/307200",
            }, ranges);
        }

        [Fact]
        public async Task Test_DriveClient_MoveFile_NoOpReturnsCompleteMetadataWithoutPatch() {
            var methods = new List<HttpMethod>();
            Uri? metadataUri = null;
            using var httpClient = new HttpClient(new FakeHandler(request => {
                methods.Add(request.Method);
                metadataUri = request.RequestUri;
                return Task.FromResult(Json("{\"id\":\"file-1\",\"name\":\"Document\",\"parents\":[\"folder-1\"],\"version\":\"7\",\"size\":\"42\",\"modifiedTime\":\"2026-07-15T18:00:00Z\"}"));
            }));
            using var client = CreateClient(httpClient);

            GoogleDriveFile file = await client.MoveFileAsync("file-1", "folder-1");

            Assert.Equal(7, file.Version);
            Assert.Equal(42, file.Size);
            Assert.Equal(DateTimeOffset.Parse("2026-07-15T18:00:00Z"), file.ModifiedTime);
            Assert.Equal(new[] { HttpMethod.Get }, methods);
            Assert.Contains("version", metadataUri!.AbsoluteUri, StringComparison.Ordinal);
            Assert.Contains("modifiedTime", metadataUri.AbsoluteUri, StringComparison.Ordinal);
        }

        [Fact]
        public async Task Test_DriveClient_RejectsFolderFromUnexpectedSharedDrive() {
            using var httpClient = new HttpClient(new FakeHandler(request => Task.FromResult(Json(
                "{\"id\":\"folder-1\",\"name\":\"Folder\",\"mimeType\":\"application/vnd.google-apps.folder\",\"driveId\":\"drive-a\"}"))));
            using var client = CreateClient(httpClient);

            var exception = await Assert.ThrowsAsync<InvalidOperationException>(() =>
                client.ResolveFolderAsync("folder-1", "drive-b"));

            Assert.Contains("drive-a", exception.Message, StringComparison.Ordinal);
            Assert.Contains("drive-b", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public async Task Test_DriveClient_CommentsAndSharedDriveChanges_UseRequiredFields() {
            var uris = new List<string>();
            using var httpClient = new HttpClient(new FakeHandler(request => {
                uris.Add(request.RequestUri!.AbsoluteUri);
                if (request.RequestUri.AbsoluteUri.Contains("/comments?", StringComparison.Ordinal)) {
                    return Task.FromResult(Json("{\"id\":\"comment-1\",\"content\":\"Review\",\"anchor\":\"{\\\"r\\\":1}\"}"));
                }

                if (request.RequestUri.AbsoluteUri.Contains("/changes?", StringComparison.Ordinal)) {
                    return Task.FromResult(Json("{\"changes\":[{\"fileId\":\"file-1\",\"removed\":false}],\"newStartPageToken\":\"next-start\"}"));
                }

                return Task.FromResult(NotFound());
            }));
            using var client = CreateClient(httpClient);
            var report = new TranslationReport();

            var comment = await client.CreateCommentAsync("file-1", "Review", "{\"r\":1}", report);
            var changes = await client.ListChangesAsync("page-1", new GoogleDriveChangeListOptions { DriveId = "drive-1" }, report);

            Assert.Equal("comment-1", comment.Id);
            Assert.Equal("next-start", changes.NewStartPageToken);
            Assert.Contains(uris, uri => uri.Contains("fields=id,content,anchor,resolved", StringComparison.Ordinal));
            Assert.Contains(uris, uri => uri.Contains("driveId=drive-1", StringComparison.Ordinal));
            Assert.Contains(uris, uri => uri.Contains("includeItemsFromAllDrives=true", StringComparison.Ordinal));
            Assert.Contains(report.Notices, notice => notice.Code == "DRIVE.COMMENT.EDITOR_ANCHOR_UNAVAILABLE");
        }

        [Fact]
        public async Task Test_DriveClient_PaginatesPermissionsAndUsesDocumentedRevisionParameters() {
            var requests = new List<string>();
            using var httpClient = new HttpClient(new FakeHandler(request => {
                requests.Add(request.RequestUri!.AbsoluteUri);
                return Task.FromResult(request.RequestUri.AbsolutePath.EndsWith("/permissions", StringComparison.Ordinal)
                    ? Json("{\"permissions\":[],\"nextPageToken\":\"permission-next\"}")
                    : Json("{\"revisions\":[{\"id\":\"revision-1\",\"size\":\"84\"}],\"nextPageToken\":\"revision-next\"}"));
            }));
            using var client = CreateClient(httpClient);

            GoogleDrivePermissionList permissions = await client.ListPermissionsAsync("file-1", "permission page", 42);
            GoogleDriveRevisionList revisions = await client.ListRevisionsAsync("file-1", "revision page");

            Assert.Equal("permission-next", permissions.NextPageToken);
            Assert.Equal("revision-next", revisions.NextPageToken);
            Assert.Equal(84, Assert.Single(revisions.Revisions).Size);
            Assert.Contains(requests, request => request.Contains("/permissions?", StringComparison.Ordinal)
                && request.Contains("pageToken=permission%20page", StringComparison.Ordinal)
                && request.Contains("pageSize=42", StringComparison.Ordinal)
                && request.Contains("supportsAllDrives=true", StringComparison.Ordinal));
            Assert.Contains(requests, request => request.Contains("/revisions?", StringComparison.Ordinal)
                && request.Contains("pageToken=revision%20page", StringComparison.Ordinal)
                && !request.Contains("supportsAllDrives", StringComparison.Ordinal));
        }

        [Fact]
        public async Task Test_DriveClient_CommentEndpoints_UseDocumentedParameters() {
            var requests = new List<string>();
            using var httpClient = new HttpClient(new FakeHandler(request => {
                requests.Add(request.Method.Method + " " + request.RequestUri!.AbsoluteUri);
                if (request.Method == HttpMethod.Get) {
                    return Task.FromResult(Json("{\"comments\":[]}"));
                }
                if (request.Method == HttpMethod.Post && request.RequestUri.AbsolutePath.EndsWith("/replies", StringComparison.Ordinal)) {
                    return Task.FromResult(Json("{\"id\":\"reply-1\",\"content\":\"Reply\"}"));
                }
                if (request.Method == HttpMethod.Post) {
                    return Task.FromResult(Json("{\"id\":\"comment-1\",\"content\":\"Comment\"}"));
                }
                return Task.FromResult(Json("{}"));
            }));
            using var client = CreateClient(httpClient);

            await client.ListCommentsAsync("file-1");
            await client.CreateCommentAsync("file-1", "Comment");
            await client.CreateReplyAsync("file-1", "comment-1", "Reply");
            await client.DeleteReplyAsync("file-1", "comment-1", "reply-1");
            await client.DeleteCommentAsync("file-1", "comment-1");

            Assert.Equal(5, requests.Count);
            Assert.DoesNotContain(requests, request => request.Contains("supportsAllDrives", StringComparison.Ordinal));
            Assert.Contains(requests, request => request == "DELETE https://www.googleapis.com/drive/v3/files/file-1/comments/comment-1/replies/reply-1");
            Assert.Contains(requests, request => request == "DELETE https://www.googleapis.com/drive/v3/files/file-1/comments/comment-1");
        }

        private static GoogleDriveClient CreateClient(HttpClient httpClient, RecordingCredentialSource? credential = null, string? quotaUser = null) {
            return new GoogleDriveClient(new GoogleWorkspaceSession(
                credential ?? new RecordingCredentialSource(),
                new GoogleWorkspaceSessionOptions {
                    HttpClient = httpClient,
                    MaxRetryCount = 1,
                    QuotaUser = quotaUser,
                }));
        }

        private static HttpResponseMessage Json(string json) {
            return new HttpResponseMessage(HttpStatusCode.OK) {
                Content = new StringContent(json, Encoding.UTF8, "application/json")
            };
        }

        private static HttpResponseMessage NotFound() {
            return new HttpResponseMessage(HttpStatusCode.NotFound) {
                Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
            };
        }

        private sealed class RecordingCredentialSource : IGoogleWorkspaceCredentialSource {
            public IReadOnlyList<string> LastScopes { get; private set; } = Array.Empty<string>();

            public Task<GoogleWorkspaceAccessToken> AcquireAccessTokenAsync(IEnumerable<string> scopes, CancellationToken cancellationToken = default) {
                LastScopes = scopes.ToArray();
                return Task.FromResult(new GoogleWorkspaceAccessToken("token", DateTimeOffset.UtcNow.AddHours(1), LastScopes));
            }
        }

        private sealed class FakeHandler : HttpMessageHandler {
            private readonly Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> _handler;

            public FakeHandler(Func<HttpRequestMessage, Task<HttpResponseMessage>> handler) {
                _handler = (request, _) => handler(request);
            }

            public FakeHandler(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> handler) {
                _handler = handler;
            }

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
                return _handler(request, cancellationToken);
            }
        }

        private sealed class InlineProgress : IProgress<GoogleDriveTransferProgress> {
            private readonly Action<GoogleDriveTransferProgress> _action;

            public InlineProgress(Action<GoogleDriveTransferProgress> action) {
                _action = action;
            }

            public void Report(GoogleDriveTransferProgress value) {
                _action(value);
            }
        }
    }
}
