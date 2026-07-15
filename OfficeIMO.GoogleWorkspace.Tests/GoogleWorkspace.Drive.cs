using OfficeIMO.GoogleWorkspace;
using OfficeIMO.GoogleWorkspace.Drive;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public class GoogleWorkspaceDriveTests {
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
            Assert.Contains(report.Notices, notice => notice.Code == "DRIVE.COMMENT.EDITOR_ANCHOR_UNAVAILABLE");
        }

        private static GoogleDriveClient CreateClient(HttpClient httpClient, RecordingCredentialSource? credential = null) {
            return new GoogleDriveClient(new GoogleWorkspaceSession(
                credential ?? new RecordingCredentialSource(),
                new GoogleWorkspaceSessionOptions {
                    HttpClient = httpClient,
                    MaxRetryCount = 1,
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
            private readonly Func<HttpRequestMessage, Task<HttpResponseMessage>> _handler;

            public FakeHandler(Func<HttpRequestMessage, Task<HttpResponseMessage>> handler) {
                _handler = handler;
            }

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
                return _handler(request);
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
