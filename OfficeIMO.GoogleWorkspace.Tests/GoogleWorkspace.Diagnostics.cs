using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Excel;
using OfficeIMO.Excel.GoogleSheets;
using OfficeIMO.Word;
using OfficeIMO.Word.GoogleDocs;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public class GoogleWorkspaceDiagnosticsTests {
        [Fact]
        public void Test_TranslationReport_ToDiagnosticEntries_ProjectsNotices() {
            var report = new TranslationReport();
            report.Add(TranslationSeverity.Warning, "DrivePlacement", "FolderId is missing.", path: "Location.FolderId");
            report.Add(TranslationSeverity.Info, "ApiRetries", "Retried request.");

            var entries = report.ToDiagnosticEntries();

            Assert.Equal(2, entries.Count);

            Assert.Equal(TranslationSeverity.Warning, entries[0].Severity);
            Assert.Equal("DrivePlacement", entries[0].Feature);
            Assert.Equal("FolderId is missing.", entries[0].Message);
            Assert.Equal("Location.FolderId", entries[0].Path);
            Assert.Null(entries[0].FailureKind);

            Assert.Equal(TranslationSeverity.Info, entries[1].Severity);
            Assert.Equal("ApiRetries", entries[1].Feature);
            Assert.Equal("Retried request.", entries[1].Message);
            Assert.Equal(string.Empty, entries[1].Path);
        }

        [Fact]
        public void Test_GoogleWorkspaceExportException_ToDiagnosticEntries_IncludesFailureSummary() {
            var report = new TranslationReport();
            report.Add(TranslationSeverity.Error, "Authentication", "Token acquisition failed.");
            report.Add(TranslationSeverity.Warning, "DrivePlacement", "Folder was not changed.");

            var exception = new GoogleWorkspaceExportException(
                "Google Docs export could not acquire a token.",
                GoogleWorkspaceFailureKind.TokenAcquisition,
                report,
                new InvalidOperationException("boom"));

            var entries = exception.ToDiagnosticEntries();

            Assert.Equal(3, entries.Count);

            Assert.Equal(TranslationSeverity.Error, entries[0].Severity);
            Assert.Equal("ExportFailure", entries[0].Feature);
            Assert.Equal("Google Docs export could not acquire a token.", entries[0].Message);
            Assert.Equal(GoogleWorkspaceFailureKind.TokenAcquisition, entries[0].FailureKind);

            Assert.Equal("Authentication", entries[1].Feature);
            Assert.Equal(GoogleWorkspaceFailureKind.TokenAcquisition, entries[1].FailureKind);
            Assert.Equal("DrivePlacement", entries[2].Feature);
            Assert.Equal(GoogleWorkspaceFailureKind.TokenAcquisition, entries[2].FailureKind);
        }

        [Fact]
        public async Task Test_GoogleSheetsExporter_DoesNotRetryAmbiguousCreate() {
            string filePath = Path.Combine(Path.GetTempPath(), "GoogleSheetsExporterNoAmbiguousCreateRetry-" + Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using var document = ExcelDocument.Create(filePath);
                var summary = document.AddWorksheet("Summary");
                summary.CellValue(1, 1, "Name");
                summary.CellValue(2, 1, "Alpha");

                var entries = new List<GoogleWorkspaceDiagnosticEntry>();
                int createAttempts = 0;
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets") {
                        createAttempts++;
                        if (createAttempts == 1) {
                            return Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable) {
                                Content = new StringContent("retry later", Encoding.UTF8, "text/plain")
                            });
                        }

                        return Task.FromResult(CreateJsonResponse("{\"spreadsheetId\":\"diagRetry\",\"spreadsheetUrl\":\"https://docs.google.com/spreadsheets/d/diagRetry/edit\",\"properties\":{\"title\":\"Diagnostic Retry Export\"}}"));
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets/diagRetry:batchUpdate") {
                        return Task.FromResult(CreateJsonResponse("{}"));
                    }

                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    });
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                        MaxRetryCount = 1,
                        DiagnosticSink = entries.Add,
                    });

                var exception = await Assert.ThrowsAsync<GoogleWorkspaceExportException>(() =>
                    document.ExportToGoogleSheetsAsync(session, new GoogleSheetsSaveOptions {
                        Title = "Diagnostic Retry Export",
                    }));

                Assert.Equal(GoogleWorkspaceFailureKind.ApiRequest, exception.FailureKind);
                Assert.Equal(1, createAttempts);
                Assert.DoesNotContain(entries, entry => entry.Code == GoogleWorkspaceDiagnosticCodes.ApiRetry);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleWorkspaceHttpTransport_RetriesSafeReads_AndSendsContext() {
            int attempts = 0;
            var requests = new List<(string Uri, string? UserAgent, string? QuotaProject, string? RequestId)>();
            using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                attempts++;
                requests.Add((
                    request.RequestUri!.AbsoluteUri,
                    request.Headers.UserAgent.ToString(),
                    request.Headers.TryGetValues("X-Goog-User-Project", out var quotaProjects) ? quotaProjects.Single() : null,
                    request.Headers.TryGetValues("X-Request-Id", out var requestIds) ? requestIds.Single() : null));

                if (attempts == 1) {
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable) {
                        Content = new StringContent("retry later", Encoding.UTF8, "text/plain")
                    });
                }

                return Task.FromResult(CreateJsonResponse("{\"id\":\"safe-read\"}"));
            }));

            var entries = new List<GoogleWorkspaceDiagnosticEntry>();
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = httpClient,
                ApplicationName = "OfficeIMO Integration Tests",
                QuotaUser = "tenant-user",
                QuotaProject = "billing-project",
                RequestIdFactory = () => "request-123",
                MaxRetryCount = 1,
                DiagnosticSink = entries.Add,
            };
            var report = new TranslationReport();
            using var transport = new GoogleWorkspaceHttpTransport(options);

            var result = await transport.SendJsonAsync<TransportReadResponse>(
                "token",
                HttpMethod.Get,
                "https://www.googleapis.com/drive/v3/files/file-1?fields=id",
                null,
                GoogleWorkspaceRequestSafety.Safe,
                "Google Drive API",
                report);

            Assert.Equal("safe-read", result.Id);
            Assert.Equal(2, attempts);
            Assert.All(requests, request => {
                Assert.Contains("quotaUser=tenant-user", request.Uri, StringComparison.Ordinal);
                Assert.Equal("OfficeIMO-Integration-Tests/2.0", request.UserAgent);
                Assert.Equal("billing-project", request.QuotaProject);
                Assert.Equal("request-123", request.RequestId);
            });
            Assert.Contains(entries, entry =>
                entry.Code == GoogleWorkspaceDiagnosticCodes.ApiRetry
                && entry.Feature == "ApiRetries"
                && entry.Severity == TranslationSeverity.Info);
        }

        [Fact]
        public async Task Test_GoogleWorkspaceHttpTransport_RetriesSafePerAttemptTimeouts() {
            int attempts = 0;
            using var httpClient = new HttpClient(new FakeHttpMessageHandler(async (_, cancellationToken) => {
                attempts++;
                if (attempts == 1) {
                    await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken).ConfigureAwait(false);
                }

                return CreateJsonResponse("{\"id\":\"safe-timeout-retry\"}");
            }));

            var entries = new List<GoogleWorkspaceDiagnosticEntry>();
            var options = new GoogleWorkspaceSessionOptions {
                HttpClient = httpClient,
                RequestTimeout = TimeSpan.FromMilliseconds(50),
                MaxRetryCount = 1,
                RetryBaseDelay = TimeSpan.FromMilliseconds(1),
                RetryMaxDelay = TimeSpan.FromMilliseconds(1),
                DiagnosticSink = entries.Add,
            };
            using var transport = new GoogleWorkspaceHttpTransport(options);

            TransportReadResponse result = await transport.SendJsonAsync<TransportReadResponse>(
                "token",
                HttpMethod.Get,
                "https://www.googleapis.com/drive/v3/files/file-1?fields=id",
                null,
                GoogleWorkspaceRequestSafety.Safe,
                "Google Drive API",
                new TranslationReport());

            Assert.Equal("safe-timeout-retry", result.Id);
            Assert.Equal(2, attempts);
            Assert.Contains(entries, entry => entry.Code == GoogleWorkspaceDiagnosticCodes.ApiRetry);
        }

        [Fact]
        public void Test_GoogleWorkspaceHttpTransport_OwnedClientDoesNotOverridePerAttemptTimeout() {
            var ownedOptions = new GoogleWorkspaceSessionOptions {
                RequestTimeout = TimeSpan.FromMinutes(5),
            };
            using var ownedTransport = new GoogleWorkspaceHttpTransport(ownedOptions);
            var clientField = typeof(GoogleWorkspaceHttpTransport).GetField(
                "_client",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            HttpClient ownedClient = Assert.IsType<HttpClient>(clientField!.GetValue(ownedTransport));
            Assert.Equal(Timeout.InfiniteTimeSpan, ownedClient.Timeout);

            using var injectedClient = new HttpClient {
                Timeout = TimeSpan.FromSeconds(17),
            };
            using var injectedTransport = new GoogleWorkspaceHttpTransport(new GoogleWorkspaceSessionOptions {
                HttpClient = injectedClient,
                RequestTimeout = TimeSpan.FromMinutes(5),
            });
            HttpClient preservedClient = Assert.IsType<HttpClient>(clientField.GetValue(injectedTransport));
            Assert.Same(injectedClient, preservedClient);
            Assert.Equal(TimeSpan.FromSeconds(17), preservedClient.Timeout);
        }

        [Fact]
        public void Test_GoogleWorkspacePreflight_BlocksUnacceptedErrorsBeforeMutation() {
            var report = new TranslationReport();
            report.Add(
                TranslationSeverity.Error,
                "Charts",
                "Chart rasterization is unavailable.",
                path: "Slides[0].Charts[0]",
                code: "SLIDES.CHART.RASTERIZER_UNAVAILABLE",
                action: TranslationAction.Fail);

            var exception = Assert.Throws<GoogleWorkspacePreflightException>(() =>
                GoogleWorkspacePreflight.Validate(report, new GoogleWorkspaceFidelityPolicy()));

            var notice = Assert.Single(exception.BlockingNotices);
            Assert.Equal("SLIDES.CHART.RASTERIZER_UNAVAILABLE", notice.Code);
            Assert.Equal(TranslationAction.Fail, notice.Action);
            Assert.Equal("Slides[0].Charts[0]", notice.Path);
        }

        [Fact]
        public async Task Test_GoogleSheetsExporter_DiagnosticSink_ReceivesTokenFailureEntries() {
            string filePath = Path.Combine(Path.GetTempPath(), "GoogleSheetsExporterDiagnosticSinkAuth-" + Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using var document = ExcelDocument.Create(filePath);
                var summary = document.AddWorksheet("Summary");
                summary.CellValue(1, 1, "Name");
                summary.CellValue(2, 1, "Alpha");

                var entries = new List<GoogleWorkspaceDiagnosticEntry>();
                var session = new GoogleWorkspaceSession(
                    new DelegateGoogleWorkspaceCredentialSource((scopes, cancellationToken) =>
                        Task.FromException<GoogleWorkspaceAccessToken>(new HttpRequestException("token endpoint unavailable"))),
                    new GoogleWorkspaceSessionOptions {
                        DiagnosticSink = entries.Add,
                    });

                var exception = await Assert.ThrowsAsync<GoogleWorkspaceExportException>(() =>
                    document.ExportToGoogleSheetsAsync(session, new GoogleSheetsSaveOptions {
                        Title = "Diagnostic Auth Export",
                    }));

                Assert.Equal(GoogleWorkspaceFailureKind.TokenAcquisition, exception.FailureKind);
                Assert.Contains(entries, entry =>
                    entry.Feature == "Authentication"
                    && entry.FailureKind == GoogleWorkspaceFailureKind.TokenAcquisition
                    && entry.Message.Contains("token endpoint unavailable", StringComparison.Ordinal));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_DiagnosticSink_ReceivesApiFailureEntries() {
            string filePath = Path.Combine(Path.GetTempPath(), "GoogleDocsExporterDiagnosticSinkApi-" + Guid.NewGuid().ToString("N") + ".docx");

            try {
                using var document = WordDocument.Create(filePath);
                document.AddParagraph("API failure");

                var entries = new List<GoogleWorkspaceDiagnosticEntry>();
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return Task.FromResult(new HttpResponseMessage(HttpStatusCode.Forbidden) {
                            Content = new StringContent("{\"error\":{\"code\":403,\"message\":\"The caller does not have permission\",\"status\":\"PERMISSION_DENIED\",\"errors\":[{\"message\":\"The caller does not have permission\",\"domain\":\"global\",\"reason\":\"forbidden\"}]}}", Encoding.UTF8, "application/json")
                        });
                    }

                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    });
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                        DiagnosticSink = entries.Add,
                    });

                var exception = await Assert.ThrowsAsync<GoogleWorkspaceExportException>(() =>
                    document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                        Title = "Diagnostic API Export",
                    }));

                Assert.Equal(GoogleWorkspaceFailureKind.ApiRequest, exception.FailureKind);
                Assert.Contains("PERMISSION_DENIED", exception.Message, StringComparison.Ordinal);
                Assert.Contains("reason=forbidden", exception.Message, StringComparison.Ordinal);
                Assert.Contains(entries, entry =>
                    entry.Feature == "ApiFailures"
                    && entry.FailureKind == GoogleWorkspaceFailureKind.ApiRequest
                    && entry.Message.Contains("https://docs.googleapis.com/v1/documents", StringComparison.Ordinal)
                    && entry.Message.Contains("PERMISSION_DENIED", StringComparison.Ordinal));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleSheetsExporter_EnrichesGoogleApiFailureMessages() {
            string filePath = Path.Combine(Path.GetTempPath(), "GoogleSheetsExporterApiError-" + Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using var document = ExcelDocument.Create(filePath);
                var summary = document.AddWorksheet("Summary");
                summary.CellValue(1, 1, "Name");
                summary.CellValue(2, 1, "Alpha");

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets") {
                        return Task.FromResult(new HttpResponseMessage(HttpStatusCode.Forbidden) {
                            Content = new StringContent("{\"error\":{\"code\":403,\"message\":\"Request had insufficient authentication scopes.\",\"status\":\"PERMISSION_DENIED\",\"errors\":[{\"message\":\"Insufficient Permission\",\"domain\":\"global\",\"reason\":\"insufficientPermissions\"}]}}", Encoding.UTF8, "application/json")
                        });
                    }

                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    });
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var exception = await Assert.ThrowsAsync<GoogleWorkspaceExportException>(() =>
                    document.ExportToGoogleSheetsAsync(session, new GoogleSheetsSaveOptions {
                        Title = "API Error Export",
                    }));

                Assert.Equal(GoogleWorkspaceFailureKind.ApiRequest, exception.FailureKind);
                Assert.Contains("PERMISSION_DENIED", exception.Message, StringComparison.Ordinal);
                Assert.Contains("reason=insufficientPermissions", exception.Message, StringComparison.Ordinal);
                Assert.Contains(exception.Report.Notices, notice =>
                    notice.Feature == "ApiFailures"
                    && notice.Message.Contains("PERMISSION_DENIED", StringComparison.Ordinal)
                    && notice.Message.Contains("reason=insufficientPermissions", StringComparison.Ordinal));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleSheetsExporter_ClassifiesRequestTimeoutFailures() {
            string filePath = Path.Combine(Path.GetTempPath(), "GoogleSheetsExporterTimeout-" + Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using var document = ExcelDocument.Create(filePath);
                var summary = document.AddWorksheet("Summary");
                summary.CellValue(1, 1, "Name");
                summary.CellValue(2, 1, "Alpha");

                using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://sheets.googleapis.com/v4/spreadsheets") {
                        throw new TaskCanceledException("request timeout");
                    }

                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    });
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions {
                        HttpClient = httpClient,
                    });

                var exception = await Assert.ThrowsAsync<GoogleWorkspaceExportException>(() =>
                    document.ExportToGoogleSheetsAsync(session, new GoogleSheetsSaveOptions {
                        Title = "Timeout Export",
                    }));

                Assert.Equal(GoogleWorkspaceFailureKind.RequestTimeout, exception.FailureKind);
                Assert.IsType<TaskCanceledException>(exception.InnerException);
                Assert.Contains(exception.Report.Notices, notice =>
                    notice.Feature == "RequestTimeout"
                    && notice.Severity == TranslationSeverity.Error);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_ClassifiesCallerCancellation() {
            string filePath = Path.Combine(Path.GetTempPath(), "GoogleDocsExporterCanceled-" + Guid.NewGuid().ToString("N") + ".docx");

            try {
                using var document = WordDocument.Create(filePath);
                document.AddParagraph("Canceled export");

                var entries = new List<GoogleWorkspaceDiagnosticEntry>();
                var session = new GoogleWorkspaceSession(
                    new DelegateGoogleWorkspaceCredentialSource((scopes, cancellationToken) =>
                        Task.FromCanceled<GoogleWorkspaceAccessToken>(cancellationToken)),
                    new GoogleWorkspaceSessionOptions {
                        DiagnosticSink = entries.Add,
                    });

                using var cancellationTokenSource = new CancellationTokenSource();
                cancellationTokenSource.Cancel();

                var exception = await Assert.ThrowsAsync<GoogleWorkspaceExportCanceledException>(() =>
                    document.ExportToGoogleDocsAsync(
                        session,
                        new GoogleDocsSaveOptions {
                            Title = "Canceled Export",
                        },
                        cancellationTokenSource.Token));

                Assert.Equal(GoogleWorkspaceFailureKind.Canceled, exception.FailureKind);
                Assert.Contains(exception.Report.Notices, notice =>
                    notice.Feature == "Cancellation"
                    && notice.Severity == TranslationSeverity.Warning);
                Assert.Contains(entries, entry =>
                    entry.Feature == "Cancellation"
                    && entry.FailureKind == GoogleWorkspaceFailureKind.Canceled);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleWorkspaceHttpTransport_BoundsUnknownLengthByteResponses() {
            using var httpClient = new HttpClient(new FakeHttpMessageHandler(_ =>
                Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new UnknownLengthContent(new byte[128])
                })));
            using var transport = new GoogleWorkspaceHttpTransport(
                new GoogleWorkspaceSessionOptions {
                    HttpClient = httpClient,
                    MaxRetryCount = 0
                });

            await Assert.ThrowsAsync<InvalidDataException>(() =>
                transport.SendBytesAsync(
                    "token",
                    HttpMethod.Get,
                    "https://lh3.googleusercontent.com/image.png",
                    GoogleWorkspaceRequestSafety.Safe,
                    "Google content",
                    new TranslationReport(),
                    maxResponseBytes: 8));
        }

        [Fact]
        public async Task Test_GoogleWorkspaceHttpTransport_TruncatesUnknownLengthErrorResponses() {
            byte[] responseBytes = Encoding.UTF8.GetBytes(
                new string('x', 128 * 1024));
            using var httpClient = new HttpClient(new FakeHttpMessageHandler(_ =>
                Task.FromResult(new HttpResponseMessage(HttpStatusCode.BadRequest) {
                    Content = new UnknownLengthContent(responseBytes)
                })));
            using var transport = new GoogleWorkspaceHttpTransport(
                new GoogleWorkspaceSessionOptions {
                    HttpClient = httpClient,
                    MaxRetryCount = 0
                });

            GoogleWorkspaceApiException exception =
                await Assert.ThrowsAsync<GoogleWorkspaceApiException>(() =>
                    transport.SendBytesAsync(
                        "token",
                        HttpMethod.Get,
                        "https://lh3.googleusercontent.com/image.png",
                        GoogleWorkspaceRequestSafety.Safe,
                        "Google content",
                        new TranslationReport(),
                        maxResponseBytes: 8));

            Assert.Equal(64 * 1024, exception.ResponseBody.Length);
        }

        [Fact]
        public async Task Test_GoogleWorkspaceHttpTransport_TimesOutWhileReadingUnboundedByteResponse() {
            using var httpClient = new HttpClient(new FakeHttpMessageHandler(_ =>
                Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new BlockingReadContent()
                })));
            using var transport = new GoogleWorkspaceHttpTransport(
                new GoogleWorkspaceSessionOptions {
                    HttpClient = httpClient,
                    RequestTimeout = TimeSpan.FromMilliseconds(50),
                    MaxRetryCount = 0
                });

            await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
                transport.SendBytesAsync(
                    "token",
                    HttpMethod.Get,
                    "https://www.googleapis.com/drive/v3/files/file-1?alt=media",
                    GoogleWorkspaceRequestSafety.Safe,
                    "Google Drive API",
                    new TranslationReport()));
        }

        private static HttpResponseMessage CreateJsonResponse(string json) {
            return new HttpResponseMessage(HttpStatusCode.OK) {
                Content = new StringContent(json, Encoding.UTF8, "application/json")
            };
        }

        private sealed class FakeGoogleWorkspaceCredentialSource : IGoogleWorkspaceCredentialSource {
            public Task<GoogleWorkspaceAccessToken> AcquireAccessTokenAsync(IEnumerable<string> scopes, CancellationToken cancellationToken = default) {
                return Task.FromResult(new GoogleWorkspaceAccessToken(
                    "fake-access-token",
                    DateTimeOffset.UtcNow.AddHours(1),
                    scopes.ToArray()));
            }
        }

        private sealed class TransportReadResponse {
            [System.Text.Json.Serialization.JsonPropertyName("id")]
            public string? Id { get; set; }
        }

        private sealed class UnknownLengthContent : HttpContent {
            private readonly byte[] _bytes;

            internal UnknownLengthContent(byte[] bytes) {
                _bytes = bytes;
            }

            protected override Task SerializeToStreamAsync(
                Stream stream,
                TransportContext? context) =>
                throw new InvalidOperationException(
                    "ResponseContentRead attempted to buffer the response.");

            protected override Task<Stream> CreateContentReadStreamAsync() =>
                Task.FromResult<Stream>(new MemoryStream(_bytes,
                    writable: false));

            protected override bool TryComputeLength(out long length) {
                length = 0;
                return false;
            }
        }

        private sealed class BlockingReadContent : HttpContent {
            protected override Task SerializeToStreamAsync(Stream stream, TransportContext? context) =>
                throw new InvalidOperationException("ResponseContentRead attempted to buffer the response.");

            protected override Task<Stream> CreateContentReadStreamAsync() =>
                Task.FromResult<Stream>(new BlockingReadStream());

            protected override bool TryComputeLength(out long length) {
                length = 0;
                return false;
            }
        }

        private sealed class BlockingReadStream : Stream {
            public override bool CanRead => true;
            public override bool CanSeek => false;
            public override bool CanWrite => false;
            public override long Length => throw new NotSupportedException();
            public override long Position {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }

            public override void Flush() { }
            public override int Read(byte[] buffer, int offset, int count) =>
                throw new InvalidOperationException("The response body must be read asynchronously.");
            public override async Task<int> ReadAsync(byte[] buffer, int offset, int count,
                CancellationToken cancellationToken) {
                await Task.Delay(Timeout.InfiniteTimeSpan, cancellationToken).ConfigureAwait(false);
                return 0;
            }
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        }

        private sealed class FakeHttpMessageHandler : HttpMessageHandler {
            private readonly Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> _handler;

            public FakeHttpMessageHandler(Func<HttpRequestMessage, Task<HttpResponseMessage>> handler) {
                if (handler == null) throw new ArgumentNullException(nameof(handler));
                _handler = (request, _) => handler(request);
            }

            public FakeHttpMessageHandler(Func<HttpRequestMessage, CancellationToken, Task<HttpResponseMessage>> handler) {
                _handler = handler ?? throw new ArgumentNullException(nameof(handler));
            }

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
                return _handler(request, cancellationToken);
            }
        }
    }
}
