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
        public async Task Test_GoogleSheetsExporter_DiagnosticSink_ReceivesRetryEntries() {
            string filePath = Path.Combine(Path.GetTempPath(), "GoogleSheetsExporterDiagnosticSinkRetry-" + Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using var document = ExcelDocument.Create(filePath);
                var summary = document.AddWorkSheet("Summary");
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

                var result = await document.ExportToGoogleSheetsAsync(session, new GoogleSheetsSaveOptions {
                    Title = "Diagnostic Retry Export",
                });

                Assert.Equal("diagRetry", result.SpreadsheetId);
                Assert.Contains(entries, entry =>
                    entry.Feature == "ApiRetries"
                    && entry.Severity == TranslationSeverity.Info
                    && entry.Message.Contains("https://sheets.googleapis.com/v4/spreadsheets", StringComparison.Ordinal));
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleSheetsExporter_DiagnosticSink_ReceivesTokenFailureEntries() {
            string filePath = Path.Combine(Path.GetTempPath(), "GoogleSheetsExporterDiagnosticSinkAuth-" + Guid.NewGuid().ToString("N") + ".xlsx");

            try {
                using var document = ExcelDocument.Create(filePath);
                var summary = document.AddWorkSheet("Summary");
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
                var summary = document.AddWorkSheet("Summary");
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
                var summary = document.AddWorkSheet("Summary");
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

        private sealed class FakeHttpMessageHandler : HttpMessageHandler {
            private readonly Func<HttpRequestMessage, Task<HttpResponseMessage>> _handler;

            public FakeHttpMessageHandler(Func<HttpRequestMessage, Task<HttpResponseMessage>> handler) {
                _handler = handler ?? throw new ArgumentNullException(nameof(handler));
            }

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
                return _handler(request);
            }
        }
    }
}
