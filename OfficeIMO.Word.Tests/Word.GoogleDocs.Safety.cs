using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Word.GoogleDocs;
using System.Net;
using System.Net.Http;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public async Task Test_GoogleDocsExporter_DefaultImagePolicy_NeverUsesPublicDriveStaging() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsDefaultImageSafety.docx");
            string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");

            try {
                using var document = BuildGoogleDocsTableImageDocument(filePath, imagePath);
                var requestUris = new List<string>();
                int batchUpdateCount = 0;
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                    requestUris.Add(request.RequestUri!.AbsoluteUri);

                    if (request.Method == HttpMethod.Post && request.RequestUri.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return Task.FromResult(CreateJsonResponse("{\"documentId\":\"doc-safe-image\",\"title\":\"Safe Image Export\"}"));
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-safe-image:batchUpdate") {
                        batchUpdateCount++;
                        return Task.FromResult(CreateJsonResponse("{}"));
                    }

                    if (request.Method == HttpMethod.Get && request.RequestUri.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-safe-image") {
                        return Task.FromResult(CreateJsonResponse(CreateBodyTableDocumentStateJson("doc-safe-image", "Safe Image Export")));
                    }

                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    });
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                var result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Title = "Safe Image Export",
                });

                Assert.Equal("doc-safe-image", result.DocumentId);
                Assert.Equal(3, batchUpdateCount);
                Assert.DoesNotContain(requestUris, uri => uri.Contains("drive/v3", StringComparison.Ordinal));
                Assert.Contains(result.Report.Notices, notice =>
                    notice.Code == "DOCS.IMAGE.STAGING_DISABLED"
                    && notice.Action == TranslationAction.Skip);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_TemporaryImageLease_CleansUpAfterContentFailure() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsImageCleanupFailure.docx");
            string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");

            try {
                using var document = BuildGoogleDocsSampleDocument(filePath, imagePath);
                bool deleted = false;
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return Task.FromResult(CreateJsonResponse("{\"documentId\":\"doc-cleanup\",\"title\":\"Cleanup Export\"}"));
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri.StartsWith("https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart", StringComparison.Ordinal)) {
                        return Task.FromResult(CreateJsonResponse("{\"id\":\"temporary-image\"}"));
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri.StartsWith("https://www.googleapis.com/drive/v3/files/temporary-image/permissions?", StringComparison.Ordinal)) {
                        return Task.FromResult(CreateJsonResponse("{\"id\":\"public-reader\",\"type\":\"anyone\",\"role\":\"reader\"}"));
                    }

                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-cleanup:batchUpdate") {
                        return Task.FromResult(new HttpResponseMessage(HttpStatusCode.BadRequest) {
                            Content = new StringContent("invalid content", Encoding.UTF8, "text/plain")
                        });
                    }

                    if (request.Method == HttpMethod.Delete && request.RequestUri!.AbsoluteUri == "https://www.googleapis.com/drive/v3/files/temporary-image?supportsAllDrives=true") {
                        deleted = true;
                        return Task.FromResult(CreateJsonResponse("{}"));
                    }

                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound) {
                        Content = new StringContent("unexpected request", Encoding.UTF8, "text/plain")
                    });
                }));

                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                var exception = await Assert.ThrowsAsync<GoogleWorkspaceExportException>(() =>
                    document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                        Title = "Cleanup Export",
                        InlineImageMode = GoogleDocsInlineImageMode.TemporaryPublicDriveLease,
                    }));

                Assert.Equal(GoogleWorkspaceFailureKind.ApiRequest, exception.FailureKind);
                Assert.True(deleted);
                Assert.Contains(exception.Report.Notices, notice => notice.Code == "DRIVE.TEMPORARY_CONTENT.CLEANED");
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
