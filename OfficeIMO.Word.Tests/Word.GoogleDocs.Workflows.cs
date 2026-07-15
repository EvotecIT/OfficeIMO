using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Word;
using OfficeIMO.Word.GoogleDocs;
using System.Net;
using System.Net.Http;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_GoogleDocsFeatureSupportCatalog_IsCodeOwnedAndBidirectional() {
            Assert.NotEmpty(GoogleDocsFeatureSupportCatalog.Features);
            Assert.Contains(GoogleDocsFeatureSupportCatalog.Features, feature => feature.Feature == "Document tabs" && feature.Export == GoogleDocsFeatureSupportLevel.Native);
            Assert.Contains(GoogleDocsFeatureSupportCatalog.Features, feature => feature.Import == GoogleDocsFeatureSupportLevel.DriveFallback);
            Assert.All(GoogleDocsFeatureSupportCatalog.Features, feature => Assert.False(string.IsNullOrWhiteSpace(feature.Notes)));
        }

        [Fact]
        public void Test_GoogleDocsResetPayload_ReplacesEveryTabAndItsSegments() {
            var document = new GoogleDocsApiDocumentResponse();
            document.Tabs.Add(CreateTabState("tab-one", "First", 8, "header-one", "footer-one", "range-one"));
            document.Tabs.Add(CreateTabState("tab-two", "Second", 5, "header-two", "footer-two", "range-two"));

            GoogleDocsApiBatchUpdatePayload payload = GoogleDocsApiPayloadBuilder.BuildResetDocumentPayload(
                document,
                new GoogleDocsTabOptions { Strategy = GoogleDocsTabStrategy.ReplaceEveryTab });

            Assert.Equal(2, payload.Requests.Count(request => request.DeleteContentRange != null));
            Assert.Contains(payload.Requests, request => request.DeleteContentRange?.Range.TabId == "tab-one");
            Assert.Contains(payload.Requests, request => request.DeleteContentRange?.Range.TabId == "tab-two");
            Assert.Contains(payload.Requests, request => request.DeleteHeader?.HeaderId == "header-one" && request.DeleteHeader.TabId == "tab-one");
            Assert.Contains(payload.Requests, request => request.DeleteFooter?.FooterId == "footer-two" && request.DeleteFooter.TabId == "tab-two");
            Assert.Contains(payload.Requests, request => request.DeleteNamedRange?.Name == "range-two" && request.DeleteNamedRange.TabsCriteria?.TabIds.Single() == "tab-two");
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_ReplaceEveryTabPreservesPerTabResetTargets() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsReplaceEveryTab.docx");
            try {
                using var document = WordDocument.Create(filePath);
                document.AddParagraph("Replacement");
                var bodies = new List<string>();
                const string state = "{\"documentId\":\"doc-tabs\",\"title\":\"Tabbed\",\"revisionId\":\"revision-1\",\"tabs\":[{\"tabProperties\":{\"tabId\":\"tab-one\",\"title\":\"One\"},\"documentTab\":{\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":5,\"paragraph\":{}}]}}},{\"tabProperties\":{\"tabId\":\"tab-two\",\"title\":\"Two\"},\"documentTab\":{\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":6,\"paragraph\":{}}]}}}]}";
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "docs.googleapis.com") {
                        return CreateJsonResponse(state);
                    }
                    if (request.Method == HttpMethod.Post && request.RequestUri!.Host == "docs.googleapis.com") {
                        bodies.Add(await request.Content!.ReadAsStringAsync().ConfigureAwait(false));
                        return CreateJsonResponse($"{{\"writeControl\":{{\"requiredRevisionId\":\"revision-{bodies.Count + 1}\"}}}}");
                    }
                    if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "www.googleapis.com") {
                        return CreateJsonResponse("{\"id\":\"doc-tabs\",\"name\":\"Tabbed\",\"mimeType\":\"application/vnd.google-apps.document\"}");
                    }
                    return new HttpResponseMessage(HttpStatusCode.NotFound);
                }));
                var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Location = new GoogleDriveFileLocation { ExistingFileId = "doc-tabs" },
                    Tabs = new GoogleDocsTabOptions { Strategy = GoogleDocsTabStrategy.ReplaceEveryTab },
                    Replace = new GoogleDocsReplaceOptions { ExpectedRevisionId = "revision-1" },
                });

                string resetBody = Assert.IsType<string>(bodies.First());
                Assert.Contains("\"tabId\":\"tab-one\"", resetBody, StringComparison.Ordinal);
                Assert.Contains("\"tabId\":\"tab-two\"", resetBody, StringComparison.Ordinal);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_RejectsStaleRevisionBeforeMutation() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsStaleRevision.docx");
            try {
                using var document = WordDocument.Create(filePath);
                document.AddParagraph("Local edit");
                int mutationCount = 0;
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri.Contains("doc-stale?includeTabsContent=true", StringComparison.Ordinal)) {
                        return Task.FromResult(CreateJsonResponse(CreateTabbedDocumentStateJson("doc-stale", "remote-revision", "tab-a", "Remote edit")));
                    }
                    if (request.Method == HttpMethod.Post) mutationCount++;
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
                }));
                var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                GoogleWorkspaceConflictException exception = await Assert.ThrowsAsync<GoogleWorkspaceConflictException>(() =>
                    document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                        Location = new GoogleDriveFileLocation { ExistingFileId = "doc-stale" },
                        Replace = new GoogleDocsReplaceOptions { ExpectedRevisionId = "observed-revision" },
                    }));

                Assert.Equal("observed-revision", exception.ExpectedVersion);
                Assert.Equal("remote-revision", exception.ActualVersion);
                Assert.Equal(0, mutationCount);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_ChainsWriteControlAndSelectedTab() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsWriteControl.docx");
            try {
                using var document = WordDocument.Create(filePath);
                document.AddParagraph("Replacement");
                var bodies = new List<string>();
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    if (request.Method == HttpMethod.Get) {
                        return CreateJsonResponse(CreateTabbedDocumentStateJson("doc-write-control", "revision-1", "tab-a", "Old"));
                    }
                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri.EndsWith(":batchUpdate", StringComparison.Ordinal)) {
                        bodies.Add(await request.Content!.ReadAsStringAsync().ConfigureAwait(false));
                        string revision = bodies.Count == 1 ? "revision-2" : "revision-3";
                        return CreateJsonResponse($"{{\"writeControl\":{{\"requiredRevisionId\":\"{revision}\"}}}}");
                    }
                    return new HttpResponseMessage(HttpStatusCode.NotFound);
                }));
                var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                GoogleDocumentReference result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Location = new GoogleDriveFileLocation { ExistingFileId = "doc-write-control" },
                    Tabs = new GoogleDocsTabOptions { Strategy = GoogleDocsTabStrategy.SelectedTab, TabId = "tab-a" },
                    Replace = new GoogleDocsReplaceOptions { ExpectedRevisionId = "revision-1" },
                });

                Assert.Equal(2, bodies.Count);
                Assert.Contains("\"requiredRevisionId\":\"revision-1\"", bodies[0]);
                Assert.Contains("\"requiredRevisionId\":\"revision-2\"", bodies[1]);
                Assert.All(bodies, body => Assert.Contains("\"tabId\":\"tab-a\"", body));
                Assert.Equal("revision-3", result.RevisionId);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public async Task Test_GoogleDocsImporter_NativeFlattensTabsWithHeadings() {
            using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                if (request.RequestUri!.Host == "www.googleapis.com") {
                    return Task.FromResult(CreateJsonResponse("{\"id\":\"doc-import\",\"name\":\"Import\",\"mimeType\":\"application/vnd.google-apps.document\",\"version\":7,\"capabilities\":{\"canDownload\":true}}"));
                }
                if (request.RequestUri.Host == "docs.googleapis.com") {
                    const string json = "{\"documentId\":\"doc-import\",\"title\":\"Import\",\"revisionId\":\"revision-7\",\"tabs\":[{\"tabProperties\":{\"tabId\":\"one\",\"title\":\"Tab One\"},\"documentTab\":{\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":7,\"paragraph\":{\"elements\":[{\"textRun\":{\"content\":\"Alpha\\n\",\"textStyle\":{\"bold\":true}}}]}}]}}},{\"tabProperties\":{\"tabId\":\"two\",\"title\":\"Tab Two\"},\"documentTab\":{\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":6,\"paragraph\":{\"elements\":[{\"textRun\":{\"content\":\"Beta\\n\"}}]}}]}}}]}";
                    return Task.FromResult(CreateJsonResponse(json));
                }
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }));
            var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

            GoogleDocsImportResult imported = await new GoogleDocsImporter().ImportAsync("doc-import", session, new GoogleDocsImportOptions { Mode = GoogleDocsImportMode.Native });
            using (imported.Document) {
                WordDocumentSnapshot snapshot = imported.Document.CreateInspectionSnapshot();
                string[] text = snapshot.Sections.SelectMany(section => section.Elements).OfType<WordParagraphSnapshot>().Select(paragraph => paragraph.Text).ToArray();
                Assert.Contains("Tab One", text);
                Assert.Contains("Alpha", text);
                Assert.Contains("Tab Two", text);
                Assert.Contains("Beta", text);
                Assert.Equal("revision-7", imported.Source.RevisionId);
                Assert.Equal(7, imported.Source.DriveVersion);
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_CreatesUnanchoredDriveCommentThreads() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsComments.docx");
            try {
                using var document = WordDocument.Create(filePath);
                document.AddParagraph("Review target").AddComment("Alice", "A", "Please review");
                WordComment comment = WordComment.GetAllComments(document).Single();
                comment.AddReply("Bob", "B", "Reviewed");
                var driveBodies = new List<string>();
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-comments\",\"title\":\"Comments\",\"revisionId\":\"revision-1\"}");
                    }
                    if (request.Method == HttpMethod.Post && request.RequestUri!.Host == "docs.googleapis.com") return CreateJsonResponse("{}");
                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsolutePath.EndsWith("/comments", StringComparison.Ordinal)) {
                        driveBodies.Add(await request.Content!.ReadAsStringAsync().ConfigureAwait(false));
                        return CreateJsonResponse("{\"id\":\"comment-1\",\"content\":\"Alice: Please review\"}");
                    }
                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsolutePath.EndsWith("/replies", StringComparison.Ordinal)) {
                        driveBodies.Add(await request.Content!.ReadAsStringAsync().ConfigureAwait(false));
                        return CreateJsonResponse("{\"id\":\"reply-1\",\"content\":\"Bob: Reviewed\"}");
                    }
                    return new HttpResponseMessage(HttpStatusCode.NotFound);
                }));
                var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                GoogleDocumentReference result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions { Title = "Comments" });

                Assert.Equal(2, driveBodies.Count);
                Assert.Contains(driveBodies, body => body.Contains("Alice: Please review", StringComparison.Ordinal));
                Assert.Contains(driveBodies, body => body.Contains("Bob: Reviewed", StringComparison.Ordinal));
                Assert.Contains(result.Report.Notices, notice => notice.Code == "DOCS.COMMENT.UNANCHORED_CREATED");
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void Test_GoogleDocsDiffPlanner_DetectsIndependentConflict() {
            var checkpoint = new GoogleDocsSyncCheckpoint();
            checkpoint.ContentHashes["section/0/paragraph/0"] = "base";
            IReadOnlyList<GoogleDocsDiffItem> items = GoogleDocsDiffPlanner.Compare(
                new Dictionary<string, string> { ["section/0/paragraph/0"] = "local" },
                new Dictionary<string, string> { ["section/0/paragraph/0"] = "remote" },
                checkpoint);

            GoogleDocsDiffItem conflict = Assert.Single(items);
            Assert.Equal(GoogleDocsDiffKind.Conflict, conflict.Kind);
        }

        private static GoogleDocsApiTabResponse CreateTabState(string tabId, string title, int endIndex, string headerId, string footerId, string rangeName) {
            return new GoogleDocsApiTabResponse {
                Properties = new GoogleDocsApiTabPropertiesResponse { TabId = tabId, Title = title },
                DocumentTab = new GoogleDocsApiDocumentTabResponse {
                    Body = new GoogleDocsApiBodyResponse { Content = new List<GoogleDocsApiStructuralElementResponse> { new GoogleDocsApiStructuralElementResponse { StartIndex = 1, EndIndex = endIndex, Paragraph = new GoogleDocsApiParagraphElementResponse() } } },
                    Headers = new Dictionary<string, GoogleDocsApiHeaderFooterResponse> { [headerId] = new GoogleDocsApiHeaderFooterResponse() },
                    Footers = new Dictionary<string, GoogleDocsApiHeaderFooterResponse> { [footerId] = new GoogleDocsApiHeaderFooterResponse() },
                    NamedRanges = new Dictionary<string, GoogleDocsApiNamedRangesResponse> { [rangeName] = new GoogleDocsApiNamedRangesResponse() },
                },
            };
        }

        private static string CreateTabbedDocumentStateJson(string documentId, string revisionId, string tabId, string text) {
            return $"{{\"documentId\":\"{documentId}\",\"title\":\"State\",\"revisionId\":\"{revisionId}\",\"tabs\":[{{\"tabProperties\":{{\"tabId\":\"{tabId}\",\"title\":\"Tab\"}},\"documentTab\":{{\"body\":{{\"content\":[{{\"startIndex\":1,\"endIndex\":{text.Length + 2},\"paragraph\":{{\"elements\":[{{\"textRun\":{{\"content\":\"{text}\\n\"}}}}]}}}}]}}}}}}]}}";
        }
    }
}
