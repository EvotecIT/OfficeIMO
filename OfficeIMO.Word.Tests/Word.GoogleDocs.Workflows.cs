using OfficeIMO.GoogleWorkspace;
using OfficeIMO.Word;
using OfficeIMO.Word.GoogleDocs;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
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
                string[] contentBodies = bodies.Skip(1).ToArray();
                Assert.Equal(2, contentBodies.Length);
                Assert.Single(contentBodies, body => body.Contains("\"tabId\":\"tab-one\"", StringComparison.Ordinal));
                Assert.Single(contentBodies, body => body.Contains("\"tabId\":\"tab-two\"", StringComparison.Ordinal));
                Assert.All(contentBodies, body => Assert.Contains("Replacement", body, StringComparison.Ordinal));
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

        [Theory]
        [InlineData(GoogleDocsRevisionConflictMode.MergeAgainstTargetRevision)]
        [InlineData(GoogleDocsRevisionConflictMode.OverwriteLatest)]
        public async Task Test_GoogleDocsExporter_NonStrictRevisionModesAcceptStaleObservedRevision(GoogleDocsRevisionConflictMode mode) {
            string filePath = Path.Combine(_directoryWithFiles, $"GoogleDocsNonStrictRevision-{mode}.docx");
            try {
                using var document = WordDocument.Create(filePath);
                document.AddParagraph("Local edit");
                var bodies = new List<string>();
                int documentReads = 0;
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "docs.googleapis.com") {
                        documentReads++;
                        string revision = documentReads == 1 ? "remote-revision" : "revision-readback";
                        return CreateJsonResponse(CreateTabbedDocumentStateJson("doc-nonstrict", revision, "tab-a", "Remote edit"));
                    }
                    if (request.Method == HttpMethod.Post && request.RequestUri!.Host == "docs.googleapis.com") {
                        bodies.Add(await request.Content!.ReadAsStringAsync().ConfigureAwait(false));
                        return mode == GoogleDocsRevisionConflictMode.OverwriteLatest
                            ? CreateJsonResponse("{}")
                            : CreateJsonResponse("{\"writeControl\":{\"requiredRevisionId\":\"revision-new\"}}");
                    }
                    return new HttpResponseMessage(HttpStatusCode.NotFound);
                }));
                var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                GoogleDocumentReference result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Location = new GoogleDriveFileLocation { ExistingFileId = "doc-nonstrict" },
                    Replace = new GoogleDocsReplaceOptions {
                        ConflictMode = mode,
                        ExpectedRevisionId = "observed-revision",
                    },
                });

                Assert.NotEmpty(bodies);
                if (mode == GoogleDocsRevisionConflictMode.MergeAgainstTargetRevision) {
                    Assert.Contains("\"targetRevisionId\":\"observed-revision\"", bodies[0], StringComparison.Ordinal);
                } else {
                    Assert.All(bodies, body => Assert.DoesNotContain("writeControl", body, StringComparison.Ordinal));
                }
                Assert.Equal(
                    mode == GoogleDocsRevisionConflictMode.OverwriteLatest ? "revision-readback" : "revision-new",
                    result.RevisionId);
                Assert.Equal(mode == GoogleDocsRevisionConflictMode.OverwriteLatest ? 2 : 1, documentReads);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_ReturnsPostWriteDriveCheckpointMetadata() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsDriveCheckpoint.docx");
            try {
                using var document = WordDocument.Create(filePath);
                document.AddParagraph("Replacement");
                DateTimeOffset modified = DateTimeOffset.Parse("2026-07-15T20:00:00Z", System.Globalization.CultureInfo.InvariantCulture);
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                    if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "docs.googleapis.com") {
                        return Task.FromResult(CreateJsonResponse(CreateTabbedDocumentStateJson("doc-checkpoint", "revision-1", "tab-a", "Remote")));
                    }
                    if (request.Method == HttpMethod.Post && request.RequestUri!.Host == "docs.googleapis.com") {
                        return Task.FromResult(CreateJsonResponse("{\"writeControl\":{\"requiredRevisionId\":\"revision-2\"}}"));
                    }
                    if (request.Method == HttpMethod.Get
                        && request.RequestUri!.AbsolutePath == "/drive/v3/files/doc-checkpoint") {
                        return Task.FromResult(CreateJsonResponse("{\"id\":\"doc-checkpoint\",\"name\":\"Checkpoint\",\"mimeType\":\"application/vnd.google-apps.document\",\"webViewLink\":\"https://docs.google.com/document/d/doc-checkpoint/edit\",\"version\":42,\"modifiedTime\":\"2026-07-15T20:00:00Z\"}"));
                    }
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
                }));
                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                GoogleDocumentReference result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Location = new GoogleDriveFileLocation { ExistingFileId = "doc-checkpoint" },
                    Replace = new GoogleDocsReplaceOptions { ExpectedRevisionId = "revision-1" },
                });

                Assert.Equal(42, result.DriveVersion);
                Assert.Equal(modified, result.ModifiedTime);
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
        public async Task Test_GoogleDocsExporter_UsesCreatedTabIdInsteadOfCallerTabId() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsCreatedTab.docx");
            try {
                using var document = WordDocument.Create(filePath);
                document.AddParagraph("Created tab content");
                var bodies = new List<string>();
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    if (request.Method == HttpMethod.Post
                        && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-created-tab\",\"title\":\"Created\",\"revisionId\":\"revision-1\",\"tabs\":[{\"tabProperties\":{\"tabId\":\"api-created-tab\",\"title\":\"Tab 1\"},\"documentTab\":{\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":1,\"paragraph\":{}}]}}}]}");
                    }
                    if (request.Method == HttpMethod.Post
                        && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-created-tab:batchUpdate") {
                        bodies.Add(await request.Content!.ReadAsStringAsync().ConfigureAwait(false));
                        return CreateJsonResponse("{\"writeControl\":{\"requiredRevisionId\":\"revision-2\"}}");
                    }
                    return new HttpResponseMessage(HttpStatusCode.NotFound);
                }));
                var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                GoogleDocumentReference result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Tabs = new GoogleDocsTabOptions { Strategy = GoogleDocsTabStrategy.SelectedTab, TabId = "caller-supplied-tab" },
                });

                string body = Assert.Single(bodies);
                Assert.Contains("\"tabId\":\"api-created-tab\"", body, StringComparison.Ordinal);
                Assert.DoesNotContain("caller-supplied-tab", body, StringComparison.Ordinal);
                Assert.Equal("revision-2", result.RevisionId);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_ScopesFirstSectionHeaderToSelectedTab() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsSelectedTabHeader.docx");
            try {
                using var document = WordDocument.Create(filePath);
                document.AddParagraph("Replacement");
                document.AddHeadersAndFooters();
                document.Sections[0].Header.Default!.AddParagraph("Selected tab header");
                var bodies = new List<string>();
                const string state = "{\"documentId\":\"doc-tab-header\",\"title\":\"Tabbed\",\"revisionId\":\"revision-1\",\"tabs\":[{\"tabProperties\":{\"tabId\":\"tab-a\",\"title\":\"One\"},\"documentTab\":{\"body\":{\"content\":[{\"startIndex\":0,\"endIndex\":1,\"sectionBreak\":{}},{\"startIndex\":1,\"endIndex\":5,\"paragraph\":{}}]}}},{\"tabProperties\":{\"tabId\":\"tab-b\",\"title\":\"Two\"},\"documentTab\":{\"body\":{\"content\":[{\"startIndex\":0,\"endIndex\":1,\"sectionBreak\":{}},{\"startIndex\":1,\"endIndex\":5,\"paragraph\":{}}]}}}]}";
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "docs.googleapis.com") {
                        return CreateJsonResponse(state);
                    }
                    if (request.Method == HttpMethod.Post && request.RequestUri!.Host == "docs.googleapis.com") {
                        string body = await request.Content!.ReadAsStringAsync().ConfigureAwait(false);
                        bodies.Add(body);
                        return body.Contains("\"createHeader\"", StringComparison.Ordinal)
                            ? CreateJsonResponse("{\"replies\":[{\"createHeader\":{\"headerId\":\"header-selected\"}}],\"writeControl\":{\"requiredRevisionId\":\"revision-2\"}}")
                            : CreateJsonResponse("{\"writeControl\":{\"requiredRevisionId\":\"revision-2\"}}");
                    }
                    if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "www.googleapis.com") {
                        return CreateJsonResponse("{\"id\":\"doc-tab-header\",\"name\":\"Tabbed\",\"mimeType\":\"application/vnd.google-apps.document\"}");
                    }
                    return new HttpResponseMessage(HttpStatusCode.NotFound);
                }));
                var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Location = new GoogleDriveFileLocation { ExistingFileId = "doc-tab-header" },
                    Tabs = new GoogleDocsTabOptions { Strategy = GoogleDocsTabStrategy.SelectedTab, TabId = "tab-b" },
                    Replace = new GoogleDocsReplaceOptions { ExpectedRevisionId = "revision-1" },
                });

                string createHeaderBody = Assert.Single(bodies, body => body.Contains("\"createHeader\"", StringComparison.Ordinal));
                using JsonDocument payload = JsonDocument.Parse(createHeaderBody);
                JsonElement createHeader = payload.RootElement.GetProperty("requests").EnumerateArray()
                    .Single(request => request.TryGetProperty("createHeader", out _))
                    .GetProperty("createHeader");
                JsonElement location = createHeader.GetProperty("sectionBreakLocation");
                Assert.Equal(0, location.GetProperty("index").GetInt32());
                Assert.Equal("tab-b", location.GetProperty("tabId").GetString());
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_AttachesLaterSectionHeaderToPrecedingBreak() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsLaterSectionHeader.docx");
            try {
                using var document = WordDocument.Create(filePath);
                document.AddParagraph("First section");
                WordSection secondSection = document.AddSection();
                secondSection.AddParagraph("Second section");
                secondSection.AddHeadersAndFooters();
                secondSection.Header.Default!.AddParagraph("Second section header");

                var bodies = new List<string>();
                const string state = "{\"documentId\":\"doc-later-header\",\"title\":\"Later header\",\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":10,\"paragraph\":{}},{\"startIndex\":10,\"endIndex\":11,\"sectionBreak\":{}},{\"startIndex\":11,\"endIndex\":30,\"paragraph\":{}}]}}";
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(async request => {
                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents") {
                        return CreateJsonResponse("{\"documentId\":\"doc-later-header\",\"title\":\"Later header\"}");
                    }
                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-later-header?includeTabsContent=true") {
                        return CreateJsonResponse(state);
                    }
                    if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri == "https://docs.googleapis.com/v1/documents/doc-later-header:batchUpdate") {
                        string body = await request.Content!.ReadAsStringAsync().ConfigureAwait(false);
                        bodies.Add(body);
                        return body.Contains("\"createHeader\"", StringComparison.Ordinal)
                            ? CreateJsonResponse("{\"replies\":[{\"createHeader\":{\"headerId\":\"header-later\"}}]}")
                            : CreateJsonResponse("{}");
                    }
                    return new HttpResponseMessage(HttpStatusCode.NotFound);
                }));
                var session = new GoogleWorkspaceSession(
                    new FakeGoogleWorkspaceCredentialSource(),
                    new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions { Title = "Later header" });

                string createHeaderBody = Assert.Single(bodies, body => body.Contains("\"createHeader\"", StringComparison.Ordinal));
                using JsonDocument payload = JsonDocument.Parse(createHeaderBody);
                JsonElement location = payload.RootElement.GetProperty("requests").EnumerateArray()
                    .Single(request => request.TryGetProperty("createHeader", out _))
                    .GetProperty("createHeader")
                    .GetProperty("sectionBreakLocation");
                Assert.Equal(10, location.GetProperty("index").GetInt32());
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public async Task Test_GoogleDocsImporter_NativeFlattensTabsWhenDownloadIsAllowed() {
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
                Assert.Equal("Import", imported.Document.BuiltinDocumentProperties.Title);
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
        public async Task Test_GoogleDocsImporter_NativeRejectsFilesThatCannotBeDownloaded() {
            int nativeReads = 0;
            using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                if (request.RequestUri!.Host == "www.googleapis.com") {
                    return Task.FromResult(CreateJsonResponse("{\"id\":\"doc-blocked\",\"mimeType\":\"application/vnd.google-apps.document\",\"capabilities\":{\"canDownload\":false}}"));
                }
                nativeReads++;
                return Task.FromResult(CreateJsonResponse("{\"documentId\":\"doc-blocked\"}"));
            }));
            var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

            InvalidOperationException exception = await Assert.ThrowsAsync<InvalidOperationException>(() =>
                new GoogleDocsImporter().ImportAsync("doc-blocked", session, new GoogleDocsImportOptions { Mode = GoogleDocsImportMode.Native }));

            Assert.Contains("cannot be exported", exception.Message, StringComparison.Ordinal);
            Assert.Equal(0, nativeReads);
        }

        [Fact]
        public async Task Test_GoogleDocsImporter_NativeEnforcesResponseAndModelBudgets() {
            const string nativeJson = "{\"documentId\":\"doc-large\",\"body\":{\"content\":[{\"paragraph\":{\"elements\":[{\"textRun\":{\"content\":\"0123456789\"}}]}}]}}";
            using var httpClient = new HttpClient(new FakeHttpMessageHandler(request =>
                request.RequestUri!.Host == "www.googleapis.com"
                    ? Task.FromResult(CreateJsonResponse("{\"id\":\"doc-large\",\"mimeType\":\"application/vnd.google-apps.document\",\"capabilities\":{\"canDownload\":true}}"))
                    : Task.FromResult(CreateJsonResponse(nativeJson))));
            var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

            await Assert.ThrowsAsync<InvalidDataException>(() =>
                new GoogleDocsImporter().ImportAsync("doc-large", session, new GoogleDocsImportOptions {
                    Mode = GoogleDocsImportMode.Native,
                    MaxResponseBytes = 32,
                }));
            await Assert.ThrowsAsync<InvalidDataException>(() =>
                new GoogleDocsImporter().ImportAsync("doc-large", session, new GoogleDocsImportOptions {
                    Mode = GoogleDocsImportMode.Native,
                    MaxTextCharacters = 4,
                }));
        }

        [Fact]
        public async Task Test_GoogleDocsImporter_NativeBoundsSparseTableRows() {
            const string nativeJson = "{\"documentId\":\"doc-sparse-table\",\"body\":{\"content\":[{\"table\":{\"tableRows\":[{\"tableCells\":[]},{\"tableCells\":[]}]}}]}}";
            using var httpClient = new HttpClient(new FakeHttpMessageHandler(request =>
                request.RequestUri!.Host == "www.googleapis.com"
                    ? Task.FromResult(CreateJsonResponse("{\"id\":\"doc-sparse-table\",\"mimeType\":\"application/vnd.google-apps.document\",\"capabilities\":{\"canDownload\":true}}"))
                    : Task.FromResult(CreateJsonResponse(nativeJson))));
            var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(),
                new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

            await Assert.ThrowsAsync<InvalidDataException>(() =>
                new GoogleDocsImporter().ImportAsync("doc-sparse-table", session, new GoogleDocsImportOptions {
                    Mode = GoogleDocsImportMode.Native,
                    MaxTableCells = 1,
                }));
        }

        [Fact]
        public async Task Test_GoogleDocsImporter_NativeChargesRectangularTableProjection() {
            const string nativeJson = "{\"documentId\":\"doc-ragged-table\",\"body\":{\"content\":[{\"table\":{\"tableRows\":[{\"tableCells\":[{}]},{\"tableCells\":[{},{},{}]}]}}]}}";
            using var httpClient = new HttpClient(new FakeHttpMessageHandler(request =>
                request.RequestUri!.Host == "www.googleapis.com"
                    ? Task.FromResult(CreateJsonResponse("{\"id\":\"doc-ragged-table\",\"mimeType\":\"application/vnd.google-apps.document\",\"capabilities\":{\"canDownload\":true}}"))
                    : Task.FromResult(CreateJsonResponse(nativeJson))));
            var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(),
                new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

            await Assert.ThrowsAsync<InvalidDataException>(() =>
                new GoogleDocsImporter().ImportAsync("doc-ragged-table", session, new GoogleDocsImportOptions {
                    Mode = GoogleDocsImportMode.Native,
                    MaxTableCells = 4,
                }));
        }

        [Fact]
        public async Task Test_GoogleDocsDiffPlanner_ReportsDriveVersionChanges() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsDriveVersionDiff.docx");
            try {
                using var source = WordDocument.Create(filePath);
                source.AddParagraph("Same");
                GoogleDocsSyncCheckpoint checkpoint = GoogleDocsDiffPlanner.CreateCheckpoint(source, revisionId: "revision-1", driveVersion: 7);
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                    if (request.RequestUri!.Host == "www.googleapis.com") {
                        return Task.FromResult(CreateJsonResponse("{\"id\":\"doc-diff\",\"name\":\"Diff\",\"mimeType\":\"application/vnd.google-apps.document\",\"version\":8,\"capabilities\":{\"canDownload\":true}}"));
                    }
                    const string docs = "{\"documentId\":\"doc-diff\",\"title\":\"Diff\",\"revisionId\":\"revision-1\",\"body\":{\"content\":[{\"startIndex\":1,\"endIndex\":6,\"paragraph\":{\"elements\":[{\"textRun\":{\"content\":\"Same\\n\"}}]}}]}}";
                    return Task.FromResult(CreateJsonResponse(docs));
                }));
                var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                GoogleDocsDiffPlan plan = await GoogleDocsDiffPlanner.BuildAsync(source, "doc-diff", session, checkpoint);

                Assert.Contains(plan.Items, item => item.Kind == GoogleDocsDiffKind.RemoteChange && item.Path == "document/driveVersion");
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void Test_GoogleDocsCheckpoint_HashesExportedLayoutImagesCommentsAndTableRuns() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsSemanticCheckpoint.docx");
            try {
                using var document = WordDocument.Create(filePath);
                WordParagraph body = document.AddParagraph("Body");
                WordParagraph imageParagraph = document.AddParagraph("Image ");
                byte[] png = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
                using (var imageStream = new MemoryStream(png)) {
                    imageParagraph.AddImage(imageStream, "pixel.png", 10, 10);
                }
                WordImage image = Assert.Single(document.Images);

                WordTable table = document.AddTable(1, 1, WordTableStyle.TableGrid);
                WordParagraph tableParagraph = table.Rows[0].Cells[0].Paragraphs[0];
                tableParagraph.Text = "Cell";

                body.AddComment("Alice", "A", "Please review");
                WordComment comment = Assert.Single(document.Comments, candidate => candidate.ParentComment == null);
                WordComment reply = comment.AddReply("Bob", "B", "Reviewed");

                GoogleDocsSyncCheckpoint baseline = GoogleDocsDiffPlanner.CreateCheckpoint(document);
                const string bodyPath = "section/0/paragraph/0";
                const string imagePath = "section/0/paragraph/1";
                const string tableCellPath = "section/0/table/2/cell/0:0";
                const string replyPath = "comment/0/reply/0";

                body.PageBreakBefore = true;
                body.IndentationBeforePoints = 18;
                GoogleDocsSyncCheckpoint layoutChanged = GoogleDocsDiffPlanner.CreateCheckpoint(document);
                Assert.NotEqual(baseline.ContentHashes[bodyPath], layoutChanged.ContentHashes[bodyPath]);

                image.Width = 20;
                GoogleDocsSyncCheckpoint imageChanged = GoogleDocsDiffPlanner.CreateCheckpoint(document);
                Assert.NotEqual(layoutChanged.ContentHashes[imagePath], imageChanged.ContentHashes[imagePath]);

                reply.Text = "Reviewed with changes";
                GoogleDocsSyncCheckpoint commentChanged = GoogleDocsDiffPlanner.CreateCheckpoint(document);
                Assert.NotEqual(imageChanged.ContentHashes[replyPath], commentChanged.ContentHashes[replyPath]);

                tableParagraph.AddFormattedText(" Bold", bold: true);
                tableParagraph.AddHyperLink(" Link", new Uri("https://example.test/"));
                GoogleDocsSyncCheckpoint tableChanged = GoogleDocsDiffPlanner.CreateCheckpoint(document);
                Assert.NotEqual(commentChanged.ContentHashes[tableCellPath], tableChanged.ContentHashes[tableCellPath]);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void Test_GoogleDocsCheckpoint_AssignsMalformedDuplicateCommentThreadRepliesOnce() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsDuplicateCommentThread.docx");
            try {
                using var document = WordDocument.Create(filePath);
                document.AddParagraph("First").AddComment("Alice", "A", "First root");
                WordComment first = WordComment.GetAllComments(document)
                    .Single(comment => comment.Text == "First root");
                first.AddReply("Bob", "B", "Only reply");
                document.AddParagraph("Second").AddComment("Carol", "C", "Second root");
                WordComment duplicate = WordComment.GetAllComments(document)
                    .Single(comment => comment.Text == "Second root");
                var commentsEx = document._wordprocessingDocument.MainDocumentPart!
                    .WordprocessingCommentsExPart!.CommentsEx!;
                DocumentFormat.OpenXml.Office2013.Word.CommentEx duplicateMetadata = commentsEx
                    .Elements<DocumentFormat.OpenXml.Office2013.Word.CommentEx>()
                    .Single(item => string.Equals(item.ParaId?.Value, duplicate.ParaId,
                        StringComparison.Ordinal));
                duplicateMetadata.ParaId = first.ParaId;

                GoogleDocsSyncCheckpoint checkpoint = GoogleDocsDiffPlanner.CreateCheckpoint(document);

                Assert.Equal(2, checkpoint.ContentHashes.Keys.Count(path =>
                    path.StartsWith("comment/", StringComparison.Ordinal)
                    && !path.Contains("/reply/", StringComparison.Ordinal)));
                Assert.Single(checkpoint.ContentHashes.Keys, path =>
                    path.Contains("/reply/", StringComparison.Ordinal));
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Theory]
        [InlineData(GoogleDocsSuggestionsMode.Default, "DEFAULT_FOR_CURRENT_ACCESS")]
        [InlineData(GoogleDocsSuggestionsMode.Accepted, "PREVIEW_SUGGESTIONS_ACCEPTED")]
        [InlineData(GoogleDocsSuggestionsMode.Inline, "SUGGESTIONS_INLINE")]
        [InlineData(GoogleDocsSuggestionsMode.Rejected, "PREVIEW_WITHOUT_SUGGESTIONS")]
        public async Task Test_GoogleDocsImporter_MapsSuggestionModesToNativeApiValues(
            GoogleDocsSuggestionsMode mode,
            string expectedApiValue) {
            Uri? docsUri = null;
            using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                if (request.RequestUri!.Host == "www.googleapis.com") {
                    return Task.FromResult(CreateJsonResponse("{\"id\":\"doc-suggestions\",\"name\":\"Suggestions\",\"mimeType\":\"application/vnd.google-apps.document\"}"));
                }

                if (request.RequestUri.Host == "docs.googleapis.com") {
                    docsUri = request.RequestUri;
                    return Task.FromResult(CreateJsonResponse("{\"documentId\":\"doc-suggestions\",\"title\":\"Suggestions\",\"body\":{\"content\":[]}}"));
                }

                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }));
            var session = new GoogleWorkspaceSession(
                new FakeGoogleWorkspaceCredentialSource(),
                new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

            GoogleDocsImportResult imported = await new GoogleDocsImporter().ImportAsync(
                "doc-suggestions",
                session,
                new GoogleDocsImportOptions { Mode = GoogleDocsImportMode.Native, Suggestions = mode });
            using (imported.Document) {
                Assert.NotNull(docsUri);
                Assert.Contains("suggestionsViewMode=" + expectedApiValue, docsUri!.Query, StringComparison.Ordinal);
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
        public async Task Test_GoogleDocsExporter_ReusesMatchingDriveCommentsOnReplacement() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsCommentReplacement.docx");
            try {
                using var document = WordDocument.Create(filePath);
                document.AddParagraph("Review target").AddComment("Alice", "A", "Please review");
                WordComment comment = WordComment.GetAllComments(document).Single();
                comment.AddReply("Bob", "B", "Reviewed");
                int createdCommentItems = 0;
                int deletedCommentItems = 0;
                int commentListReads = 0;
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                    if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "docs.googleapis.com") {
                        return Task.FromResult(CreateJsonResponse(CreateTabbedDocumentStateJson("doc-comments", "revision-1", "tab-a", "Remote")));
                    }
                    if (request.Method == HttpMethod.Post && request.RequestUri!.Host == "docs.googleapis.com") {
                        return Task.FromResult(CreateJsonResponse("{\"writeControl\":{\"requiredRevisionId\":\"revision-2\"}}"));
                    }
                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsolutePath.EndsWith("/comments", StringComparison.Ordinal)) {
                        commentListReads++;
                        return Task.FromResult(request.RequestUri.Query.Contains("pageToken=next", StringComparison.Ordinal)
                            ? CreateJsonResponse("{\"comments\":[{\"id\":\"comment-1\",\"content\":\"Alice: Please review\",\"replies\":[{\"id\":\"reply-1\",\"content\":\"Bob: Reviewed\"},{\"id\":\"reply-stale\",\"content\":\"Old reply\"}]}]}")
                            : CreateJsonResponse("{\"nextPageToken\":\"next\",\"comments\":[]}"));
                    }
                    if (request.Method == HttpMethod.Post
                        && (request.RequestUri!.AbsolutePath.EndsWith("/comments", StringComparison.Ordinal)
                            || request.RequestUri.AbsolutePath.EndsWith("/replies", StringComparison.Ordinal))) {
                        createdCommentItems++;
                        return Task.FromResult(CreateJsonResponse("{}"));
                    }
                    if (request.Method == HttpMethod.Delete && request.RequestUri!.AbsolutePath.EndsWith("/replies/reply-stale", StringComparison.Ordinal)) {
                        deletedCommentItems++;
                        return Task.FromResult(CreateJsonResponse("{}"));
                    }
                    if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "www.googleapis.com") {
                        return Task.FromResult(CreateJsonResponse("{\"id\":\"doc-comments\",\"name\":\"Comments\",\"mimeType\":\"application/vnd.google-apps.document\",\"version\":2}"));
                    }
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
                }));
                var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                GoogleDocumentReference result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Location = new GoogleDriveFileLocation { ExistingFileId = "doc-comments" },
                    Replace = new GoogleDocsReplaceOptions {
                        ConflictMode = GoogleDocsRevisionConflictMode.OverwriteLatest,
                    },
                });

                Assert.Equal(2, commentListReads);
                Assert.Equal(0, createdCommentItems);
                Assert.Equal(0, deletedCommentItems);
                Assert.Contains(result.Report.Notices, notice => notice.Code == "DOCS.COMMENT.UNANCHORED_REUSED" && notice.Count == 2);
                Assert.DoesNotContain(result.Report.Notices, notice => notice.Code == "DOCS.COMMENT.UNANCHORED_DELETED");
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public async Task Test_GoogleDocsExporter_PreservesUnrelatedDriveCommentsOnReplacement() {
            string filePath = Path.Combine(_directoryWithFiles, "GoogleDocsCommentRemoval.docx");
            try {
                using var document = WordDocument.Create(filePath);
                document.AddParagraph("Replacement without comments");
                var deletedUris = new List<string>();
                using var httpClient = new HttpClient(new FakeHttpMessageHandler(request => {
                    if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "docs.googleapis.com") {
                        return Task.FromResult(CreateJsonResponse(CreateTabbedDocumentStateJson("doc-comments", "revision-1", "tab-a", "Remote")));
                    }
                    if (request.Method == HttpMethod.Post && request.RequestUri!.Host == "docs.googleapis.com") {
                        return Task.FromResult(CreateJsonResponse("{\"writeControl\":{\"requiredRevisionId\":\"revision-2\"}}"));
                    }
                    if (request.Method == HttpMethod.Get && request.RequestUri!.AbsolutePath.EndsWith("/comments", StringComparison.Ordinal)) {
                        return Task.FromResult(CreateJsonResponse("{\"comments\":[{\"id\":\"stale-comment\",\"content\":\"Old review\"}]}"));
                    }
                    if (request.Method == HttpMethod.Delete && request.RequestUri!.AbsolutePath.EndsWith("/comments/stale-comment", StringComparison.Ordinal)) {
                        deletedUris.Add(request.RequestUri.AbsoluteUri);
                        return Task.FromResult(CreateJsonResponse("{}"));
                    }
                    if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "www.googleapis.com") {
                        return Task.FromResult(CreateJsonResponse("{\"id\":\"doc-comments\",\"name\":\"Comments\",\"mimeType\":\"application/vnd.google-apps.document\",\"version\":2}"));
                    }
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
                }));
                var session = new GoogleWorkspaceSession(new FakeGoogleWorkspaceCredentialSource(), new GoogleWorkspaceSessionOptions { HttpClient = httpClient });

                GoogleDocumentReference result = await document.ExportToGoogleDocsAsync(session, new GoogleDocsSaveOptions {
                    Location = new GoogleDriveFileLocation { ExistingFileId = "doc-comments" },
                    Replace = new GoogleDocsReplaceOptions { ConflictMode = GoogleDocsRevisionConflictMode.OverwriteLatest },
                });

                Assert.Empty(deletedUris);
                Assert.DoesNotContain(result.Report.Notices, notice => notice.Code == "DOCS.COMMENT.UNANCHORED_DELETED");
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
