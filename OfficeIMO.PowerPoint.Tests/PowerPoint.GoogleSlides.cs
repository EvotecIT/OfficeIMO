using OfficeIMO.GoogleWorkspace;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.GoogleSlides;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class GoogleSlidesTests {
        [Fact]
        public void BatchCompiler_MapsEditableCoreAndDeterministicIds() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.Hidden = true;
            slide.AddTextBoxPoints("Hello Slides", 20, 30, 300, 80).Paragraphs[0].Runs[0].Bold = true;
            PowerPointTextBox hiddenText = slide.AddTextBoxPoints("Hidden shape", 20, 110, 300, 40);
            hiddenText.Hidden = true;
            PowerPointTable table = slide.AddTablePoints(2, 2, 40, 140, 400, 160);
            table.RowItems[0].Cells[0].Text = "A1";
            slide.Notes.Text = "Speaker note";

            GoogleSlidesBatch batch = presentation.BuildGoogleSlidesBatch(new GoogleSlidesSaveOptions { Title = "Deck" });

            Assert.Single(batch.Slides);
            Assert.Equal("officeimo_slide_0001_0001", batch.Slides[0].ObjectId);
            Assert.True(batch.Slides[0].IsSkipped);
            Assert.Contains(batch.Slides[0].Elements, element => element is GoogleSlidesTextBox text && text.Text == "Hello Slides" && text.Bold);
            Assert.DoesNotContain(batch.Slides[0].Elements, element => element is GoogleSlidesTextBox text && text.Text == "Hidden shape");
            Assert.Contains(batch.Slides[0].Elements, element => element is GoogleSlidesTable grid && grid.Cells[0][0] == "A1");
            Assert.Equal("Speaker note", batch.Slides[0].SpeakerNotes);
            Assert.Equal(1, batch.Plan.NativeTextBoxCount);
            Assert.Equal(1, batch.Plan.NativeTableCount);
        }

        [Fact]
        public void BatchCompiler_DoesNotSendUnsupportedNativeImageFormats() {
            string svgPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".svg");
            try {
                File.WriteAllText(svgPath, "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"2\" height=\"2\"><rect width=\"2\" height=\"2\"/></svg>");
                using PowerPointPresentation presentation = PowerPointPresentation.Create();
                presentation.AddSlide().AddPicture(svgPath, left: 0, top: 0, width: 1000, height: 1000);

                GoogleSlidesBatch batch = presentation.BuildGoogleSlidesBatch(new GoogleSlidesSaveOptions {
                    ComplexSlides = GoogleSlidesComplexSlideMode.PreferNativeAndReport,
                });

                Assert.Empty(Assert.Single(batch.Slides).Elements.OfType<GoogleSlidesImage>());
                Assert.Contains(batch.Plan.Report.Notices, notice => notice.Code == "SLIDES.IMAGE.FORMAT_SKIPPED");
            } finally {
                if (File.Exists(svgPath)) File.Delete(svgPath);
            }
        }

        [Fact]
        public async Task Exporter_CreatesAndReplacesInitialSlideWithRevisionGuard() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide authoredSlide = presentation.AddSlide();
            authoredSlide.Hidden = true;
            authoredSlide.AddTextBoxPoints("Hello Slides", 20, 30, 300, 80);
            var batchBodies = new List<string>();
            using var httpClient = new HttpClient(new DelegateHandler(async request => {
                string uri = request.RequestUri!.AbsoluteUri;
                if (request.Method == HttpMethod.Post && uri == "https://slides.googleapis.com/v1/presentations") return Json("{\"presentationId\":\"deck-1\",\"title\":\"Deck\"}");
                if (request.Method == HttpMethod.Get && uri == "https://slides.googleapis.com/v1/presentations/deck-1") return Json("{\"presentationId\":\"deck-1\",\"title\":\"Deck\",\"revisionId\":\"revision-1\",\"slides\":[{\"objectId\":\"initial-slide\"}]}");
                if (request.Method == HttpMethod.Post && uri.EndsWith(":batchUpdate", StringComparison.Ordinal)) {
                    batchBodies.Add(await request.Content!.ReadAsStringAsync().ConfigureAwait(false));
                    return Json("{\"presentationId\":\"deck-1\",\"writeControl\":{\"requiredRevisionId\":\"revision-2\"}}");
                }
                if (request.Method == HttpMethod.Get && request.RequestUri.Host == "www.googleapis.com") return Json("{\"id\":\"deck-1\",\"name\":\"Deck\",\"mimeType\":\"application/vnd.google-apps.presentation\",\"version\":2,\"webViewLink\":\"https://docs.google.com/presentation/d/deck-1/edit\"}");
                return new HttpResponseMessage(HttpStatusCode.NotFound) { Content = new StringContent("unexpected " + uri) };
            }));
            var session = Session(httpClient);

            GooglePresentationReference result = await presentation.ExportToGoogleSlidesAsync(session, new GoogleSlidesSaveOptions { Title = "Deck" });

            string body = Assert.Single(batchBodies);
            Assert.Contains("\"deleteObject\":{\"objectId\":\"initial-slide\"}", body);
            Assert.Contains("\"createSlide\"", body);
            Assert.Contains("\"slideProperties\":{\"isSkipped\":true},\"fields\":\"isSkipped\"", body);
            Assert.Contains("\"createShape\"", body);
            Assert.Contains("Hello Slides", body);
            Assert.Contains("\"requiredRevisionId\":\"revision-1\"", body);
            Assert.Equal("revision-2", result.RevisionId);
            Assert.Equal(2, result.DriveVersion);
        }

        [Fact]
        public async Task Exporter_ScalesAndCentersElementsToTargetPageSize() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.SlideSize.SetSizePoints(960, 540);
            presentation.AddSlide().AddTextBoxPoints("Edge", 800, 400, 100, 100);
            string? batchBody = null;
            using var httpClient = new HttpClient(new DelegateHandler(async request => {
                string uri = request.RequestUri!.AbsoluteUri;
                if (request.Method == HttpMethod.Post && uri == "https://slides.googleapis.com/v1/presentations") return Json("{\"presentationId\":\"deck-scaled\"}");
                if (request.Method == HttpMethod.Get && uri == "https://slides.googleapis.com/v1/presentations/deck-scaled") {
                    return Json("{\"presentationId\":\"deck-scaled\",\"revisionId\":\"revision-1\",\"pageSize\":{\"width\":{\"magnitude\":720,\"unit\":\"PT\"},\"height\":{\"magnitude\":405,\"unit\":\"PT\"}},\"slides\":[{\"objectId\":\"initial-slide\"}]}");
                }
                if (request.Method == HttpMethod.Post && uri.EndsWith(":batchUpdate", StringComparison.Ordinal)) {
                    batchBody = await request.Content!.ReadAsStringAsync().ConfigureAwait(false);
                    return Json("{\"writeControl\":{\"requiredRevisionId\":\"revision-2\"}}");
                }
                if (request.Method == HttpMethod.Get && request.RequestUri.Host == "www.googleapis.com") return Json("{\"id\":\"deck-scaled\",\"mimeType\":\"application/vnd.google-apps.presentation\"}");
                return new HttpResponseMessage(HttpStatusCode.NotFound);
            }));

            GooglePresentationReference result = await presentation.ExportToGoogleSlidesAsync(Session(httpClient));

            using JsonDocument payload = JsonDocument.Parse(Assert.IsType<string>(batchBody));
            JsonElement shapeRequest = payload.RootElement.GetProperty("requests").EnumerateArray()
                .Single(request => request.TryGetProperty("createShape", out _));
            JsonElement properties = shapeRequest.GetProperty("createShape").GetProperty("elementProperties");
            Assert.Equal(75, properties.GetProperty("size").GetProperty("width").GetProperty("magnitude").GetDouble());
            Assert.Equal(75, properties.GetProperty("size").GetProperty("height").GetProperty("magnitude").GetDouble());
            Assert.Equal(600, properties.GetProperty("transform").GetProperty("translateX").GetDouble());
            Assert.Equal(300, properties.GetProperty("transform").GetProperty("translateY").GetDouble());
            Assert.Contains(result.Report.Notices, notice => notice.Code == "SLIDES.PAGE_SIZE.SCALED");
        }

        [Fact]
        public async Task Exporter_WritesSpeakerNotesAfterSlidesExist() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddTextBox("Slide body");
            slide.Notes.Text = "Presenter-only context";
            int presentationReads = 0;
            var batchBodies = new List<string>();
            using var httpClient = new HttpClient(new DelegateHandler(async request => {
                string uri = request.RequestUri!.AbsoluteUri;
                if (request.Method == HttpMethod.Post && uri == "https://slides.googleapis.com/v1/presentations") return Json("{\"presentationId\":\"deck-notes\"}");
                if (request.Method == HttpMethod.Get && uri == "https://slides.googleapis.com/v1/presentations/deck-notes") {
                    presentationReads++;
                    return presentationReads == 1
                        ? Json("{\"presentationId\":\"deck-notes\",\"revisionId\":\"revision-1\",\"slides\":[{\"objectId\":\"initial-slide\"}]}")
                        : Json("{\"presentationId\":\"deck-notes\",\"revisionId\":\"revision-2\",\"slides\":[{\"objectId\":\"officeimo_slide_0001_0001\",\"slideProperties\":{\"notesPage\":{\"notesProperties\":{\"speakerNotesObjectId\":\"notes-body\"}}}}]}");
                }
                if (request.Method == HttpMethod.Post && uri.EndsWith(":batchUpdate", StringComparison.Ordinal)) {
                    batchBodies.Add(await request.Content!.ReadAsStringAsync().ConfigureAwait(false));
                    string revision = batchBodies.Count == 1 ? "revision-2" : "revision-3";
                    return Json("{\"presentationId\":\"deck-notes\",\"writeControl\":{\"requiredRevisionId\":\"" + revision + "\"}}");
                }
                if (request.Method == HttpMethod.Get && request.RequestUri.Host == "www.googleapis.com") return Json("{\"id\":\"deck-notes\",\"name\":\"Notes\",\"mimeType\":\"application/vnd.google-apps.presentation\",\"version\":3}");
                return new HttpResponseMessage(HttpStatusCode.NotFound) { Content = new StringContent("unexpected " + uri) };
            }));

            GooglePresentationReference result = await presentation.ExportToGoogleSlidesAsync(Session(httpClient));

            Assert.Equal(2, presentationReads);
            Assert.Equal(2, batchBodies.Count);
            Assert.Contains("\"deleteText\":{\"objectId\":\"notes-body\"", batchBodies[1]);
            Assert.Contains("Presenter-only context", batchBodies[1]);
            Assert.Contains("\"requiredRevisionId\":\"revision-2\"", batchBodies[1]);
            Assert.Equal("revision-3", result.RevisionId);
        }

        [Fact]
        public async Task Exporter_OverwriteLatest_DoesNotSendRevisionGuard() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.AddSlide().AddTextBox("Replacement");
            string? batchBody = null;
            using var httpClient = new HttpClient(new DelegateHandler(async request => {
                if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "slides.googleapis.com") return Json("{\"presentationId\":\"existing\",\"revisionId\":\"remote\",\"slides\":[]}");
                if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri.EndsWith(":batchUpdate", StringComparison.Ordinal)) {
                    batchBody = await request.Content!.ReadAsStringAsync().ConfigureAwait(false);
                    return Json("{\"presentationId\":\"existing\"}");
                }
                if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "www.googleapis.com") return Json("{\"id\":\"existing\",\"name\":\"Existing\",\"mimeType\":\"application/vnd.google-apps.presentation\"}");
                return new HttpResponseMessage(HttpStatusCode.NotFound);
            }));

            await presentation.ExportToGoogleSlidesAsync(Session(httpClient), new GoogleSlidesSaveOptions {
                Location = new GoogleDriveFileLocation { ExistingFileId = "existing" },
                Replace = new GoogleSlidesReplaceOptions { ConflictMode = GoogleSlidesRevisionConflictMode.OverwriteLatest },
            });

            Assert.NotNull(batchBody);
            Assert.DoesNotContain("writeControl", batchBody);
        }

        [Fact]
        public async Task Exporter_RejectsStaleExistingRevisionBeforeMutation() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.AddSlide().AddTextBox("Local");
            int mutations = 0;
            using var httpClient = new HttpClient(new DelegateHandler(request => {
                if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "slides.googleapis.com") return Task.FromResult(Json("{\"presentationId\":\"existing\",\"revisionId\":\"remote\",\"slides\":[]}"));
                if (request.Method == HttpMethod.Post) mutations++;
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }));

            GoogleWorkspaceConflictException exception = await Assert.ThrowsAsync<GoogleWorkspaceConflictException>(() => presentation.ExportToGoogleSlidesAsync(Session(httpClient), new GoogleSlidesSaveOptions {
                Location = new GoogleDriveFileLocation { ExistingFileId = "existing" },
                Replace = new GoogleSlidesReplaceOptions { ExpectedRevisionId = "observed" },
            }));

            Assert.Equal("observed", exception.ExpectedVersion);
            Assert.Equal("remote", exception.ActualVersion);
            Assert.Equal(0, mutations);
        }

        [Fact]
        public async Task NativeImporter_ProjectsTextTableAndNotesWhenDriveExportIsDisabled() {
            using var httpClient = new HttpClient(new DelegateHandler(request => {
                if (request.RequestUri!.Host == "www.googleapis.com") return Task.FromResult(Json("{\"id\":\"deck-import\",\"name\":\"Import\",\"mimeType\":\"application/vnd.google-apps.presentation\",\"version\":4,\"capabilities\":{\"canDownload\":false}}"));
                const string slides = "{\"presentationId\":\"deck-import\",\"title\":\"Import\",\"revisionId\":\"r4\",\"pageSize\":{\"width\":{\"magnitude\":720,\"unit\":\"PT\"},\"height\":{\"magnitude\":405,\"unit\":\"PT\"}},\"slides\":[{\"objectId\":\"slide-1\",\"pageElements\":[{\"objectId\":\"text-1\",\"size\":{\"width\":{\"magnitude\":300,\"unit\":\"PT\"},\"height\":{\"magnitude\":80,\"unit\":\"PT\"}},\"transform\":{\"translateX\":20,\"translateY\":30,\"unit\":\"PT\"},\"shape\":{\"shapeType\":\"TEXT_BOX\",\"text\":{\"textElements\":[{\"textRun\":{\"content\":\"Imported text\",\"style\":{\"bold\":true}}}]}}},{\"objectId\":\"table-1\",\"size\":{\"width\":{\"magnitude\":300,\"unit\":\"PT\"},\"height\":{\"magnitude\":100,\"unit\":\"PT\"}},\"transform\":{\"translateX\":30,\"translateY\":130,\"unit\":\"PT\"},\"table\":{\"rows\":1,\"columns\":1,\"tableRows\":[{\"tableCells\":[{\"text\":{\"textElements\":[{\"textRun\":{\"content\":\"Cell\"}}]}}]}]}}],\"slideProperties\":{\"isSkipped\":true,\"notesPage\":{\"notesProperties\":{\"speakerNotesObjectId\":\"notes-body\"},\"pageElements\":[{\"objectId\":\"notes-body\",\"shape\":{\"text\":{\"textElements\":[{\"textRun\":{\"content\":\"Imported notes\"}}]}}}]}}}]}";
                return Task.FromResult(Json(slides));
            }));

            GoogleSlidesImportResult imported = await new GoogleSlidesImporter().ImportAsync("deck-import", Session(httpClient), new GoogleSlidesImportOptions { Mode = GoogleSlidesImportMode.Native });
            using (imported.Presentation) {
                PowerPointSlide slide = Assert.Single(imported.Presentation.Slides);
                Assert.True(slide.Hidden);
                Assert.Contains(slide.TextBoxes, text => text.Text == "Imported text");
                Assert.Equal("Cell", Assert.Single(slide.Tables).RowItems[0].Cells[0].Text);
                Assert.Equal("Imported notes", slide.Notes.Text);
                Assert.Equal("r4", imported.Source.RevisionId);
            }
        }

        [Fact]
        public async Task NativeImporter_PreservesGifImages() {
            byte[] gif = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==");
            using var httpClient = new HttpClient(new DelegateHandler(request => {
                if (request.RequestUri!.Host == "www.googleapis.com") return Task.FromResult(Json("{\"id\":\"deck-gif\",\"mimeType\":\"application/vnd.google-apps.presentation\",\"capabilities\":{\"canDownload\":true}}"));
                if (request.RequestUri.Host == "images.example.test") {
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(gif) });
                }
                const string slides = "{\"presentationId\":\"deck-gif\",\"pageSize\":{\"width\":{\"magnitude\":720,\"unit\":\"PT\"},\"height\":{\"magnitude\":405,\"unit\":\"PT\"}},\"slides\":[{\"objectId\":\"slide-1\",\"pageElements\":[{\"objectId\":\"image-1\",\"size\":{\"width\":{\"magnitude\":100,\"unit\":\"PT\"},\"height\":{\"magnitude\":100,\"unit\":\"PT\"}},\"transform\":{\"translateX\":20,\"translateY\":30,\"unit\":\"PT\"},\"image\":{\"contentUrl\":\"https://images.example.test/image.gif\"}}]}]}";
                return Task.FromResult(Json(slides));
            }));

            GoogleSlidesImportResult imported = await new GoogleSlidesImporter().ImportAsync("deck-gif", Session(httpClient), new GoogleSlidesImportOptions { Mode = GoogleSlidesImportMode.Native });
            using (imported.Presentation) {
                PowerPointPicture picture = Assert.Single(Assert.Single(imported.Presentation.Slides).Pictures);
                Assert.Equal("image/gif", picture.ContentType);
                Assert.DoesNotContain(imported.Report.Notices, notice => notice.Code == "SLIDES.IMPORT.IMAGE_FALLBACK");
            }
        }

        [Fact]
        public void DiffPlanner_DetectsIndependentEdits() {
            var checkpoint = new GoogleSlidesSyncCheckpoint(); checkpoint.ContentHashes["slide/1"] = "base";
            List<GoogleSlidesDiffItem> items = GoogleSlidesDiffPlanner.Compare(new Dictionary<string, string> { ["slide/1"] = "local" }, new Dictionary<string, string> { ["slide/1"] = "remote" }, checkpoint);
            Assert.Equal(GoogleSlidesDiffKind.Conflict, Assert.Single(items).Kind);
        }

        [Fact]
        public void SupportCatalog_IsExplicitAboutRasterAndDriveFallbacks() {
            Assert.Contains(GoogleSlidesFeatureSupportCatalog.Features, row => row.Feature == "Charts and SmartArt" && row.Export == GoogleSlidesFeatureSupportLevel.Rasterized);
            Assert.Contains(GoogleSlidesFeatureSupportCatalog.Features, row => row.Import == GoogleSlidesFeatureSupportLevel.DriveFallback);
        }

        private static GoogleWorkspaceSession Session(HttpClient client) => new GoogleWorkspaceSession(new StaticAccessTokenCredentialSource("token"), new GoogleWorkspaceSessionOptions { HttpClient = client });
        private static HttpResponseMessage Json(string value) => new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(value, Encoding.UTF8, "application/json") };
        private sealed class DelegateHandler : HttpMessageHandler {
            private readonly Func<HttpRequestMessage, Task<HttpResponseMessage>> _handler;
            public DelegateHandler(Func<HttpRequestMessage, Task<HttpResponseMessage>> handler) { _handler = handler; }
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => _handler(request);
        }
    }
}
