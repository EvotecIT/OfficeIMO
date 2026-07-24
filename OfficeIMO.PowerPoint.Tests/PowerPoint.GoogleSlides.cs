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
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Tests {
    public sealed partial class GoogleSlidesTests {
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
        public void BatchCompiler_PreservesTextBearingPresetGeometry() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddTextShapePoints(A.ShapeTypeValues.RightArrow, "Next step", 20, 30, 180, 60);
            slide.AddTextBoxPoints("Plain text", 20, 110, 180, 60);

            GoogleSlidesBatch batch = presentation.BuildGoogleSlidesBatch();

            GoogleSlidesTextBox[] textShapes = Assert.Single(batch.Slides).Elements.OfType<GoogleSlidesTextBox>().ToArray();
            Assert.Equal("RIGHT_ARROW", Assert.Single(textShapes, shape => shape.Text == "Next step").ShapeType);
            Assert.Equal("TEXT_BOX", Assert.Single(textShapes, shape => shape.Text == "Plain text").ShapeType);
        }

        [Fact]
        public void BatchCompiler_DropsMalformedRichTextColorsBeforeJsonGeneration() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointTextRun run = presentation.AddSlide()
                .AddTextBoxPoints("Unsafe color", 20, 30, 300, 80)
                .Paragraphs[0].Runs[0];
            run.Color = "GG0000";

            GoogleSlidesTextBox text = Assert.Single(
                Assert.Single(presentation.BuildGoogleSlidesBatch().Slides)
                    .Elements.OfType<GoogleSlidesTextBox>());

            Assert.Null(text.ForegroundColorHex);
            Assert.Null(Assert.Single(text.TextRuns).ForegroundColorHex);
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
        public void BatchCompiler_RasterizesUnmappedAutoShapesInsteadOfChangingTheirGeometry() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.AddSlide().AddShapePoints(A.ShapeTypeValues.Cloud, 20, 20, 160, 90);

            GoogleSlidesBatch batch = presentation.BuildGoogleSlidesBatch();

            GoogleSlidesSlide slide = Assert.Single(batch.Slides);
            Assert.True(slide.IsRasterized);
            Assert.Single(slide.Elements.OfType<GoogleSlidesImage>());
            Assert.Empty(slide.Elements.OfType<GoogleSlidesShape>());
            Assert.Equal(1, batch.Plan.RasterizedSlideCount);
        }

        [Fact]
        public void BatchCompiler_RasterizesMergedTablesToPreserveCellLayout() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointTable table = presentation.AddSlide().AddTablePoints(2, 2, 20, 20, 300, 120);
            table.MergeCells(0, 0, 1, 1);

            GoogleSlidesBatch batch = presentation.BuildGoogleSlidesBatch();

            GoogleSlidesSlide slide = Assert.Single(batch.Slides);
            Assert.True(slide.IsRasterized);
            Assert.Single(slide.Elements.OfType<GoogleSlidesImage>());
            Assert.Empty(slide.Elements.OfType<GoogleSlidesTable>());
            Assert.Equal(1, batch.Plan.RasterizedSlideCount);
        }

        [Fact]
        public void PlanBuilder_ReportsRasterFallbackWithoutMaterializingPngBytes() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.AddSlide().AddShapePoints(A.ShapeTypeValues.Cloud, 20, 20, 160, 90);

            GoogleSlidesBatch planningBatch = GoogleSlidesBatchCompiler.Build(
                presentation,
                new GoogleSlidesSaveOptions(),
                materializeRasterImages: false);
            GoogleSlidesTranslationPlan publicPlan = new GoogleSlidesExporter().BuildPlan(presentation);

            Assert.True(Assert.Single(planningBatch.Slides).IsRasterized);
            Assert.Empty(planningBatch.Slides[0].Elements.OfType<GoogleSlidesImage>());
            Assert.Equal(1, planningBatch.Plan.RasterizedSlideCount);
            Assert.Equal(1, publicPlan.RasterizedSlideCount);
        }

        [Fact]
        public void PlanBuilder_RasterizesOrReportsUnsupportedSlideBackgrounds() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.SetBackgroundGradient("112233", "445566", 45);

            GoogleSlidesBatch rasterPlan = GoogleSlidesBatchCompiler.Build(
                presentation,
                new GoogleSlidesSaveOptions { ComplexSlides = GoogleSlidesComplexSlideMode.RasterizeComplexSlides },
                materializeRasterImages: false);
            GoogleSlidesBatch nativePlan = GoogleSlidesBatchCompiler.Build(
                presentation,
                new GoogleSlidesSaveOptions { ComplexSlides = GoogleSlidesComplexSlideMode.PreferNativeAndReport },
                materializeRasterImages: false);

            Assert.True(Assert.Single(rasterPlan.Slides).IsRasterized);
            Assert.Empty(rasterPlan.Slides[0].Elements.OfType<GoogleSlidesImage>());
            Assert.Equal(1, rasterPlan.Plan.UnsupportedElementCount);
            Assert.False(Assert.Single(nativePlan.Slides).IsRasterized);
            Assert.Contains(nativePlan.Plan.Report.Notices, notice => notice.Code == "SLIDES.BACKGROUND.SKIPPED");
        }

        [Fact]
        public void BatchCompiler_PreservesSupportedImageBackgroundNatively() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.SetBackgroundImage(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Images", "BackgroundImage.png"));

            GoogleSlidesBatch batch = GoogleSlidesBatchCompiler.Build(
                presentation,
                new GoogleSlidesSaveOptions(),
                materializeRasterImages: false);

            GoogleSlidesSlide compiled = Assert.Single(batch.Slides);
            Assert.False(compiled.IsRasterized);
            Assert.NotNull(compiled.BackgroundImage);
            Assert.Equal("image/png", compiled.BackgroundImage!.ContentType);
            Assert.DoesNotContain(batch.Plan.Report.Notices, notice => notice.Code == "SLIDES.BACKGROUND.SKIPPED");

            slide.AddShapePoints(A.ShapeTypeValues.Cloud, 20, 20, 160, 90);
            GoogleSlidesBatch rasterized = GoogleSlidesBatchCompiler.Build(
                presentation,
                new GoogleSlidesSaveOptions(),
                materializeRasterImages: false);
            GoogleSlidesSlide rasterizedSlide = Assert.Single(rasterized.Slides);
            Assert.True(rasterizedSlide.IsRasterized);
            Assert.Null(rasterizedSlide.BackgroundImage);
        }

        [Fact]
        public async Task Exporter_CreatesAndReplacesInitialSlideWithRevisionGuard() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide authoredSlide = presentation.AddSlide();
            authoredSlide.Hidden = true;
            authoredSlide.BackgroundColor = "112233";
            PowerPointTextBox mixedText = authoredSlide.AddTextBoxPoints("Hello ", 20, 30, 300, 80);
            mixedText.Paragraphs[0].AddRun("Slides", run => run.Bold = true);
            authoredSlide.AddTextShapePoints(A.ShapeTypeValues.RightArrow, "Next step", 340, 30, 160, 80);
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
            using (JsonDocument payload = JsonDocument.Parse(body)) {
                JsonElement[] requests = payload.RootElement.GetProperty("requests").EnumerateArray().ToArray();
                int keeperCreateIndex = Array.FindIndex(requests, request =>
                    request.TryGetProperty("createSlide", out JsonElement createSlide)
                    && createSlide.GetProperty("objectId").GetString()!.StartsWith("officeimo_replacement_keeper", StringComparison.Ordinal));
                Assert.True(keeperCreateIndex >= 0);
                string keeperSlideId = requests[keeperCreateIndex].GetProperty("createSlide").GetProperty("objectId").GetString()!;
                int oldSlideDeleteIndex = Array.FindIndex(requests, request =>
                    request.TryGetProperty("deleteObject", out JsonElement deleteObject)
                    && deleteObject.GetProperty("objectId").GetString() == "initial-slide");
                int authoredSlideCreateIndex = Array.FindIndex(requests, request =>
                    request.TryGetProperty("createSlide", out JsonElement createSlide)
                    && createSlide.GetProperty("objectId").GetString() == "officeimo_slide_0001_0001");
                int keeperDeleteIndex = Array.FindIndex(requests, request =>
                    request.TryGetProperty("deleteObject", out JsonElement deleteObject)
                    && deleteObject.GetProperty("objectId").GetString() == keeperSlideId);
                Assert.True(keeperCreateIndex < oldSlideDeleteIndex);
                Assert.True(oldSlideDeleteIndex < authoredSlideCreateIndex);
                Assert.True(authoredSlideCreateIndex < keeperDeleteIndex);
                JsonElement update = Assert.Single(requests, request => request.TryGetProperty("updatePageProperties", out _))
                    .GetProperty("updatePageProperties");
                Assert.Equal("officeimo_slide_0001_0001", update.GetProperty("objectId").GetString());
                Assert.Equal("pageBackgroundFill.solidFill.color", update.GetProperty("fields").GetString());
                JsonElement rgb = update.GetProperty("pageProperties").GetProperty("pageBackgroundFill")
                    .GetProperty("solidFill").GetProperty("color").GetProperty("rgbColor");
                Assert.Equal(0x11 / 255d, rgb.GetProperty("red").GetDouble(), 6);
                Assert.DoesNotContain(requests, request => request.TryGetProperty("updateSlideProperties", out JsonElement slideProperties)
                    && slideProperties.GetProperty("fields").GetString() == "background");
                JsonElement textStyle = Assert.Single(requests, request =>
                    request.TryGetProperty("updateTextStyle", out JsonElement updateTextStyle)
                    && updateTextStyle.GetProperty("style").TryGetProperty("bold", out _))
                    .GetProperty("updateTextStyle");
                Assert.Equal("FIXED_RANGE", textStyle.GetProperty("textRange").GetProperty("type").GetString());
                Assert.Equal(6, textStyle.GetProperty("textRange").GetProperty("startIndex").GetInt32());
                Assert.Equal(12, textStyle.GetProperty("textRange").GetProperty("endIndex").GetInt32());
            }
            Assert.Contains("\"createShape\"", body);
            Assert.Contains("Hello Slides", body);
            Assert.Contains("\"shapeType\":\"RIGHT_ARROW\"", body);
            Assert.Contains("Next step", body);
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
        public async Task Exporter_ProjectsPowerPointRotationIntoSlidesAffineTransform() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointTextBox textBox = presentation.AddSlide().AddTextBoxPoints("Rotated", 10, 20, 100, 40);
            textBox.Rotation = 90;
            string? batchBody = null;
            using var httpClient = new HttpClient(new DelegateHandler(async request => {
                string uri = request.RequestUri!.AbsoluteUri;
                if (request.Method == HttpMethod.Post && uri == "https://slides.googleapis.com/v1/presentations") return Json("{\"presentationId\":\"deck-rotated\"}");
                if (request.Method == HttpMethod.Get && uri == "https://slides.googleapis.com/v1/presentations/deck-rotated") {
                    return Json("{\"presentationId\":\"deck-rotated\",\"revisionId\":\"revision-1\",\"slides\":[{\"objectId\":\"initial-slide\"}]}");
                }
                if (request.Method == HttpMethod.Post && uri.EndsWith(":batchUpdate", StringComparison.Ordinal)) {
                    batchBody = await request.Content!.ReadAsStringAsync().ConfigureAwait(false);
                    return Json("{\"writeControl\":{\"requiredRevisionId\":\"revision-2\"}}");
                }
                if (request.Method == HttpMethod.Get && request.RequestUri.Host == "www.googleapis.com") return Json("{\"id\":\"deck-rotated\",\"mimeType\":\"application/vnd.google-apps.presentation\"}");
                return new HttpResponseMessage(HttpStatusCode.NotFound);
            }));

            await presentation.ExportToGoogleSlidesAsync(Session(httpClient));

            using JsonDocument payload = JsonDocument.Parse(Assert.IsType<string>(batchBody));
            JsonElement transform = payload.RootElement.GetProperty("requests").EnumerateArray()
                .Single(request => request.TryGetProperty("createShape", out _))
                .GetProperty("createShape").GetProperty("elementProperties").GetProperty("transform");
            Assert.Equal(0, transform.GetProperty("scaleX").GetDouble(), 12);
            Assert.Equal(0, transform.GetProperty("scaleY").GetDouble(), 12);
            Assert.Equal(-1, transform.GetProperty("shearX").GetDouble(), 12);
            Assert.Equal(1, transform.GetProperty("shearY").GetDouble(), 12);
            Assert.Equal(80, transform.GetProperty("translateX").GetDouble(), 12);
            Assert.Equal(-10, transform.GetProperty("translateY").GetDouble(), 12);
        }

        [Fact]
        public async Task Exporter_PreservesShapeReflectionAndBasicAppearance() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointAutoShape shape = presentation.AddSlide().AddRectanglePoints(10, 20, 100, 40);
            shape.HorizontalFlip = true;
            shape.FillColor = "336699";
            shape.FillTransparency = 25;
            shape.OutlineColor = "CC3300";
            shape.OutlineWidthPoints = 3;
            string? batchBody = null;
            using var httpClient = new HttpClient(new DelegateHandler(async request => {
                string uri = request.RequestUri!.AbsoluteUri;
                if (request.Method == HttpMethod.Post && uri == "https://slides.googleapis.com/v1/presentations") return Json("{\"presentationId\":\"deck-styled-shape\"}");
                if (request.Method == HttpMethod.Get && uri == "https://slides.googleapis.com/v1/presentations/deck-styled-shape") {
                    return Json("{\"presentationId\":\"deck-styled-shape\",\"revisionId\":\"revision-1\",\"slides\":[{\"objectId\":\"initial-slide\"}]}");
                }
                if (request.Method == HttpMethod.Post && uri.EndsWith(":batchUpdate", StringComparison.Ordinal)) {
                    batchBody = await request.Content!.ReadAsStringAsync().ConfigureAwait(false);
                    return Json("{\"writeControl\":{\"requiredRevisionId\":\"revision-2\"}}");
                }
                if (request.Method == HttpMethod.Get && request.RequestUri.Host == "www.googleapis.com") return Json("{\"id\":\"deck-styled-shape\",\"mimeType\":\"application/vnd.google-apps.presentation\"}");
                return new HttpResponseMessage(HttpStatusCode.NotFound);
            }));

            await presentation.ExportToGoogleSlidesAsync(Session(httpClient));

            using JsonDocument payload = JsonDocument.Parse(Assert.IsType<string>(batchBody));
            JsonElement[] requests = payload.RootElement.GetProperty("requests").EnumerateArray().ToArray();
            JsonElement transform = Assert.Single(requests, request => request.TryGetProperty("createShape", out JsonElement create)
                && create.GetProperty("shapeType").GetString() == "RECTANGLE")
                .GetProperty("createShape").GetProperty("elementProperties").GetProperty("transform");
            Assert.Equal(-1, transform.GetProperty("scaleX").GetDouble(), 12);
            Assert.Equal(1, transform.GetProperty("scaleY").GetDouble(), 12);
            Assert.Equal(110, transform.GetProperty("translateX").GetDouble(), 12);
            Assert.Equal(20, transform.GetProperty("translateY").GetDouble(), 12);

            JsonElement styleUpdate = Assert.Single(requests, request => request.TryGetProperty("updateShapeProperties", out _))
                .GetProperty("updateShapeProperties");
            Assert.Contains("shapeBackgroundFill.solidFill.color", styleUpdate.GetProperty("fields").GetString(), StringComparison.Ordinal);
            Assert.Contains("outline.outlineFill.solidFill.color", styleUpdate.GetProperty("fields").GetString(), StringComparison.Ordinal);
            JsonElement shapeProperties = styleUpdate.GetProperty("shapeProperties");
            JsonElement fill = shapeProperties.GetProperty("shapeBackgroundFill").GetProperty("solidFill");
            Assert.Equal(0.75, fill.GetProperty("alpha").GetDouble(), 12);
            Assert.Equal(0x33 / 255d, fill.GetProperty("color").GetProperty("rgbColor").GetProperty("red").GetDouble(), 12);
            JsonElement outline = shapeProperties.GetProperty("outline");
            Assert.Equal(3, outline.GetProperty("weight").GetProperty("magnitude").GetDouble(), 12);
            Assert.Equal(0xCC / 255d, outline.GetProperty("outlineFill").GetProperty("solidFill").GetProperty("color")
                .GetProperty("rgbColor").GetProperty("red").GetDouble(), 12);
        }

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task Exporter_WritesSpeakerNotesAfterSlidesExist(bool existingNotes) {
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
                        : Json(existingNotes
                            ? "{\"presentationId\":\"deck-notes\",\"revisionId\":\"revision-2\",\"slides\":[{\"objectId\":\"officeimo_slide_0001_0001\",\"slideProperties\":{\"notesPage\":{\"notesProperties\":{\"speakerNotesObjectId\":\"notes-body\"},\"pageElements\":[{\"objectId\":\"notes-body\",\"shape\":{\"text\":{\"textElements\":[{\"textRun\":{\"content\":\"Existing notes\\n\"}}]}}}]}}}] }"
                            : "{\"presentationId\":\"deck-notes\",\"revisionId\":\"revision-2\",\"slides\":[{\"objectId\":\"officeimo_slide_0001_0001\",\"slideProperties\":{\"notesPage\":{\"notesProperties\":{\"speakerNotesObjectId\":\"notes-body\"},\"pageElements\":[{\"objectId\":\"notes-body\",\"shape\":{\"text\":{\"textElements\":[{\"textRun\":{\"content\":\"\\n\"}}]}}}]}}}] }");
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
            if (existingNotes) {
                Assert.Contains("\"deleteText\":{\"objectId\":\"notes-body\"", batchBodies[1]);
            } else {
                Assert.DoesNotContain("\"deleteText\":{\"objectId\":\"notes-body\"", batchBodies[1]);
            }
            Assert.Contains("Presenter-only context", batchBodies[1]);
            Assert.Contains("\"requiredRevisionId\":\"revision-2\"", batchBodies[1]);
            Assert.Equal("revision-3", result.RevisionId);
        }

        [Fact]
        public async Task Exporter_OverwriteLatest_DoesNotSendRevisionGuard() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.AddSlide().AddTextBox("Replacement");
            string? batchBody = null;
            int presentationReads = 0;
            using var httpClient = new HttpClient(new DelegateHandler(async request => {
                if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "slides.googleapis.com") {
                    presentationReads++;
                    string revision = presentationReads == 1 ? "remote" : "fresh";
                    return Json("{\"presentationId\":\"existing\",\"revisionId\":\"" + revision + "\",\"slides\":[]}");
                }
                if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri.EndsWith(":batchUpdate", StringComparison.Ordinal)) {
                    batchBody = await request.Content!.ReadAsStringAsync().ConfigureAwait(false);
                    return Json("{\"presentationId\":\"existing\"}");
                }
                if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "www.googleapis.com") return Json("{\"id\":\"existing\",\"name\":\"Existing\",\"mimeType\":\"application/vnd.google-apps.presentation\"}");
                return new HttpResponseMessage(HttpStatusCode.NotFound);
            }));

            GooglePresentationReference result = await presentation.ExportToGoogleSlidesAsync(Session(httpClient), new GoogleSlidesSaveOptions {
                Location = new GoogleDriveFileLocation { ExistingFileId = "existing" },
                Replace = new GoogleSlidesReplaceOptions { ConflictMode = GoogleSlidesRevisionConflictMode.OverwriteLatest },
            });

            Assert.NotNull(batchBody);
            Assert.DoesNotContain("writeControl", batchBody);
            Assert.Equal(2, presentationReads);
            Assert.Equal("fresh", result.RevisionId);
        }

        [Fact]
        public async Task Exporter_RejectsStaleExistingRevisionBeforeMutation() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.AddSlide().AddTextBox("Local");
            int mutations = 0;
            using var httpClient = new HttpClient(new DelegateHandler(request => {
                if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "www.googleapis.com") return Task.FromResult(Json("{\"id\":\"existing\",\"capabilities\":{\"canEdit\":true}}"));
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
        public async Task Exporter_RejectsFolderFromUnexpectedSharedDriveBeforeTemplateCopy() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.AddSlide().AddTextBox("Local");
            int mutations = 0;
            using var httpClient = new HttpClient(new DelegateHandler(request => {
                if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "www.googleapis.com") {
                    return Task.FromResult(Json("{\"id\":\"folder-1\",\"name\":\"Folder\",\"mimeType\":\"application/vnd.google-apps.folder\",\"driveId\":\"drive-a\"}"));
                }
                if (request.Method != HttpMethod.Get) mutations++;
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }));

            InvalidOperationException exception = await Assert.ThrowsAsync<InvalidOperationException>(() => presentation.ExportToGoogleSlidesAsync(
                Session(httpClient),
                new GoogleSlidesSaveOptions {
                    TemplatePresentationId = "template-1",
                    Location = new GoogleDriveFileLocation { FolderId = "folder-1", DriveId = "drive-b" },
                }));

            Assert.Contains("drive-a", exception.Message, StringComparison.Ordinal);
            Assert.Contains("drive-b", exception.Message, StringComparison.Ordinal);
            Assert.Equal(0, mutations);
        }

        [Fact]
        public async Task Exporter_PreservesInvalidRequestDiagnosticsWhenRevisionStillMatches() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.AddSlide().AddTextBox("Local");
            int presentationReads = 0;
            using var httpClient = new HttpClient(new DelegateHandler(request => {
                if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "www.googleapis.com") return Task.FromResult(Json("{\"id\":\"existing\",\"capabilities\":{\"canEdit\":true}}"));
                if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "slides.googleapis.com") {
                    presentationReads++;
                    return Task.FromResult(Json("{\"presentationId\":\"existing\",\"revisionId\":\"observed\",\"slides\":[]}"));
                }
                if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri.EndsWith(":batchUpdate", StringComparison.Ordinal)) {
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.BadRequest) {
                        Content = new StringContent("{\"error\":{\"status\":\"INVALID_ARGUMENT\",\"message\":\"Invalid requests[0].createShape\"}}", Encoding.UTF8, "application/json")
                    });
                }
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }));

            GoogleWorkspaceApiException exception = await Assert.ThrowsAsync<GoogleWorkspaceApiException>(() => presentation.ExportToGoogleSlidesAsync(Session(httpClient), new GoogleSlidesSaveOptions {
                Location = new GoogleDriveFileLocation { ExistingFileId = "existing" },
                Replace = new GoogleSlidesReplaceOptions { ExpectedRevisionId = "observed" },
            }));

            Assert.Contains("Invalid requests[0].createShape", exception.ResponseBody);
            Assert.Equal(2, presentationReads);
        }

        [Fact]
        public async Task Exporter_ClassifiesBatchFailureAsConflictAfterRevisionChanges() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.AddSlide().AddTextBox("Local");
            int presentationReads = 0;
            using var httpClient = new HttpClient(new DelegateHandler(request => {
                if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "www.googleapis.com") return Task.FromResult(Json("{\"id\":\"existing\",\"capabilities\":{\"canEdit\":true}}"));
                if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "slides.googleapis.com") {
                    presentationReads++;
                    string revision = presentationReads == 1 ? "observed" : "remote";
                    return Task.FromResult(Json("{\"presentationId\":\"existing\",\"revisionId\":\"" + revision + "\",\"slides\":[]}"));
                }
                if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri.EndsWith(":batchUpdate", StringComparison.Ordinal)) {
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.BadRequest) {
                        Content = new StringContent("{\"error\":{\"status\":\"INVALID_ARGUMENT\"}}", Encoding.UTF8, "application/json")
                    });
                }
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }));

            GoogleWorkspaceConflictException exception = await Assert.ThrowsAsync<GoogleWorkspaceConflictException>(() => presentation.ExportToGoogleSlidesAsync(Session(httpClient), new GoogleSlidesSaveOptions {
                Location = new GoogleDriveFileLocation { ExistingFileId = "existing" },
                Replace = new GoogleSlidesReplaceOptions { ExpectedRevisionId = "observed" },
            }));

            Assert.Equal("observed", exception.ExpectedVersion);
            Assert.Equal("remote", exception.ActualVersion);
            Assert.Equal(2, presentationReads);
        }

        [Fact]
        public async Task Exporter_PreservesLaterBatchDiagnosticsWhenActiveRevisionStillMatches() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddTextBox("Local");
            slide.Notes.Text = "Notes";
            int presentationReads = 0;
            int batchWrites = 0;
            var batchBodies = new List<string>();
            using var httpClient = new HttpClient(new DelegateHandler(async request => {
                if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "www.googleapis.com") return Json("{\"id\":\"existing\",\"capabilities\":{\"canEdit\":true}}");
                if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "slides.googleapis.com") {
                    presentationReads++;
                    return presentationReads == 1
                        ? Json("{\"presentationId\":\"existing\",\"revisionId\":\"observed\",\"slides\":[]}")
                        : Json("{\"presentationId\":\"existing\",\"revisionId\":\"revision-2\",\"slides\":[{\"objectId\":\"officeimo_slide_0001_0001\",\"slideProperties\":{\"notesPage\":{\"notesProperties\":{\"speakerNotesObjectId\":\"notes-body\"}}}}]}");
                }
                if (request.Method == HttpMethod.Post && request.RequestUri!.AbsoluteUri.EndsWith(":batchUpdate", StringComparison.Ordinal)) {
                    batchWrites++;
                    batchBodies.Add(await request.Content!.ReadAsStringAsync().ConfigureAwait(false));
                    if (batchWrites == 1) return Json("{\"writeControl\":{\"requiredRevisionId\":\"revision-2\"}}");
                    return new HttpResponseMessage(HttpStatusCode.BadRequest) {
                        Content = new StringContent("{\"error\":{\"status\":\"INVALID_ARGUMENT\",\"message\":\"Invalid speaker notes request\"}}", Encoding.UTF8, "application/json")
                    };
                }
                return new HttpResponseMessage(HttpStatusCode.NotFound);
            }));

            GoogleWorkspaceApiException exception = await Assert.ThrowsAsync<GoogleWorkspaceApiException>(() => presentation.ExportToGoogleSlidesAsync(Session(httpClient), new GoogleSlidesSaveOptions {
                Location = new GoogleDriveFileLocation { ExistingFileId = "existing" },
                Replace = new GoogleSlidesReplaceOptions { ExpectedRevisionId = "observed" },
            }));

            Assert.Contains("Invalid speaker notes request", exception.ResponseBody);
            Assert.Equal(3, presentationReads);
            Assert.Equal(2, batchWrites);
            Assert.Contains("\"requiredRevisionId\":\"revision-2\"", batchBodies[1]);
        }

        [Fact]
        public async Task NativeImporter_RejectsFilesThatCannotBeDownloaded() {
            int nativeReads = 0;
            using var httpClient = new HttpClient(new DelegateHandler(request => {
                if (request.RequestUri!.Host == "www.googleapis.com") {
                    return Task.FromResult(Json("{\"id\":\"deck-blocked\",\"mimeType\":\"application/vnd.google-apps.presentation\",\"capabilities\":{\"canDownload\":false}}"));
                }
                nativeReads++;
                return Task.FromResult(Json("{\"presentationId\":\"deck-blocked\"}"));
            }));

            InvalidOperationException exception = await Assert.ThrowsAsync<InvalidOperationException>(() =>
                new GoogleSlidesImporter().ImportAsync("deck-blocked", Session(httpClient), new GoogleSlidesImportOptions { Mode = GoogleSlidesImportMode.Native }));

            Assert.Contains("cannot be exported", exception.Message, StringComparison.Ordinal);
            Assert.Equal(0, nativeReads);
        }

        [Fact]
        public async Task NativeImporter_ProjectsTextTableAndNotesWhenDownloadIsAllowed() {
            using var httpClient = new HttpClient(new DelegateHandler(request => {
                if (request.RequestUri!.Host == "www.googleapis.com") return Task.FromResult(Json("{\"id\":\"deck-import\",\"name\":\"Import\",\"mimeType\":\"application/vnd.google-apps.presentation\",\"version\":4,\"capabilities\":{\"canDownload\":true}}"));
                const string slides = "{\"presentationId\":\"deck-import\",\"title\":\"Import\",\"revisionId\":\"r4\",\"pageSize\":{\"width\":{\"magnitude\":720,\"unit\":\"PT\"},\"height\":{\"magnitude\":405,\"unit\":\"PT\"}},\"slides\":[{\"objectId\":\"slide-1\",\"pageProperties\":{\"pageBackgroundFill\":{\"solidFill\":{\"color\":{\"rgbColor\":{\"red\":0.2,\"green\":0.4,\"blue\":0.6}}}}},\"pageElements\":[{\"objectId\":\"text-1\",\"size\":{\"width\":{\"magnitude\":300,\"unit\":\"PT\"},\"height\":{\"magnitude\":80,\"unit\":\"PT\"}},\"transform\":{\"translateX\":20,\"translateY\":30,\"unit\":\"PT\"},\"shape\":{\"shapeType\":\"TEXT_BOX\",\"text\":{\"textElements\":[{\"textRun\":{\"content\":\"Imported \",\"style\":{\"bold\":true,\"foregroundColor\":{\"opaqueColor\":{\"rgbColor\":{\"red\":0.2,\"green\":0.4,\"blue\":0.6}}}}}},{\"textRun\":{\"content\":\"text\\n\",\"style\":{\"italic\":true,\"underline\":true,\"foregroundColor\":{\"opaqueColor\":{\"rgbColor\":{\"red\":0.8,\"green\":0.2,\"blue\":0.1}}}}}}]}}},{\"objectId\":\"shape-1\",\"size\":{\"width\":{\"magnitude\":120,\"unit\":\"PT\"},\"height\":{\"magnitude\":60,\"unit\":\"PT\"}},\"transform\":{\"translateX\":400,\"translateY\":30,\"unit\":\"PT\"},\"shape\":{\"shapeType\":\"RECTANGLE\",\"shapeProperties\":{\"shapeBackgroundFill\":{\"propertyState\":\"RENDERED\",\"solidFill\":{\"color\":{\"rgbColor\":{\"red\":0.8,\"green\":0.4,\"blue\":0.2}},\"alpha\":0.75}},\"outline\":{\"propertyState\":\"RENDERED\",\"outlineFill\":{\"solidFill\":{\"color\":{\"rgbColor\":{\"red\":0.2,\"green\":0.4,\"blue\":0.6}}}},\"weight\":{\"magnitude\":2.5,\"unit\":\"PT\"}}}}},{\"objectId\":\"table-1\",\"size\":{\"width\":{\"magnitude\":300,\"unit\":\"PT\"},\"height\":{\"magnitude\":100,\"unit\":\"PT\"}},\"transform\":{\"translateX\":30,\"translateY\":130,\"unit\":\"PT\"},\"table\":{\"rows\":1,\"columns\":1,\"tableRows\":[{\"tableCells\":[{\"text\":{\"textElements\":[{\"textRun\":{\"content\":\"Cell\\n\"}}]}}]}]}}],\"slideProperties\":{\"isSkipped\":true,\"notesPage\":{\"notesProperties\":{\"speakerNotesObjectId\":\"notes-body\"},\"pageElements\":[{\"objectId\":\"notes-body\",\"shape\":{\"text\":{\"textElements\":[{\"textRun\":{\"content\":\"Imported notes\\n\"}}]}}}]}}}]}";
                return Task.FromResult(Json(slides));
            }));

            GoogleSlidesImportResult imported = await new GoogleSlidesImporter().ImportAsync("deck-import", Session(httpClient), new GoogleSlidesImportOptions { Mode = GoogleSlidesImportMode.Native });
            using (imported.Presentation) {
                PowerPointSlide slide = Assert.Single(imported.Presentation.Slides);
                Assert.True(slide.Hidden);
                Assert.Equal("336699", slide.BackgroundColor);
                PowerPointTextBox importedText = Assert.Single(slide.TextBoxes, text => text.Text == "Imported text");
                IReadOnlyList<PowerPointTextRun> importedRuns = Assert.Single(importedText.Paragraphs).Runs;
                Assert.Collection(importedRuns,
                    run => {
                        Assert.Equal("Imported ", run.Text);
                        Assert.True(run.Bold);
                        Assert.Equal("336699", run.Color);
                    },
                    run => {
                        Assert.Equal("text", run.Text);
                        Assert.True(run.Italic);
                        Assert.True(run.Underline);
                        Assert.Equal("CC331A", run.Color);
                    });
                GoogleSlidesTextBox roundTripText = Assert.Single(
                    Assert.Single(GoogleSlidesBatchCompiler.Build(imported.Presentation, new GoogleSlidesSaveOptions()).Slides)
                        .Elements.OfType<GoogleSlidesTextBox>());
                Assert.Collection(roundTripText.TextRuns,
                    run => {
                        Assert.Equal(0, run.StartIndex);
                        Assert.Equal(9, run.EndIndex);
                        Assert.True(run.Bold);
                        Assert.Equal("336699", run.ForegroundColorHex);
                    },
                    run => {
                        Assert.Equal(9, run.StartIndex);
                        Assert.Equal(13, run.EndIndex);
                        Assert.True(run.Italic);
                        Assert.True(run.Underline);
                        Assert.Equal("CC331A", run.ForegroundColorHex);
                    });
                PowerPointAutoShape importedShape = Assert.Single(slide.Shapes.OfType<PowerPointAutoShape>());
                Assert.Equal(A.ShapeTypeValues.Rectangle, importedShape.ShapeType);
                Assert.Equal("CC6633", importedShape.FillColor);
                Assert.Equal(25, importedShape.FillTransparency);
                Assert.Equal("336699", importedShape.OutlineColor);
                Assert.Equal(2.5d, importedShape.OutlineWidthPoints);
                Assert.Equal("Cell", Assert.Single(slide.Tables).RowItems[0].Cells[0].Text);
                Assert.Equal("Imported notes", slide.Notes.Text);
                Assert.Equal("r4", imported.Source.RevisionId);
                GoogleSlidesShape roundTripShape = Assert.Single(
                    Assert.Single(GoogleSlidesBatchCompiler.Build(imported.Presentation, new GoogleSlidesSaveOptions()).Slides)
                        .Elements.OfType<GoogleSlidesShape>());
                Assert.Equal("CC6633", roundTripShape.Style.FillColorHex);
                Assert.Equal(25, roundTripShape.Style.FillTransparencyPercent);
                Assert.Equal("336699", roundTripShape.Style.OutlineColorHex);
                Assert.Equal(2.5d, roundTripShape.Style.OutlineWidthPoints);
            }
        }

        [Fact]
        public async Task NativeImporter_PreservesGeometryForTextBearingShapes() {
            using var httpClient = new HttpClient(new DelegateHandler(request => {
                if (request.RequestUri!.Host == "www.googleapis.com") {
                    return Task.FromResult(Json("{\"id\":\"deck-text-shape\",\"mimeType\":\"application/vnd.google-apps.presentation\",\"capabilities\":{\"canDownload\":true}}"));
                }
                const string slides = "{\"presentationId\":\"deck-text-shape\",\"slides\":[{\"objectId\":\"slide-1\",\"pageElements\":[{\"objectId\":\"arrow-1\",\"size\":{\"width\":{\"magnitude\":180,\"unit\":\"PT\"},\"height\":{\"magnitude\":60,\"unit\":\"PT\"}},\"transform\":{\"translateX\":20,\"translateY\":30,\"unit\":\"PT\"},\"shape\":{\"shapeType\":\"RIGHT_ARROW\",\"text\":{\"textElements\":[{\"textRun\":{\"content\":\"Next step\",\"style\":{\"bold\":true}}}]}}}]}]}";
                return Task.FromResult(Json(slides));
            }));

            GoogleSlidesImportResult imported = await new GoogleSlidesImporter().ImportAsync(
                "deck-text-shape",
                Session(httpClient),
                new GoogleSlidesImportOptions { Mode = GoogleSlidesImportMode.Native });

            using (imported.Presentation) {
                PowerPointTextBox shape = Assert.Single(
                    Assert.Single(imported.Presentation.Slides).TextBoxes,
                    candidate => candidate.Text == "Next step");
                Assert.Equal(A.ShapeTypeValues.RightArrow, shape.ShapeType);
                Assert.True(shape.Paragraphs[0].Runs[0].Bold);
            }
        }

        [Fact]
        public async Task NativeImporter_PreservesRotationAndReportsUnrepresentableShear() {
            byte[] gif = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==");
            using var httpClient = new HttpClient(new DelegateHandler(request => {
                if (request.RequestUri!.Host == "www.googleapis.com") {
                    return Task.FromResult(Json("{\"id\":\"deck-transform\",\"mimeType\":\"application/vnd.google-apps.presentation\",\"capabilities\":{\"canDownload\":true}}"));
                }
                if (request.RequestUri.Host == "lh3.googleusercontent.com") {
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(gif) });
                }
                const string slides = "{\"presentationId\":\"deck-transform\",\"slides\":[{\"objectId\":\"slide-1\",\"pageElements\":[{\"objectId\":\"shape-rotated\",\"size\":{\"width\":{\"magnitude\":80,\"unit\":\"PT\"},\"height\":{\"magnitude\":40,\"unit\":\"PT\"}},\"transform\":{\"scaleX\":0,\"scaleY\":0,\"shearX\":-1,\"shearY\":1,\"translateX\":100,\"translateY\":50,\"unit\":\"PT\"},\"shape\":{\"shapeType\":\"RECTANGLE\"}},{\"objectId\":\"table-rotated\",\"size\":{\"width\":{\"magnitude\":100,\"unit\":\"PT\"},\"height\":{\"magnitude\":40,\"unit\":\"PT\"}},\"transform\":{\"scaleX\":0,\"scaleY\":0,\"shearX\":-1,\"shearY\":1,\"translateX\":200,\"translateY\":50,\"unit\":\"PT\"},\"table\":{\"rows\":1,\"columns\":1,\"tableRows\":[{\"tableCells\":[{\"text\":{\"textElements\":[]}}]}]}},{\"objectId\":\"image-rotated\",\"size\":{\"width\":{\"magnitude\":60,\"unit\":\"PT\"},\"height\":{\"magnitude\":30,\"unit\":\"PT\"}},\"transform\":{\"scaleX\":0,\"scaleY\":0,\"shearX\":-1,\"shearY\":1,\"translateX\":300,\"translateY\":50,\"unit\":\"PT\"},\"image\":{\"contentUrl\":\"https://lh3.googleusercontent.com/image.gif\"}},{\"objectId\":\"shape-skewed\",\"size\":{\"width\":{\"magnitude\":80,\"unit\":\"PT\"},\"height\":{\"magnitude\":40,\"unit\":\"PT\"}},\"transform\":{\"scaleX\":1,\"scaleY\":1,\"shearX\":0.25,\"translateX\":400,\"translateY\":50,\"unit\":\"PT\"},\"shape\":{\"shapeType\":\"RECTANGLE\"}}]}]}";
                return Task.FromResult(Json(slides));
            }));

            GoogleSlidesImportResult imported = await new GoogleSlidesImporter().ImportAsync(
                "deck-transform",
                Session(httpClient),
                new GoogleSlidesImportOptions { Mode = GoogleSlidesImportMode.Native });

            using (imported.Presentation) {
                PowerPointSlide slide = Assert.Single(imported.Presentation.Slides);
                PowerPointAutoShape rotatedShape = Assert.Single(slide.Shapes.OfType<PowerPointAutoShape>(), shape => shape.Name == "shape-rotated");
                Assert.Equal(90d, rotatedShape.Rotation ?? 0, 6);
                Assert.Equal(40d, rotatedShape.LeftPoints, 6);
                Assert.Equal(70d, rotatedShape.TopPoints, 6);
                Assert.Equal(90d, Assert.Single(slide.Tables).Rotation ?? 0, 6);
                Assert.Equal(90d, Assert.Single(slide.Pictures).Rotation ?? 0, 6);
                Assert.Contains(imported.Report.Notices, notice => notice.Code == "SLIDES.IMPORT.TRANSFORM_PARTIAL"
                    && notice.Message.Contains("shape-skewed", StringComparison.Ordinal));
            }
        }

        [Fact]
        public async Task NativeImporter_PreservesGifImages() {
            byte[] gif = Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==");
            using var httpClient = new HttpClient(new DelegateHandler(request => {
                if (request.RequestUri!.Host == "www.googleapis.com") return Task.FromResult(Json("{\"id\":\"deck-gif\",\"mimeType\":\"application/vnd.google-apps.presentation\",\"capabilities\":{\"canDownload\":true}}"));
                if (request.RequestUri.Host == "lh3.googleusercontent.com") {
                    Assert.Equal("https://lh3.googleusercontent.com/image.gif", request.RequestUri.AbsoluteUri);
                    Assert.Null(request.Headers.Authorization);
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(gif) });
                }
                const string slides = "{\"presentationId\":\"deck-gif\",\"pageSize\":{\"width\":{\"magnitude\":720,\"unit\":\"PT\"},\"height\":{\"magnitude\":405,\"unit\":\"PT\"}},\"slides\":[{\"objectId\":\"slide-1\",\"pageProperties\":{\"pageBackgroundFill\":{\"stretchedPictureFill\":{\"contentUrl\":\"https://lh3.googleusercontent.com/image.gif\"}}},\"pageElements\":[{\"objectId\":\"image-1\",\"size\":{\"width\":{\"magnitude\":100,\"unit\":\"PT\"},\"height\":{\"magnitude\":100,\"unit\":\"PT\"}},\"transform\":{\"translateX\":20,\"translateY\":30,\"unit\":\"PT\"},\"image\":{\"contentUrl\":\"https://lh3.googleusercontent.com/image.gif\"}}]}]}";
                return Task.FromResult(Json(slides));
            }));

            GoogleSlidesImportResult imported = await new GoogleSlidesImporter().ImportAsync("deck-gif", Session(httpClient, quotaUser: "tenant-user"), new GoogleSlidesImportOptions { Mode = GoogleSlidesImportMode.Native });
            using (imported.Presentation) {
                PowerPointSlide slide = Assert.Single(imported.Presentation.Slides);
                PowerPointPicture picture = Assert.Single(slide.Pictures);
                Assert.Equal("image/gif", picture.ContentType);
                Assert.Equal(PowerPointSlideBackgroundKind.Image, slide.GetBackground().Kind);
                Assert.Equal("image/gif", slide.GetBackground().ImageContentType);
                Assert.DoesNotContain(imported.Report.Notices, notice => notice.Code == "SLIDES.IMPORT.IMAGE_FALLBACK");
                Assert.DoesNotContain(imported.Report.Notices, notice => notice.Code == "SLIDES.IMPORT.BACKGROUND_IMAGE_FALLBACK");
            }
        }

        [Fact]
        public async Task NativeImporter_RejectsUntrustedImageContentUrls() {
            int untrustedRequests = 0;
            using var httpClient = new HttpClient(new DelegateHandler(request => {
                if (request.RequestUri!.Host == "www.googleapis.com") {
                    return Task.FromResult(Json(
                        "{\"id\":\"deck-untrusted\",\"mimeType\":\"application/vnd.google-apps.presentation\"}"));
                }
                if (request.RequestUri.Host == "attacker.example.test") {
                    untrustedRequests++;
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                        Content = new ByteArrayContent(new byte[1])
                    });
                }
                const string slides = "{\"presentationId\":\"deck-untrusted\",\"slides\":[{\"objectId\":\"slide-1\",\"pageElements\":[{\"objectId\":\"image-1\",\"size\":{\"width\":{\"magnitude\":100,\"unit\":\"PT\"},\"height\":{\"magnitude\":100,\"unit\":\"PT\"}},\"image\":{\"contentUrl\":\"https://attacker.example.test/image.png\"}}]}]}";
                return Task.FromResult(Json(slides));
            }));

            GoogleSlidesImportResult imported = await new GoogleSlidesImporter()
                .ImportAsync("deck-untrusted", Session(httpClient),
                    new GoogleSlidesImportOptions {
                        Mode = GoogleSlidesImportMode.Native
                    });

            using (imported.Presentation) {
                Assert.Equal(0, untrustedRequests);
                Assert.Empty(Assert.Single(imported.Presentation.Slides).Pictures);
                Assert.Contains(imported.Report.Notices, notice =>
                    notice.Code == "SLIDES.IMPORT.IMAGE_FALLBACK");
            }
        }

        [Fact]
        public async Task NativeImporter_EnforcesPerImageResponseLimit() {
            byte[] gif = Convert.FromBase64String(
                "R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==");
            using var httpClient = new HttpClient(new DelegateHandler(request => {
                if (request.RequestUri!.Host == "www.googleapis.com") {
                    return Task.FromResult(Json(
                        "{\"id\":\"deck-image-limit\",\"mimeType\":\"application/vnd.google-apps.presentation\"}"));
                }
                if (request.RequestUri.Host == "lh3.googleusercontent.com") {
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                        Content = new ByteArrayContent(gif)
                    });
                }
                const string slides = "{\"presentationId\":\"deck-image-limit\",\"slides\":[{\"objectId\":\"slide-1\",\"pageElements\":[{\"objectId\":\"image-1\",\"size\":{\"width\":{\"magnitude\":100,\"unit\":\"PT\"},\"height\":{\"magnitude\":100,\"unit\":\"PT\"}},\"image\":{\"contentUrl\":\"https://lh3.googleusercontent.com/image.gif\"}}]}]}";
                return Task.FromResult(Json(slides));
            }));

            GoogleSlidesImportResult imported = await new GoogleSlidesImporter()
                .ImportAsync("deck-image-limit", Session(httpClient),
                    new GoogleSlidesImportOptions {
                        Mode = GoogleSlidesImportMode.Native,
                        MaxImageBytes = 8
                    });

            using (imported.Presentation) {
                Assert.Empty(Assert.Single(imported.Presentation.Slides).Pictures);
                Assert.Contains(imported.Report.Notices, notice =>
                    notice.Code == "SLIDES.IMPORT.IMAGE_FALLBACK");
            }
        }

        [Fact]
        public void DiffPlanner_DetectsIndependentEdits() {
            var checkpoint = new GoogleSlidesSyncCheckpoint(); checkpoint.ContentHashes["slide/1"] = "base";
            List<GoogleSlidesDiffItem> items = GoogleSlidesDiffPlanner.Compare(new Dictionary<string, string> { ["slide/1"] = "local" }, new Dictionary<string, string> { ["slide/1"] = "remote" }, checkpoint);
            Assert.Equal(GoogleSlidesDiffKind.Conflict, Assert.Single(items).Kind);
        }

        [Fact]
        public void DiffPlanner_HashesPictureContentAndCrop() {
            byte[] firstImage = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
            byte[] secondImage = (byte[])firstImage.Clone();
            secondImage[secondImage.Length - 12] ^= 0x01;
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            using var firstStream = new MemoryStream(firstImage);
            PowerPointPicture picture = presentation.AddSlide().AddPicture(firstStream, ImagePartType.Png);
            string baseline = ElementHash(GoogleSlidesDiffPlanner.CreateCheckpoint(presentation));

            using (var secondStream = new MemoryStream(secondImage)) picture.UpdateImage(secondStream, ImagePartType.Png);
            string contentChanged = ElementHash(GoogleSlidesDiffPlanner.CreateCheckpoint(presentation));
            picture.Crop(10, 0, 0, 0);
            string cropChanged = ElementHash(GoogleSlidesDiffPlanner.CreateCheckpoint(presentation));

            Assert.NotEqual(baseline, contentChanged);
            Assert.NotEqual(contentChanged, cropChanged);
        }

        [Fact]
        public void DiffPlanner_HashesNativeGeometryAndEffectiveTextStyle() {
            string ShapeHash(A.ShapeTypeValues shapeType) {
                using PowerPointPresentation presentation = PowerPointPresentation.Create();
                PowerPointAutoShape shape = presentation.AddSlide().AddShapePoints(shapeType, 20, 20, 160, 90);
                Assert.Equal(shapeType, shape.ShapeType);
                return ElementHash(GoogleSlidesDiffPlanner.CreateCheckpoint(presentation));
            }

            Assert.NotEqual(ShapeHash(A.ShapeTypeValues.Rectangle), ShapeHash(A.ShapeTypeValues.RightArrow));

            using PowerPointPresentation transformedPresentation = PowerPointPresentation.Create();
            PowerPointAutoShape transformedShape = transformedPresentation.AddSlide().AddShapePoints(A.ShapeTypeValues.Rectangle, 20, 20, 160, 90);
            string unrotated = ElementHash(GoogleSlidesDiffPlanner.CreateCheckpoint(transformedPresentation));
            transformedShape.Rotation = 30;
            string rotated = ElementHash(GoogleSlidesDiffPlanner.CreateCheckpoint(transformedPresentation));
            Assert.NotEqual(unrotated, rotated);

            using PowerPointPresentation styledPresentation = PowerPointPresentation.Create();
            PowerPointTextBox textBox = styledPresentation.AddSlide().AddTextBoxPoints("Styled text", 20, 20, 180, 60);
            PowerPointTextRun run = Assert.Single(Assert.Single(textBox.Paragraphs).Runs);
            string baseline = ElementHash(GoogleSlidesDiffPlanner.CreateCheckpoint(styledPresentation));

            run.Bold = true;
            run.FontSize = 18;
            run.FontName = "Aptos";
            run.Color = "336699";
            run.Underline = true;
            run.Hyperlink = new Uri("https://example.test/");
            string styled = ElementHash(GoogleSlidesDiffPlanner.CreateCheckpoint(styledPresentation));

            Assert.NotEqual(baseline, styled);
        }

        [Fact]
        public async Task DiffPlanner_ReportsDriveVersionChanges() {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.AddSlide().AddTextBox("Same");
            GoogleSlidesSyncCheckpoint checkpoint = GoogleSlidesDiffPlanner.CreateCheckpoint(presentation, revisionId: "revision-1", driveVersion: 4);
            using var httpClient = new HttpClient(new DelegateHandler(request => {
                if (request.RequestUri!.Host == "www.googleapis.com") return Task.FromResult(Json("{\"id\":\"deck-diff\",\"name\":\"Diff\",\"mimeType\":\"application/vnd.google-apps.presentation\",\"version\":5,\"capabilities\":{\"canDownload\":true}}"));
                const string slides = "{\"presentationId\":\"deck-diff\",\"revisionId\":\"revision-1\",\"slides\":[{\"objectId\":\"slide-1\",\"pageElements\":[{\"objectId\":\"text-1\",\"size\":{\"width\":{\"magnitude\":100,\"unit\":\"PT\"},\"height\":{\"magnitude\":40,\"unit\":\"PT\"}},\"shape\":{\"shapeType\":\"TEXT_BOX\",\"text\":{\"textElements\":[{\"textRun\":{\"content\":\"Same\"}}]}}}]}]}";
                return Task.FromResult(Json(slides));
            }));

            GoogleSlidesDiffPlan plan = await GoogleSlidesDiffPlanner.BuildAsync(presentation, "deck-diff", Session(httpClient), checkpoint);

            Assert.Contains(plan.Items, item => item.Kind == GoogleSlidesDiffKind.RemoteChange && item.Path == "presentation/driveVersion");
        }

        [Fact]
        public void SupportCatalog_IsExplicitAboutRasterAndDriveFallbacks() {
            Assert.Contains(GoogleSlidesFeatureSupportCatalog.Features, row => row.Feature == "Charts and SmartArt" && row.Export == GoogleSlidesFeatureSupportLevel.Rasterized);
            Assert.Contains(GoogleSlidesFeatureSupportCatalog.Features, row => row.Import == GoogleSlidesFeatureSupportLevel.DriveFallback);
        }

        private static GoogleWorkspaceSession Session(HttpClient client, string? quotaUser = null) => new GoogleWorkspaceSession(
            new StaticAccessTokenCredentialSource("token"),
            new GoogleWorkspaceSessionOptions { HttpClient = client, QuotaUser = quotaUser });
        private static HttpResponseMessage Json(string value) => new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(value, Encoding.UTF8, "application/json") };
        private static string ElementHash(GoogleSlidesSyncCheckpoint checkpoint) =>
            Assert.Single(checkpoint.ContentHashes, pair => pair.Key.Contains("/element/", StringComparison.Ordinal)).Value;
        private sealed class DelegateHandler : HttpMessageHandler {
            private readonly Func<HttpRequestMessage, Task<HttpResponseMessage>> _handler;
            public DelegateHandler(Func<HttpRequestMessage, Task<HttpResponseMessage>> handler) { _handler = handler; }
            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) => _handler(request);
        }
    }
}
