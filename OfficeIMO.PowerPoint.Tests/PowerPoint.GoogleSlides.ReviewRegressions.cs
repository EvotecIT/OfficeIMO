using OfficeIMO.GoogleWorkspace;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.GoogleSlides;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed partial class GoogleSlidesTests {
        [Theory]
        [InlineData(true, "SLIDES.REPLACE.DRIVE_ACCESS_REQUIRED")]
        [InlineData(false, "SLIDES.REPLACE.DRIVE_EDIT_REQUIRED")]
        public async Task Exporter_PreflightsExistingDriveTargetBeforeSlidesMutation(bool inaccessible, string expectedCode) {
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            presentation.AddSlide().AddTextBox("Local");
            int slidesReads = 0;
            int mutations = 0;
            using var httpClient = new HttpClient(new DelegateHandler(request => {
                if (request.Method == HttpMethod.Get && request.RequestUri!.Host == "www.googleapis.com") {
                    if (inaccessible) {
                        return Task.FromResult(new HttpResponseMessage(HttpStatusCode.Forbidden) {
                            Content = new StringContent("{\"error\":{\"message\":\"forbidden\"}}", Encoding.UTF8, "application/json")
                        });
                    }
                    return Task.FromResult(Json("{\"id\":\"existing\",\"capabilities\":{\"canEdit\":false}}"));
                }
                if (request.RequestUri!.Host == "slides.googleapis.com") slidesReads++;
                if (request.Method != HttpMethod.Get) mutations++;
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }));

            GoogleWorkspacePreflightException exception = await Assert.ThrowsAsync<GoogleWorkspacePreflightException>(() =>
                presentation.ExportToGoogleSlidesAsync(Session(httpClient), new GoogleSlidesSaveOptions {
                    Location = new GoogleDriveFileLocation { ExistingFileId = "existing" },
                    Replace = new GoogleSlidesReplaceOptions { ConflictMode = GoogleSlidesRevisionConflictMode.OverwriteLatest },
                }));

            Assert.Contains(exception.Report.Notices, notice => notice.Code == expectedCode);
            Assert.Equal(0, slidesReads);
            Assert.Equal(0, mutations);
        }

        [Fact]
        public void BatchCompiler_RasterizesOrReportsCroppedImages() {
            byte[] image = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
            using PowerPointPresentation presentation = PowerPointPresentation.Create();
            using var stream = new MemoryStream(image);
            PowerPointPicture picture = presentation.AddSlide().AddPicture(stream, ImagePartType.Png);
            picture.Crop(10, 0, 0, 0);

            GoogleSlidesBatch rasterized = GoogleSlidesBatchCompiler.Build(
                presentation,
                new GoogleSlidesSaveOptions(),
                materializeRasterImages: false);
            GoogleSlidesBatch reported = GoogleSlidesBatchCompiler.Build(
                presentation,
                new GoogleSlidesSaveOptions { ComplexSlides = GoogleSlidesComplexSlideMode.PreferNativeAndReport },
                materializeRasterImages: false);

            Assert.True(Assert.Single(rasterized.Slides).IsRasterized);
            Assert.Contains(rasterized.Plan.Report.Notices, notice => notice.Code == "SLIDES.COMPLEX_SLIDE.RASTERIZED");
            Assert.False(Assert.Single(reported.Slides).IsRasterized);
            Assert.Empty(reported.Slides[0].Elements.OfType<GoogleSlidesImage>());
            Assert.Contains(reported.Plan.Report.Notices, notice => notice.Code == "SLIDES.IMAGE.CROP_SKIPPED");
        }
    }
}
