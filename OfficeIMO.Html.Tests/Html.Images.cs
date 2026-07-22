using OfficeIMO.Word.Html;
using OfficeIMO.Word;
using OfficeIMO.Html;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_RelativeImage_UsesBaseUrl() {
            var dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(dir);
            var source = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png");
            var dest = Path.Combine(dir, "logo.png");
            File.Copy(source, dest);
            try {
                var baseHref = new Uri(new Uri(Path.Combine(dir, "dummy"), UriKind.Absolute), ".").AbsoluteUri;
                Assert.EndsWith("/", baseHref);
                string html = $"<base href=\"{baseHref}\"><img src=\"logo.png\" alt=\"Logo\" />";
                var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
                    .ToWordDocument(HtmlToWordOptions.CreateTrustedDocumentProfile());
                Assert.Single(doc.Images);
                Assert.Equal("Logo", doc.Images[0].Description);
            } finally {
                File.Delete(dest);
                Directory.Delete(dir);
            }
        }

        [Fact]
        public void HtmlToWord_RelativeImage_UsesOptionsBasePath() {
            var dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(dir);
            var source = Path.Combine(AppContext.BaseDirectory, "..", "..", "..", "..", "Assets", "OfficeIMO.png");
            var dest = Path.Combine(dir, "logo.png");
            File.Copy(source, dest);
            try {
                string html = "<img src=\"logo.png\" alt=\"Logo\" />";
                var options = HtmlToWordOptions.CreateTrustedDocumentProfile();
                options.BasePath = dir;
                HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
                var doc = conversion.Value;
                Assert.Single(doc.Images);
                Assert.Equal("Logo", doc.Images[0].Description);
            } finally {
                File.Delete(dest);
                Directory.Delete(dir);
            }
        }

        [Fact]
        public void HtmlToWord_ImageTitleRoundTripsThroughSavedDocument() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string html = $"<img src=\"{path}\" alt=\"Company logo\" title=\"Quarterly report logo\" width=\"32\" height=\"32\" />";

            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
                .ToWordDocument(HtmlToWordOptions.CreateTrustedDocumentProfile());

            var image = Assert.Single(doc.Images);
            Assert.Equal("Company logo", image.Description);
            Assert.Equal("Quarterly report logo", image.Title);

            using MemoryStream stream = doc.ToStream();
            using var loaded = WordDocument.Load(stream);
            string roundTrip = loaded.ToHtml();

            Assert.Contains("alt=\"Company logo\"", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("title=\"Quarterly report logo\"", roundTrip, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HtmlToWord_UnreachableImage_InsertsPlaceholder() {
            string html = "<img src=\"http://localhost:1/missing.png\" alt=\"Missing\" />";
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
                .ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Empty(doc.Images);
            Assert.Single(doc.Paragraphs);
            Assert.Equal("Missing", doc.Paragraphs[0].Text);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("ImageSkippedByPolicy", diagnostic.Code);
            Assert.Equal(HtmlDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("http://localhost:1/missing.png", diagnostic.Source);
            Assert.Contains("data URI", diagnostic.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Null(diagnostic.Detail);
        }

        [Fact]
        public void HtmlToWord_UnreachableImage_NoAlt_SkipsImage() {
            string html = "<img src=\"http://localhost:1/missing.png\" />";
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
                .ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Empty(doc.Images);
            Assert.Empty(doc.Paragraphs);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("ImageSkippedByPolicy", diagnostic.Code);
            Assert.Equal("http://localhost:1/missing.png", diagnostic.Source);
        }

        [Fact]
        public async Task HtmlToWord_RemoteImageOverMaxBytes_SkipsWithDiagnostic() {
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ =>
                Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new ByteArrayContent(new byte[16])
            })));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed,
                MaxImageBytes = 4
            };
            string html = "<img src=\"https://example.test/large.png\" alt=\"Too large\" />";

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Empty(doc.Images);
            Assert.Single(doc.Paragraphs);
            Assert.Equal("Too large", doc.Paragraphs[0].Text);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("ImageResourceTooLarge", diagnostic.Code);
            Assert.Equal("https://example.test/large.png", diagnostic.Source);
            Assert.Contains("limit", diagnostic.Detail!, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task HtmlToWord_RemoteImageWithRejectedContentType_SkipsWithDiagnostic() {
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ => {
                var response = new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new ByteArrayContent(new byte[] { 1, 2, 3 })
                };
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("text/html");
                return Task.FromResult(response);
            }));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed
            };
            string html = "<img src=\"https://example.test/not-image.png\" alt=\"Not image\" />";

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Empty(doc.Images);
            Assert.Single(doc.Paragraphs);
            Assert.Equal("Not image", doc.Paragraphs[0].Text);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("ImageContentTypeRejected", diagnostic.Code);
            Assert.Equal("https://example.test/not-image.png", diagnostic.Source);
            Assert.Contains("text/html", diagnostic.Detail!, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HtmlToWord_RemoteImageWithRejectedScheme_SkipsBeforeFetchWithDiagnostic() {
            var fetched = false;
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ => {
                fetched = true;
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK));
            }));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed
            };
            options.AllowedImageUriSchemes.Remove(Uri.UriSchemeHttps);
            string html = "<img src=\"https://example.test/scheme.png\" alt=\"Blocked scheme\" />";

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.False(fetched);
            Assert.Empty(doc.Images);
            Assert.Single(doc.Paragraphs);
            Assert.Equal("Blocked scheme", doc.Paragraphs[0].Text);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("ImageResourceRejectedByPolicy", diagnostic.Code);
            Assert.Equal("https://example.test/scheme.png", diagnostic.Source);
            Assert.Contains("https", diagnostic.Detail!, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void HtmlToWord_UntrustedProfile_SkipsExternalImageBeforeFetch() {
            var fetched = false;
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ => {
                fetched = true;
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK));
            }));
            var options = HtmlToWordOptions.CreateUntrustedHtmlProfile();
            options.HttpClient = httpClient;
            string html = "<img src=\"https://example.test/untrusted.png\" alt=\"Blocked by profile\" />";

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.False(fetched);
            Assert.Empty(doc.Images);
            Assert.Single(doc.Paragraphs);
            Assert.Equal("Blocked by profile", doc.Paragraphs[0].Text);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("ImageSkippedByPolicy", diagnostic.Code);
            Assert.Equal("https://example.test/untrusted.png", diagnostic.Source);
        }

        [Fact]
        public void HtmlToWord_RemoteImageWithRejectedHost_SkipsBeforeFetchWithDiagnostic() {
            var fetched = false;
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ => {
                fetched = true;
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK));
            }));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed
            };
            options.AllowedImageHosts.Add("images.example.test");
            string html = "<img src=\"https://other.example.test/host.png\" alt=\"Blocked host\" />";

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.False(fetched);
            Assert.Empty(doc.Images);
            Assert.Single(doc.Paragraphs);
            Assert.Equal("Blocked host", doc.Paragraphs[0].Text);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("ImageResourceRejectedByPolicy", diagnostic.Code);
            Assert.Equal("https://other.example.test/host.png", diagnostic.Source);
            Assert.Contains("other.example.test", diagnostic.Detail!, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task HtmlToWord_PictureSourceSet_UsesFirstAllowedImageCandidate() {
            var requested = new List<Uri>();
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                requested.Add(request.RequestUri!);
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new ByteArrayContent(Convert.FromBase64String(validPng)) {
                        Headers = {
                            ContentType = new MediaTypeHeaderValue("image/png")
                        }
                    }
                });
            }));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed
            };
            options.AllowedImageHosts.Add("images.example.test");
            string html = """
<picture>
  <source srcset="https://blocked.example.test/bad.png 1x, https://images.example.test/good.png 2x">
  <img src="https://blocked.example.test/fallback.png" alt="Allowed image" />
</picture>
""";

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            var request = Assert.Single(requested);
            Assert.Equal("images.example.test", request.Host);
            Assert.Equal("/good.png", request.AbsolutePath);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public async Task HtmlToWord_PictureSource_UsesDataOriginalBeforeImageFallback() {
            var requested = new List<Uri>();
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                requested.Add(request.RequestUri!);
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new ByteArrayContent(Convert.FromBase64String(validPng)) {
                        Headers = {
                            ContentType = new MediaTypeHeaderValue("image/png")
                        }
                    }
                });
            }));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed
            };
            options.AllowedImageHosts.Add("images.example.test");
            string html = """
<picture>
  <source data-original="https://images.example.test/high.png">
  <img src="https://images.example.test/fallback.png" alt="Lazy picture" />
</picture>
""";

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            var request = Assert.Single(requested);
            Assert.Equal("images.example.test", request.Host);
            Assert.Equal("/high.png", request.AbsolutePath);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public async Task HtmlToWord_ImageSourceSet_UsesResponsiveCandidateBeforeSourceFallback() {
            var requested = new List<Uri>();
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                requested.Add(request.RequestUri!);
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new ByteArrayContent(Convert.FromBase64String(validPng)) {
                        Headers = {
                            ContentType = new MediaTypeHeaderValue("image/png")
                        }
                    }
                });
            }));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed
            };
            options.AllowedImageHosts.Add("images.example.test");
            string html = """<img src="https://images.example.test/fallback.png" srcset="https://images.example.test/hero.png 1x" alt="Responsive image" />""";

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            var request = Assert.Single(requested);
            Assert.Equal("images.example.test", request.Host);
            Assert.Equal("/hero.png", request.AbsolutePath);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public async Task HtmlToWord_ImageSelection_LimitsRemoteCandidateProbesPerElement() {
            var requested = new List<Uri>();
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                requested.Add(request.RequestUri!);
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed
            };
            string html = $"<img srcset=\"https://cdn.example.test/one.png 1x, https://cdn.example.test/two.png 2x, https://cdn.example.test/three.png 3x\" src=\"data:image/png;base64,{base64}\" alt=\"Logo\" />";

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Equal(new Uri("https://cdn.example.test/one.png"), Assert.Single(requested));
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public async Task HtmlToWord_ImageSelection_LimitsResponsiveCandidatesAndUsesSourceFallback() {
            var requested = new List<Uri>();
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                requested.Add(request.RequestUri!);
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed,
                MaxRemoteImageCandidateProbes = null,
                MaxImageSourceCandidates = 2
            };
            string html = $"<img srcset=\"https://cdn.example.test/one.png 1x, https://cdn.example.test/two.png 2x, https://cdn.example.test/three.png 3x\" src=\"data:image/png;base64,{base64}\" alt=\"Logo\" />";

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Equal(new[] { "/one.png", "/two.png" }, requested.Select(uri => uri.AbsolutePath).ToArray());
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public async Task HtmlToWord_ImageSelection_DedupesResponsiveCandidatesBeforeApplyingLimit() {
            var requested = new List<Uri>();
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                requested.Add(request.RequestUri!);
                if (request.RequestUri!.AbsolutePath == "/good.png") {
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                        Content = new ByteArrayContent(Convert.FromBase64String(validPng)) {
                            Headers = {
                                ContentType = new MediaTypeHeaderValue("image/png")
                            }
                        }
                    });
                }

                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed,
                MaxRemoteImageCandidateProbes = null,
                MaxImageSourceCandidates = 2
            };
            string html = """<img srcset="https://cdn.example.test/missing.png 1x, https://cdn.example.test/missing.png 2x, https://cdn.example.test/good.png 3x" alt="Logo" />""";

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Equal(new[] { "/missing.png", "/good.png" }, requested.Select(uri => uri.AbsolutePath).ToArray());
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public async Task HtmlToWord_ImageSelection_DedupesLazyAndResponsiveCandidatesBeforeProbing() {
            var requested = new List<Uri>();
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                requested.Add(request.RequestUri!);
                if (request.RequestUri!.AbsolutePath == "/good.png") {
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                        Content = new ByteArrayContent(Convert.FromBase64String(validPng)) {
                            Headers = {
                                ContentType = new MediaTypeHeaderValue("image/png")
                            }
                        }
                    });
                }

                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed,
                MaxRemoteImageCandidateProbes = 2,
                MaxImageSourceCandidates = 2
            };
            string html = """<img data-src="https://cdn.example.test/missing.png" srcset="https://cdn.example.test/missing.png 1x, https://cdn.example.test/good.png 2x" alt="Logo" />""";

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Equal(new[] { "/missing.png", "/good.png" }, requested.Select(uri => uri.AbsolutePath).ToArray());
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public async Task HtmlToWord_ImageSelection_AppliesPolicyBeforeResponsiveCandidateLimit() {
            var requested = new List<Uri>();
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                requested.Add(request.RequestUri!);
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new ByteArrayContent(Convert.FromBase64String(validPng)) {
                        Headers = {
                            ContentType = new MediaTypeHeaderValue("image/png")
                        }
                    }
                });
            }));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed,
                MaxRemoteImageCandidateProbes = null,
                MaxImageSourceCandidates = 2
            };
            options.AllowedImageHosts.Add("cdn.good.test");
            string html = """<img srcset="https://bad.example.test/one.png 1x, https://bad.example.test/two.png 2x, https://cdn.good.test/good.png 3x" alt="Logo" />""";

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            var request = Assert.Single(requested);
            Assert.Equal("cdn.good.test", request.Host);
            Assert.Equal("/good.png", request.AbsolutePath);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public void HtmlToWord_ImageSelection_PreservesAltTextWhenResponsiveSourcesAreRejected() {
            var options = new HtmlToWordOptions {
                ImageProcessing = ImageProcessingMode.Embed
            };
            options.AllowedImageHosts.Add("cdn.good.test");
            string html = """<img srcset="https://blocked.example.test/logo.png 1x" alt="Logo" />""";

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Empty(doc.Images);
            var paragraph = Assert.Single(doc.Paragraphs);
            Assert.Equal("Logo", paragraph.Text);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("ImageResourceRejectedByPolicy", diagnostic.Code);
        }

        [Fact]
        public async Task HtmlToWord_ImageSelection_DoesNotCountAbsentPictureAttributesTowardScanLimit() {
            var requested = new List<Uri>();
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                requested.Add(request.RequestUri!);
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new ByteArrayContent(Convert.FromBase64String(validPng)) {
                        Headers = {
                            ContentType = new MediaTypeHeaderValue("image/png")
                        }
                    }
                });
            }));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed,
                MaxRemoteImageCandidateProbes = null,
                MaxImageSourceCandidates = 1
            };
            string html = """
<picture>
  <source data-lazy-src="https://cdn.example.test/good.png">
  <img src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=" alt="Logo" />
</picture>
""";

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            var request = Assert.Single(requested);
            Assert.Equal("/good.png", request.AbsolutePath);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public async Task HtmlToWord_ImageSelection_DoesNotLetOverLimitSrcSetSuppressSourceFallback() {
            var requested = new List<Uri>();
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                requested.Add(request.RequestUri!);
                if (request.RequestUri!.AbsolutePath == "/fallback.png") {
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                        Content = new ByteArrayContent(Convert.FromBase64String(validPng)) {
                            Headers = {
                                ContentType = new MediaTypeHeaderValue("image/png")
                            }
                        }
                    });
                }

                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed,
                MaxRemoteImageCandidateProbes = null,
                MaxImageSourceCandidates = 1
            };
            string html = """
<picture>
  <source srcset="https://cdn.example.test/missing.png 1x">
  <img srcset="https://cdn.example.test/fallback.png 1x" src="https://cdn.example.test/fallback.png" alt="Logo" />
</picture>
""";

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Equal(new[] { "/missing.png", "/fallback.png" }, requested.Select(uri => uri.AbsolutePath).ToArray());
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public async Task HtmlToWord_ImageSelection_AllowsTrustedCallersToProbeAllRemoteCandidates() {
            var requested = new List<Uri>();
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                requested.Add(request.RequestUri!);
                if (request.RequestUri!.AbsolutePath == "/two.png") {
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                        Content = new ByteArrayContent(Convert.FromBase64String(validPng)) {
                            Headers = {
                                ContentType = new MediaTypeHeaderValue("image/png")
                            }
                        }
                    });
                }

                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed,
                MaxRemoteImageCandidateProbes = null
            };
            string html = """<img srcset="https://cdn.example.test/one.png 1x, https://cdn.example.test/two.png 2x" alt="Logo" />""";

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Equal(new[] { "/one.png", "/two.png" }, requested.Select(uri => uri.AbsolutePath).ToArray());
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public async Task HtmlToWord_RemoteImageOverTotalMaxBytes_SkipsWithDiagnostic() {
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ =>
                Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new ByteArrayContent(new byte[16])
                })));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed,
                MaxTotalImageBytes = 4
            };
            string html = "<img src=\"https://example.test/budget.png\" alt=\"Budget\" />";

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Empty(doc.Images);
            Assert.Single(doc.Paragraphs);
            Assert.Equal("Budget", doc.Paragraphs[0].Text);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("ImageResourceBudgetExceeded", diagnostic.Code);
            Assert.Equal("https://example.test/budget.png", diagnostic.Source);
            Assert.Contains("budget", diagnostic.Detail!, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task HtmlToWord_RemoteImageOverRemainingTotalMaxBytes_SkipsBeforeReadingBody() {
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ => {
                var response = new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new ThrowIfReadContent(128)
                };
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("image/png");
                return Task.FromResult(response);
            }));
            string html = $"<img src=\"data:image/png;base64,{validPng}\" alt=\"Valid\" /><img src=\"https://example.test/too-large.png\" alt=\"Too large\" />";
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed,
                MaxTotalImageBytes = 80
            };

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Contains(conversion.Report.Diagnostics, diagnostic =>
                diagnostic.Code == "ImageResourceBudgetExceeded" &&
                string.Equals(diagnostic.Source, "https://example.test/too-large.png", StringComparison.OrdinalIgnoreCase));
            Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "ImageLoadFailed");
        }

        [Fact]
        public void HtmlToWord_DataImageOverMaxBytes_SkipsBeforeDecodeWithDiagnostic() {
            var data = Convert.ToBase64String(new byte[16]);
            string html = $"<img src=\"data:image/png;base64,{data}\" alt=\"Too large data\" />";
            var options = new HtmlToWordOptions {
                MaxImageBytes = 4
            };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Empty(doc.Images);
            Assert.Single(doc.Paragraphs);
            Assert.Equal("Too large data", doc.Paragraphs[0].Text);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("ImageResourceTooLarge", diagnostic.Code);
            Assert.Equal("data:image", diagnostic.Source);
        }

        [Fact]
        public void HtmlToWord_DataImageOverTotalMaxBytes_SkipsBeforeDecodeWithDiagnostic() {
            var data = Convert.ToBase64String(new byte[16]);
            string html = $"<img src=\"data:image/png;base64,{data}\" alt=\"Too much data\" />";
            var options = new HtmlToWordOptions {
                MaxTotalImageBytes = 4
            };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Empty(doc.Images);
            Assert.Single(doc.Paragraphs);
            Assert.Equal("Too much data", doc.Paragraphs[0].Text);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("ImageResourceBudgetExceeded", diagnostic.Code);
            Assert.Equal("data:image", diagnostic.Source);
        }

        [Fact]
        public void HtmlToWord_InvalidDataImage_DoesNotConsumeTotalBudget() {
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            string html = $"<img src=\"data:image/png;base64,not-valid-base64\" alt=\"Broken\" /><img src=\"data:image/png;base64,{validPng}\" alt=\"Valid\" />";
            var options = new HtmlToWordOptions {
                MaxTotalImageBytes = 70
            };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Contains(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "ImageDataUriInvalid");
            Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "ImageResourceBudgetExceeded");
        }

        [Fact]
        public async Task HtmlToWord_InvalidRemoteImage_DoesNotConsumeTotalBudget() {
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ => {
                var response = new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new ByteArrayContent(Encoding.ASCII.GetBytes("not an image"))
                };
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("image/png");
                return Task.FromResult(response);
            }));
            string html = $"<img src=\"https://example.test/broken.png\" alt=\"Broken\" /><img src=\"data:image/png;base64,{validPng}\" alt=\"Valid\" />";
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed,
                MaxTotalImageBytes = 70
            };

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Contains(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "ImageLoadFailed");
            Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "ImageResourceBudgetExceeded");
        }

        [Fact]
        public void HtmlToWord_InvalidLocalImage_DoesNotConsumeTotalBudget() {
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".png");
            File.WriteAllText(path, "not an image");
            try {
                string html = $"<img src=\"{path}\" alt=\"Broken\" /><img src=\"data:image/png;base64,{validPng}\" alt=\"Valid\" />";
                var options = HtmlToWordOptions.CreateTrustedDocumentProfile();
                options.MaxTotalImageBytes = 70;

                HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
                    .ToWordDocumentResult(options);
                var doc = conversion.Value;

                Assert.Single(doc.Images);
                Assert.Contains(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "ImageLoadFailed");
                Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "ImageResourceBudgetExceeded");
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void HtmlToWord_InvalidSvgDataImage_DoesNotConsumeTotalBudget() {
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            var invalidSvgData = Convert.ToBase64String(Encoding.UTF8.GetBytes("<svg xmlns=\"http://www.w3.org/2000/svg\"><path></svg>"));
            string html = $"<img src=\"data:image/svg+xml;base64,{invalidSvgData}\" alt=\"Broken svg\" /><img src=\"data:image/png;base64,{validPng}\" alt=\"Valid\" />";
            var options = new HtmlToWordOptions {
                MaxTotalImageBytes = 100
            };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Contains(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "SvgEmbedFailed");
            Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "ImageResourceBudgetExceeded");
        }

        [Fact]
        public void HtmlToWord_InvalidLocalSvgImage_DoesNotConsumeTotalBudget() {
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".svg");
            File.WriteAllText(path, "<svg xmlns=\"http://www.w3.org/2000/svg\"><path></svg>");
            try {
                string html = $"<img src=\"{path}\" alt=\"Broken svg\" /><img src=\"data:image/png;base64,{validPng}\" alt=\"Valid\" />";
                var options = HtmlToWordOptions.CreateTrustedDocumentProfile();
                options.MaxTotalImageBytes = 100;

                HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
                    .ToWordDocumentResult(options);
                var doc = conversion.Value;

                Assert.Single(doc.Images);
                Assert.Contains(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "SvgEmbedFailed");
                Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "ImageResourceBudgetExceeded");
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public async Task HtmlToWord_InvalidRemoteSvgImage_DoesNotConsumeTotalBudget() {
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ => {
                var response = new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new ByteArrayContent(Encoding.UTF8.GetBytes("<svg xmlns=\"http://www.w3.org/2000/svg\"><path></svg>"))
                };
                response.Content.Headers.ContentType = new MediaTypeHeaderValue("image/svg+xml");
                return Task.FromResult(response);
            }));
            string html = $"<img src=\"https://example.test/broken.svg\" alt=\"Broken svg\" /><img src=\"data:image/png;base64,{validPng}\" alt=\"Valid\" />";
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed,
                MaxTotalImageBytes = 100
            };

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Contains(conversion.Report.Diagnostics, diagnostic =>
                diagnostic.Code == "SvgLoadFailed" &&
                diagnostic.Source == "https://example.test/broken.svg");
            Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "ImageResourceBudgetExceeded");
        }

        [Fact]
        public void HtmlToWord_InvalidInlineSvg_DoesNotConsumeTotalBudget() {
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            string html = $"<svg xmlns=\"http://www.w3.org/2000/svg\"><path></svg><img src=\"data:image/png;base64,{validPng}\" alt=\"Valid\" />";
            var options = new HtmlToWordOptions {
                MaxTotalImageBytes = 100
            };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Contains(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "InlineSvgEmbedFailed");
            Assert.DoesNotContain(conversion.Report.Diagnostics, diagnostic => diagnostic.Code == "ImageResourceBudgetExceeded");
        }

        [Fact]
        public void HtmlToWord_DataImageWithRejectedContentType_SkipsWithDiagnostic() {
            string html = "<img src=\"data:image/x-officeimo;base64,AAAA\" alt=\"Bad mime\" />";
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Empty(doc.Images);
            Assert.Single(doc.Paragraphs);
            Assert.Equal("Bad mime", doc.Paragraphs[0].Text);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("ImageContentTypeRejected", diagnostic.Code);
            Assert.Equal("data:image/x-officeimo", diagnostic.Source);
        }

        [Fact]
        public void HtmlToWord_SvgDataImageOverMaxBytes_SkipsBeforeEmbedWithDiagnostic() {
            var svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"10\" height=\"10\"><text>too large</text></svg>";
            var data = Convert.ToBase64String(Encoding.UTF8.GetBytes(svg));
            string html = $"<img src=\"data:image/svg+xml;base64,{data}\" alt=\"Too large svg\" />";
            var options = new HtmlToWordOptions {
                MaxImageBytes = 4
            };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Empty(doc.Images);
            Assert.Single(doc.Paragraphs);
            Assert.Equal("Too large svg", doc.Paragraphs[0].Text);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("ImageResourceTooLarge", diagnostic.Code);
            Assert.Equal("data:image/svg+xml", diagnostic.Source);
        }

        [Fact]
        public void HtmlToWord_SvgDataImageOverTotalMaxBytes_SkipsBeforeEmbedWithDiagnostic() {
            var svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"10\" height=\"10\"><text>budget</text></svg>";
            var data = Convert.ToBase64String(Encoding.UTF8.GetBytes(svg));
            string html = $"<img src=\"data:image/svg+xml;base64,{data}\" alt=\"Budget svg\" />";
            var options = new HtmlToWordOptions {
                MaxTotalImageBytes = 4
            };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Empty(doc.Images);
            Assert.Single(doc.Paragraphs);
            Assert.Equal("Budget svg", doc.Paragraphs[0].Text);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("ImageResourceBudgetExceeded", diagnostic.Code);
            Assert.Equal("data:image/svg+xml", diagnostic.Source);
        }

        [Fact]
        public void InlineImagePreservesTextOrder() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            var base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"<p>before<img src=\"data:image/png;base64,{base64}\">after</p>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
            Assert.Equal(3, doc.Paragraphs.Count);
            Assert.Equal("before", doc.Paragraphs[0].Text);
            Assert.NotNull(doc.Paragraphs[1].Image);
            Assert.Equal("after", doc.Paragraphs[2].Text);
        }

        [Fact]
        public void DuplicateImageSrcSharesPart() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            var base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            var dataUri = $"data:image/png;base64,{base64}";
            string html = $"<p><img src=\"{dataUri}\"/><img src=\"{dataUri}\"/></p>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
            Assert.Collection(doc.Images, _ => { }, _ => { });
            Assert.Equal(doc.Images[0].RelationshipId, doc.Images[1].RelationshipId);
            var wordDoc = doc._wordprocessingDocument;
            Assert.NotNull(wordDoc);
            var mainPart = wordDoc.MainDocumentPart;
            Assert.NotNull(mainPart);
            Assert.Single(mainPart.ImageParts);
        }

        [Fact]
        public void DuplicateImageSrcKeepsPerImageAltAndTitleMetadata() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            var base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            var dataUri = $"data:image/png;base64,{base64}";
            string html = $"<p><img src=\"{dataUri}\" alt=\"First logo\" title=\"First title\"/><img src=\"{dataUri}\" alt=\"Second logo\" title=\"Second title\"/></p>";

            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());

            Assert.Equal(2, doc.Images.Count);
            Assert.Equal(doc.Images[0].RelationshipId, doc.Images[1].RelationshipId);
            Assert.Equal("First logo", doc.Images[0].Description);
            Assert.Equal("First title", doc.Images[0].Title);
            Assert.Equal("Second logo", doc.Images[1].Description);
            Assert.Equal("Second title", doc.Images[1].Title);

            using MemoryStream stream = doc.ToStream();
            using var loaded = WordDocument.Load(stream);
            string roundTrip = loaded.ToHtml();

            Assert.Contains("alt=\"First logo\"", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("title=\"First title\"", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("alt=\"Second logo\"", roundTrip, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("title=\"Second title\"", roundTrip, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void DuplicateImageFileSrcSharesPart() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string html = $"<p><img src=\"{path}\"/><img src=\"{path}\"/></p>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
                .ToWordDocument(HtmlToWordOptions.CreateTrustedDocumentProfile());
            Assert.Equal(2, doc.Images.Count);
            Assert.Equal(doc.Images[0].RelationshipId, doc.Images[1].RelationshipId);
            var wordDoc = doc._wordprocessingDocument;
            Assert.NotNull(wordDoc);
            var mainPart = wordDoc.MainDocumentPart;
            Assert.NotNull(mainPart);
            Assert.Single(mainPart.ImageParts);
        }

        [Fact]
        public void ImageFloatLeftWrapsLeft() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            var base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"<img src=\"data:image/png;base64,{base64}\" style=\"float:left\"/>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
            var img = Assert.Single(doc.Images);
            Assert.Equal(WrapTextImage.Square, img.WrapText);
            var hPos = img.horizontalPosition;
            Assert.NotNull(hPos.HorizontalAlignment);
            Assert.Equal("left", hPos.HorizontalAlignment.Text);
        }

        [Fact]
        public void ImageFloatRightWrapsRight() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            var base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"<img src=\"data:image/png;base64,{base64}\" style=\"float:right\"/>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
            var img = Assert.Single(doc.Images);
            Assert.Equal(WrapTextImage.Square, img.WrapText);
            var hPos = img.horizontalPosition;
            Assert.NotNull(hPos.HorizontalAlignment);
            Assert.Equal("right", hPos.HorizontalAlignment.Text);
        }

        [Fact]
        public void HtmlToWord_ImagePercentWidth_UsesPageContentWidth() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            var base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"<img src=\"data:image/png;base64,{base64}\" style=\"width:50%\"/>";

            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());

            var img = Assert.Single(doc.Images);
            var section = doc.Sections.Last();
            var contentWidthTwips = (section.PageSettings.Width?.Value ?? WordPageSizes.A4.Width!.Value)
                - (section.Margins.Left?.Value ?? 1440U)
                - (section.Margins.Right?.Value ?? 1440U);
            var expectedWidthPixels = contentWidthTwips / 15D * 0.5D;
            Assert.NotNull(img.Width);
            Assert.InRange(img.Width!.Value, expectedWidthPixels - 0.5D, expectedWidthPixels + 0.5D);
        }

        [Fact]
        public void HtmlToWord_ImageWidthOnly_PreservesRequestedWidth() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            var base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"<img src=\"data:image/png;base64,{base64}\" width=\"64\"/>";

            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());

            var img = Assert.Single(doc.Images);
            Assert.NotNull(img.Width);
            Assert.Equal(64D, Math.Round(img.Width!.Value));
            Assert.NotNull(img.Height);
            Assert.True(img.Height!.Value > 0);
        }

        [Fact]
        public void HtmlToWord_ImageDimensionsPersistInSavedDrawingExtent() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string html = $"<img src=\"{path}\" width=\"64\" height=\"32\" alt=\"Logo\" />";

            using var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
                .ToWordDocument(HtmlToWordOptions.CreateTrustedDocumentProfile());

            var img = Assert.Single(doc.Images);
            Assert.Equal(64D, Math.Round(img.Width!.Value));
            Assert.Equal(32D, Math.Round(img.Height!.Value));

            using MemoryStream stream = doc.ToStream();
            using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
            var errors = new OpenXmlValidator().Validate(package).ToList();
            Assert.True(errors.Count == 0, OpenXmlValidationFormatting.FormatValidationErrors(errors));

            var drawing = Assert.Single(package.MainDocumentPart!.Document.Body!.Descendants<DocumentFormat.OpenXml.Wordprocessing.Drawing>());
            var extent = drawing.Inline!.Extent!;
            Assert.Equal(64L * 9525L, extent.Cx!.Value);
            Assert.Equal(32L * 9525L, extent.Cy!.Value);
        }

        [Fact]
        public void HtmlToWord_ImageProcessing_LinkExternal_UsesExternalRelationship() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            var uri = new Uri(path).AbsoluteUri;
            string html = $"<img src=\"{uri}\" width=\"64\" height=\"64\" alt=\"Logo\" />";
            var options = HtmlToWordOptions.CreateTrustedDocumentProfile();
            options.ImageProcessing = ImageProcessingMode.LinkExternal;
            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
                .ToWordDocumentResult(options);
            var doc = conversion.Value;
            var img = Assert.Single(doc.Images);
            Assert.True(img.IsExternal);
            Assert.Equal(new Uri(uri), img.ExternalUri);
            var mainPart = doc._wordprocessingDocument?.MainDocumentPart;
            Assert.NotNull(mainPart);
            Assert.Empty(mainPart!.ImageParts);
        }

        [Fact]
        public void HtmlToWord_ImageProcessing_LinkExternal_RejectsDisallowedHost() {
            string html = "<img src=\"https://blocked.example.test/logo.png\" width=\"64\" height=\"64\" alt=\"Blocked link\" />";
            var options = new HtmlToWordOptions { ImageProcessing = ImageProcessingMode.LinkExternal };
            options.AllowedImageHosts.Add("images.example.test");

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Empty(doc.Images);
            Assert.Single(doc.Paragraphs);
            Assert.Equal("Blocked link", doc.Paragraphs[0].Text);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("ImageResourceRejectedByPolicy", diagnostic.Code);
            Assert.Equal("https://blocked.example.test/logo.png", diagnostic.Source);
        }

        [Fact]
        public void HtmlToWord_ImageProcessing_LinkExternal_ContinuesPastExternalCandidateWithoutDimensions() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"<img src=\"data:image/png;base64,{base64}\" srcset=\"https://cdn.example.test/logo.png 1x\" alt=\"Logo\" />";
            var options = new HtmlToWordOptions { ImageProcessing = ImageProcessingMode.LinkExternal };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            var image = Assert.Single(doc.Images);
            Assert.False(image.IsExternal);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public void HtmlToWord_ImageProcessing_LinkExternal_ContinuesPastCssOnlyExternalCandidate() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"<img src=\"data:image/png;base64,{base64}\" srcset=\"https://cdn.example.test/logo.png 1x\" style=\"width:32px;height:32px\" alt=\"Logo\" />";
            var options = new HtmlToWordOptions { ImageProcessing = ImageProcessingMode.LinkExternal };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            var image = Assert.Single(doc.Images);
            Assert.False(image.IsExternal);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public void HtmlToWord_Picture_SkipsUnsupportedSourceTypeBeforeFallbackImage() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"""
<picture>
  <source type="image/avif" srcset="https://cdn.example.test/logo.avif 1x" />
  <img src="data:image/png;base64,{base64}" alt="Logo" />
</picture>
""";
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public void HtmlToWord_Picture_AllowsSourceTypeWithParametersBeforeFallbackImage() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"""
<picture>
  <source type="image/png; charset=binary" srcset="data:image/png;base64,{base64}" />
  <img src="https://cdn.example.test/logo.png" alt="Logo" />
</picture>
""";
            var options = new HtmlToWordOptions { ImageProcessing = ImageProcessingMode.EmbedDataUriOnly };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public void HtmlToWord_DataImage_DecodesPercentEscapedBase64Payload() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string base64 = Uri.EscapeDataString(Convert.ToBase64String(File.ReadAllBytes(path)));
            string html = $"<img src=\"data:image/png;base64,{base64}\" alt=\"Logo\" />";
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public void HtmlToWord_ImageSelection_ContinuesPastUnsupportedDataCandidate() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"<img data-src=\"data:image/avif;base64,AQID\" src=\"data:image/png;base64,{base64}\" alt=\"Logo\" />";
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public void HtmlToWord_ImageSelection_ContinuesPastOversizedDataCandidate() {
            string oversized = Convert.ToBase64String(new byte[16]);
            const string fallback = "https://cdn.example.test/logo.png";
            string html = $"<img data-src=\"data:image/png;base64,{oversized}\" src=\"{fallback}\" width=\"32\" height=\"32\" alt=\"Logo\" />";
            var options = new HtmlToWordOptions {
                ImageProcessing = ImageProcessingMode.LinkExternal,
                MaxImageBytes = 4
            };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            var image = Assert.Single(doc.Images);
            Assert.True(image.IsExternal);
            Assert.Equal(new Uri(fallback), image.ExternalUri);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public void HtmlToWord_ImageSelection_ContinuesPastUndecodableDataCandidate() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"<img data-src=\"data:image/png;base64,not-valid\" src=\"data:image/png;base64,{base64}\" alt=\"Logo\" />";
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public void HtmlToWord_ImageSelection_ContinuesPastInvalidDataImageCandidate() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"<img data-src=\"data:image/png;base64,AQID\" src=\"data:image/png;base64,{base64}\" alt=\"Logo\" />";
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public void HtmlToWord_ImageSelection_ContinuesPastInvalidTextSvgDataCandidate() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string invalidSvg = Uri.EscapeDataString("<svg><path></svg>");
            string html = $"<img data-src=\"data:image/svg+xml,{invalidSvg}\" src=\"data:image/png;base64,{base64}\" alt=\"Logo\" />";
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public void HtmlToWord_ImageSelection_ContinuesPastSvgDataCandidateWithNonSvgPayload() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"<img data-src=\"data:image/svg+xml;base64,{base64}\" src=\"data:image/png;base64,{base64}\" alt=\"Logo\" />";
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public void HtmlToWord_ImageSelection_ContinuesPastDataCandidateWithDisallowedDetectedPayloadType() {
            string svg = Convert.ToBase64String(Encoding.UTF8.GetBytes("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"1\" height=\"1\"></svg>"));
            const string fallback = "https://cdn.example.test/logo.png";
            string html = $"<img data-src=\"data:image/png;base64,{svg}\" src=\"{fallback}\" width=\"32\" height=\"32\" alt=\"Logo\" />";
            var options = new HtmlToWordOptions {
                ImageProcessing = ImageProcessingMode.LinkExternal
            };
            options.AllowedImageContentTypes.Remove("image/svg+xml");

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            var image = Assert.Single(doc.Images);
            Assert.True(image.IsExternal);
            Assert.Equal(new Uri(fallback), image.ExternalUri);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public void HtmlToWord_ImageSelection_ContinuesPastNonImageDataCandidate() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"<img data-src=\"data:text/plain,placeholder\" src=\"data:image/png;base64,{base64}\" alt=\"Logo\" />";
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public async Task HtmlToWord_ImageSelection_ContinuesPastProtocolRelativeCandidateWithoutBase() {
            var requested = new List<Uri>();
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                requested.Add(request.RequestUri!);
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }));
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"<img srcset=\"//cdn.example.test/missing.png 1x\" src=\"data:image/png;base64,{base64}\" alt=\"Logo\" />";
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed
            };

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            var request = Assert.Single(requested);
            Assert.Equal("https://cdn.example.test/missing.png", request.AbsoluteUri);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public void HtmlToWord_ImageSelection_ContinuesPastUnresolvedRelativeCandidate() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"<img data-src=\"missing.png\" src=\"data:image/png;base64,{base64}\" alt=\"Logo\" />";
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public void HtmlToWord_ImageSelection_PreservesAltTextForUnresolvedRelativeCandidateWithoutFallback() {
            const string html = "<img src=\"missing.png\" alt=\"Missing\" />";
            var options = new HtmlToWordOptions();

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Empty(doc.Images);
            Assert.Single(doc.Paragraphs);
            Assert.Equal("Missing", doc.Paragraphs[0].Text);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("ImageSkippedByPolicy", diagnostic.Code);
            Assert.Equal("missing.png", diagnostic.Source);
        }

        [Fact]
        public async Task HtmlToWord_ImageSelection_ContinuesPastFailedRemoteCandidate() {
            var requested = new List<Uri>();
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                requested.Add(request.RequestUri!);
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed
            };
            string html = $"<img srcset=\"https://cdn.example.test/missing.png 1x\" src=\"data:image/png;base64,{base64}\" alt=\"Logo\" />";

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Equal(new Uri("https://cdn.example.test/missing.png"), Assert.Single(requested));
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public async Task HtmlToWord_ImageSelection_DoesNotRetryFailedRemoteCandidateWithoutFallback() {
            var requested = new List<Uri>();
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                requested.Add(request.RequestUri!);
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
            }));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed
            };
            string html = """<img src="https://cdn.example.test/missing.png" alt="Missing" />""";

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Empty(doc.Images);
            Assert.Equal(new Uri("https://cdn.example.test/missing.png"), Assert.Single(requested));
        }

        [Fact]
        public async Task HtmlToWord_ImageSelection_ContinuesPastRemoteProbeTimeout() {
            var requested = new List<Uri>();
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                requested.Add(request.RequestUri!);
                using var cts = new CancellationTokenSource();
                cts.Cancel();
                return Task.FromCanceled<HttpResponseMessage>(cts.Token);
            }));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed,
                ResourceTimeout = TimeSpan.FromMilliseconds(1)
            };
            string html = $"<img data-src=\"https://cdn.example.test/slow.png\" src=\"data:image/png;base64,{base64}\" alt=\"Logo\" />";

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Equal(new Uri("https://cdn.example.test/slow.png"), Assert.Single(requested));
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public void HtmlToWord_ImageSelection_ContinuesPastInvalidLocalCandidate() {
            string badPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".png");
            File.WriteAllText(badPath, "not an image");
            try {
                var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
                string base64 = Convert.ToBase64String(File.ReadAllBytes(path));
                string html = $"<img data-src=\"{new Uri(badPath).AbsoluteUri}\" src=\"data:image/png;base64,{base64}\" alt=\"Logo\" />";
                var options = HtmlToWordOptions.CreateTrustedDocumentProfile();

                HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
                var doc = conversion.Value;

                Assert.Single(doc.Images);
                Assert.Empty(conversion.Report.Diagnostics);
            } finally {
                File.Delete(badPath);
            }
        }

        [Fact]
        public void HtmlToWord_ImageSelection_ContinuesPastOversizedLocalCandidate() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            byte[] imageBytes = File.ReadAllBytes(path);
            string oversizedPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".png");
            File.WriteAllBytes(oversizedPath, imageBytes.Concat(new byte[16]).ToArray());
            try {
                string base64 = Convert.ToBase64String(imageBytes);
                string html = $"<img data-src=\"{new Uri(oversizedPath).AbsoluteUri}\" src=\"data:image/png;base64,{base64}\" alt=\"Logo\" />";
                var options = HtmlToWordOptions.CreateTrustedDocumentProfile();
                options.MaxImageBytes = imageBytes.LongLength;

                HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
                var doc = conversion.Value;

                Assert.Single(doc.Images);
                Assert.Empty(conversion.Report.Diagnostics);
            } finally {
                File.Delete(oversizedPath);
            }
        }

        [Fact]
        public void HtmlToWord_ImageSelection_SkipsLocalFileCandidateByDefault() {
            string localPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".png");
            File.WriteAllText(localPath, "not an image");
            try {
                var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
                string base64 = Convert.ToBase64String(File.ReadAllBytes(path));
                string html = $"<img data-src=\"{new Uri(localPath).AbsoluteUri}\" src=\"data:image/png;base64,{base64}\" alt=\"Logo\" />";
                var options = new HtmlToWordOptions();

                HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
                var doc = conversion.Value;

                Assert.Single(doc.Images);
                Assert.Empty(conversion.Report.Diagnostics);
            } finally {
                File.Delete(localPath);
            }
        }

        [Fact]
        public void HtmlToWord_ImageSelection_ReusesCachedDataCandidatePastTotalBudget() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            byte[] bytes = File.ReadAllBytes(path);
            string base64 = Convert.ToBase64String(bytes);
            string dataUri = $"data:image/png;base64,{base64}";
            string html = $"<img src=\"{dataUri}\" alt=\"Logo\" /><img data-src=\"{dataUri}\" src=\"https://cdn.example.test/logo.png\" width=\"32\" height=\"32\" alt=\"Logo\" />";
            var options = new HtmlToWordOptions {
                ImageProcessing = ImageProcessingMode.LinkExternal,
                MaxTotalImageBytes = bytes.LongLength
            };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Equal(2, doc.Images.Count);
            Assert.All(doc.Images, image => Assert.False(image.IsExternal));
            Assert.Empty(conversion.Report.Diagnostics);
        }

        [Fact]
        public async Task HtmlToWord_RemoteImageCache_ReservesRepeatedFloatedImageParts() {
            var requested = new List<Uri>();
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            byte[] bytes = Convert.FromBase64String(validPng);
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                requested.Add(request.RequestUri!);
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new ByteArrayContent(bytes) {
                        Headers = {
                            ContentType = new MediaTypeHeaderValue("image/png")
                        }
                    }
                });
            }));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed,
                MaxTotalImageBytes = bytes.LongLength
            };
            string html = """
<img src="https://example.test/logo.png" style="float:left" alt="One" />
<img src="https://example.test/logo.png" style="float:left" alt="Two" />
""";

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Single(requested);
            Assert.Contains(conversion.Report.Diagnostics, diagnostic =>
                diagnostic.Code == "ImageResourceBudgetExceeded" &&
                diagnostic.Source == "https://example.test/logo.png");
        }

        [Fact]
        public async Task HtmlToWord_RemoteImageCache_DoesNotRetainRejectedCandidateBytes() {
            var requested = new List<Uri>();
            byte[] invalidBytes = Encoding.UTF8.GetBytes("not an image");
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                requested.Add(request.RequestUri!);
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new ByteArrayContent(invalidBytes) {
                        Headers = {
                            ContentType = new MediaTypeHeaderValue("image/png")
                        }
                    }
                });
            }));
            string html = $"""
<img data-src="https://cdn.example.test/bad-one.png" src="data:image/png;base64,{validPng}" alt="One" />
<img data-src="https://cdn.example.test/bad-two.png" src="data:image/png;base64,{validPng}" alt="Two" />
""";
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed
            };
            var converter = new HtmlToWordConverter();
            OfficeIMO.Html.HtmlConversionDocument source = OfficeIMO.Html.HtmlConversionDocument.Parse(html);

            using var doc = await converter.ConvertAsync(
                source.CreateDocumentForConversion(),
                options);

            Assert.Equal(2, doc.Images.Count);
            Assert.Equal(2, requested.Count);
            var field = typeof(HtmlToWordConverter).GetField("_remoteImageBytesCache", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            var cache = Assert.IsAssignableFrom<System.Collections.IDictionary>(field!.GetValue(converter));
            Assert.DoesNotContain(cache.Keys.Cast<string>(), key => key.Contains("bad-", StringComparison.Ordinal));
        }

        [Fact]
        public async Task HtmlToWord_RemoteImageCache_UsesCaseSensitiveUrlKeys() {
            var requested = new List<Uri>();
            const string validPng = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
            byte[] bytes = Convert.FromBase64String(validPng);
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                requested.Add(request.RequestUri!);
                if (request.RequestUri?.AbsolutePath == "/Logo.png") {
                    return Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound));
                }

                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new ByteArrayContent(bytes) {
                        Headers = {
                            ContentType = new MediaTypeHeaderValue("image/png")
                        }
                    }
                });
            }));
            var options = new HtmlToWordOptions {
                HttpClient = httpClient,
                ImageProcessing = ImageProcessingMode.Embed
            };
            string html = """
<img src="https://cdn.example.test/Logo.png" alt="Missing" />
<img src="https://cdn.example.test/logo.png" alt="Logo" />
""";

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Equal(2, requested.Count);
            Assert.Contains(requested, uri => uri.AbsolutePath == "/Logo.png");
            Assert.Contains(requested, uri => uri.AbsolutePath == "/logo.png");
            Assert.Contains(conversion.Report.Diagnostics, diagnostic =>
                diagnostic.Code == "ImageLoadFailed" &&
                diagnostic.Source == "https://cdn.example.test/Logo.png");
        }

        [Fact]
        public void HtmlToWord_ImageProcessing_EmbedDataUriOnly_SkipsExternalImages() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            var uri = new Uri(path).AbsoluteUri;
            string html = $"<img src=\"{uri}\" alt=\"Logo\" />";
            var options = new HtmlToWordOptions { ImageProcessing = ImageProcessingMode.EmbedDataUriOnly };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html, HtmlConversionDocumentOptions.CreateTrustedProfile())
                .ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Empty(doc.Images);
            Assert.Single(doc.Paragraphs);
            Assert.Equal("Logo", doc.Paragraphs[0].Text);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("ImageSkippedByPolicy", diagnostic.Code);
            Assert.Equal(uri, diagnostic.Source);
        }

        [Fact]
        public void HtmlToWord_ImageProcessing_EmbedDataUriOnly_ContinuesToDataUriCandidate() {
            var path = Path.Combine(AppContext.BaseDirectory, "Images", "EvotecLogo.png");
            string base64 = Convert.ToBase64String(File.ReadAllBytes(path));
            string html = $"<img data-src=\"https://cdn.example.test/logo.png\" src=\"data:image/png;base64,{base64}\" alt=\"Logo\" />";
            var options = new HtmlToWordOptions { ImageProcessing = ImageProcessingMode.EmbedDataUriOnly };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Single(doc.Images);
            Assert.Empty(conversion.Report.Diagnostics);
        }

        private sealed class FakeHtmlHttpMessageHandler : HttpMessageHandler {
            private readonly Func<HttpRequestMessage, Task<HttpResponseMessage>> _handler;

            internal FakeHtmlHttpMessageHandler(Func<HttpRequestMessage, Task<HttpResponseMessage>> handler) {
                _handler = handler;
            }

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken) {
                return _handler(request);
            }
        }

        private sealed class ThrowIfReadContent : HttpContent {
            private readonly long _length;

            internal ThrowIfReadContent(long length) {
                _length = length;
            }

            protected override Task SerializeToStreamAsync(Stream stream, TransportContext? context) {
                throw new InvalidOperationException("Content body should not be read when headers exceed the remaining image byte budget.");
            }

            protected override bool TryComputeLength(out long length) {
                length = _length;
                return true;
            }
        }
    }
}
