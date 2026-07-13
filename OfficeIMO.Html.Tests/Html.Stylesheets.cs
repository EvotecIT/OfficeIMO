using System;
using OfficeIMO.Word.Html;
using Xunit;
using System.Net;
using System.Net.Http;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Word;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_StyleElement_AppliesToMultipleParagraphs() {
            string html = "<style>p { color:#ff0000; }</style><p>First</p><p>Second</p>";
            var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions());
            var run1 = doc.Paragraphs[0].GetRuns().First();
            var run2 = doc.Paragraphs[1].GetRuns().First();
            Assert.Equal("FF0000", run1.ColorHex);
            Assert.Equal("FF0000", run2.ColorHex);
        }

        [Fact]
        public void HtmlToWord_LinkStylesheet_AppliesToMultipleParagraphs() {
            var path = Path.GetTempFileName();
            File.WriteAllText(path, "p { color:#00ff00; }");
            string html = $"<link rel=\"stylesheet\" href=\"{path}\" /><p>One</p><p>Two</p>";
            try {
                var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions { AllowDocumentStylesheetLinks = true });
                var run1 = doc.Paragraphs[0].GetRuns().First();
                var run2 = doc.Paragraphs[1].GetRuns().First();
                Assert.Equal("00FF00", run1.ColorHex);
                Assert.Equal("00FF00", run2.ColorHex);
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void HtmlToWord_OptionsStylesheet_AppliesToMultipleParagraphs() {
            string html = "<p>First</p><p>Second</p>";
            var options = new HtmlToWordOptions();
            options.StylesheetContents.Add("p { color:#0000ff; }");
            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;
            var run1 = doc.Paragraphs[0].GetRuns().First();
            var run2 = doc.Paragraphs[1].GetRuns().First();
            Assert.Equal("0000FF", run1.ColorHex);
            Assert.Equal("0000FF", run2.ColorHex);
        }

        [Fact]
        public void HtmlToWord_RemoteStylesheet_Applies() {
            // Deterministic version of the remote stylesheet test: inject the stylesheet content via options
            // (AngleSharp may not be able to fetch on some CI runners, and HttpListener is not supported everywhere.)
            string html = "<link rel=\\\"stylesheet\\\" href=\\\"https://example.com/style.css\\\" /><p>Test</p>";
            var options = new HtmlToWordOptions();
            options.StylesheetContents.Add("p { color:#123456; }");
            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;
            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.Equal("123456", run.ColorHex);
            string roundTrip = doc.ToHtml();
            Assert.Contains("<p>Test</p>", roundTrip, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task HtmlToWord_RemoteStylesheet_UsesConfiguredHttpClient() {
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                Assert.Equal(new Uri("https://styles.example.test/site.css"), request.RequestUri);
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new StringContent("p { color:#abcdef; }", Encoding.UTF8, "text/css")
                });
            }));
            string html = "<link rel=\"stylesheet\" href=\"https://styles.example.test/site.css\" /><p>Remote CSS</p>";
            var options = new HtmlToWordOptions {
                AllowDocumentStylesheetLinks = true,
                HttpClient = httpClient
            };

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.Equal("ABCDEF", run.ColorHex);
        }

        [Theory]
        [InlineData("body")]
        [InlineData("header")]
        [InlineData("footer")]
        public async Task HtmlToWord_AppendResolvesConfiguredStylesheetsFromTheSharedDocumentBaseUri(string target) {
            Uri? requestedUri = null;
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(request => {
                requestedUri = request.RequestUri;
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new StringContent("p { color:#abcdef; }", Encoding.UTF8, "text/css")
                });
            }));
            OfficeIMO.Html.HtmlConversionDocument source = OfficeIMO.Html.HtmlConversionDocument.Parse(
                "<p>Base-aware append</p>",
                new OfficeIMO.Html.HtmlConversionDocumentOptions {
                    BaseUri = new Uri("https://styles.example.test/reports/2026/page.html")
                });
            var options = new HtmlToWordOptions { HttpClient = httpClient };
            options.StylesheetPaths.Add("../site.css");
            options.AllowedStylesheetHosts.Add("styles.example.test");
            using WordDocument document = WordDocument.Create();

            switch (target) {
                case "body":
                    await document.AddHtmlToBodyAsync(source, options);
                    break;
                case "header":
                    await document.AddHtmlToHeaderAsync(source, options: options);
                    break;
                case "footer":
                    await document.AddHtmlToFooterAsync(source, options: options);
                    break;
                default:
                    throw new ArgumentOutOfRangeException(nameof(target));
            }

            Assert.Equal(new Uri("https://styles.example.test/reports/site.css"), requestedUri);
        }

        [Fact]
        public void HtmlToWord_UntrustedProfile_SkipsDocumentStylesheetLinkBeforeFetch() {
            var fetched = false;
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ => {
                fetched = true;
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new StringContent("p { color:#abcdef; }", Encoding.UTF8, "text/css")
                });
            }));
            string html = "<link rel=\"stylesheet\" href=\"https://styles.example.test/untrusted.css\" /><p>Remote CSS</p>";
            var options = HtmlToWordOptions.CreateUntrustedHtmlProfile();
            options.HttpClient = httpClient;

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.False(fetched);
            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.NotEqual("ABCDEF", run.ColorHex);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics, item => item.Code == "HtmlStylesheetLinkSkipped");
            Assert.Equal("https://styles.example.test/untrusted.css", diagnostic.Source);
        }

        [Fact]
        public void HtmlToWord_RemoteStylesheet_DisallowedHost_EmitsDiagnosticAndSkipsFetch() {
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ => throw new InvalidOperationException("Disallowed stylesheet host should not be fetched.")));
            string html = "<link rel=\"stylesheet\" href=\"https://blocked.example.test/site.css\" /><p>Remote CSS</p>";
            var options = new HtmlToWordOptions {
                AllowDocumentStylesheetLinks = true,
                HttpClient = httpClient
            };
            options.AllowedStylesheetHosts.Add("styles.example.test");

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.NotEqual("ABCDEF", run.ColorHex);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("StylesheetResourceRejectedByPolicy", diagnostic.Code);
            Assert.Equal("https://blocked.example.test/site.css", diagnostic.Source);
        }

        [Fact]
        public async Task HtmlToWord_RemoteStylesheet_MaxCssBytes_StopsBeforeReadingContent() {
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                Content = new ThrowIfReadContent(1024)
            })));
            string html = "<link rel=\"stylesheet\" href=\"https://styles.example.test/max.css\" /><p>Remote CSS</p>";
            var options = new HtmlToWordOptions {
                AllowDocumentStylesheetLinks = true,
                HttpClient = httpClient,
                MaxCssBytes = 8
            };

            var exception = await Assert.ThrowsAsync<HtmlConversionLimitException>(() => OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentAsync(options));

            Assert.Equal("CssSizeLimitExceeded", exception.Code);
            Assert.Equal("https://styles.example.test/max.css", exception.LimitSource);
            Assert.Equal(8, exception.Limit);
            Assert.Equal(1024, exception.Actual);
        }

        [Fact]
        public async Task HtmlToWord_RemoteStylesheet_MaxTotalCssBytes_StopsBeforeReadingContent() {
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                Content = new ThrowIfReadContent(1024)
            })));
            string seededCss = "body { color:#111111; }";
            string html = "<link rel=\"stylesheet\" href=\"https://styles.example.test/total.css\" /><p>Remote CSS</p>";
            var options = new HtmlToWordOptions {
                AllowDocumentStylesheetLinks = true,
                HttpClient = httpClient,
                MaxTotalCssBytes = Encoding.UTF8.GetByteCount(seededCss) + 8
            };
            options.StylesheetContents.Add(seededCss);

            var exception = await Assert.ThrowsAsync<HtmlConversionLimitException>(() => OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentAsync(options));

            Assert.Equal("CssTotalSizeLimitExceeded", exception.Code);
            Assert.Equal("https://styles.example.test/total.css", exception.LimitSource);
            Assert.Equal(options.MaxTotalCssBytes, exception.Limit);
            Assert.True(exception.Actual > exception.Limit);
        }

        [Fact]
        public async Task HtmlToWord_RemoteStylesheet_DisallowedContentType_EmitsDiagnosticAndSkipsStylesheet() {
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                Content = new StringContent("p { color:#fedcba; }", Encoding.UTF8, "application/json")
            })));
            string html = "<link rel=\"stylesheet\" href=\"https://styles.example.test/rejected-content-type.css\" /><p>Remote CSS</p>";
            var options = new HtmlToWordOptions {
                AllowDocumentStylesheetLinks = true,
                HttpClient = httpClient
            };

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.NotEqual("FEDCBA", run.ColorHex);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("StylesheetContentTypeRejected", diagnostic.Code);
            Assert.Equal("https://styles.example.test/rejected-content-type.css", diagnostic.Source);
        }

        [Fact]
        public async Task HtmlToWord_RemoteStylesheet_ContentTypeValidationCanBeDisabled() {
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ => Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                Content = new StringContent("p { color:#112233; }", Encoding.UTF8, "application/json")
            })));
            string html = "<link rel=\"stylesheet\" href=\"https://styles.example.test/json.css\" /><p>Remote CSS</p>";
            var options = new HtmlToWordOptions {
                AllowDocumentStylesheetLinks = true,
                HttpClient = httpClient,
                ValidateStylesheetContentTypes = false
            };

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.Equal("112233", run.ColorHex);
        }

        [Fact]
        public async Task HtmlToWord_RemoteStylesheet_NonSuccessStatus_EmitsSpecificDiagnosticAndSkipsStylesheet() {
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ => Task.FromResult(new HttpResponseMessage(HttpStatusCode.NotFound) {
                Content = new StringContent("p { color:#445566; }", Encoding.UTF8, "text/css")
            })));
            string html = "<link rel=\"stylesheet\" href=\"https://styles.example.test/not-found.css\" /><p>Remote CSS</p>";
            var options = new HtmlToWordOptions {
                AllowDocumentStylesheetLinks = true,
                HttpClient = httpClient
            };

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.NotEqual("445566", run.ColorHex);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("StylesheetHttpStatusRejected", diagnostic.Code);
            Assert.Equal("https://styles.example.test/not-found.css", diagnostic.Source);
            Assert.Contains("404", diagnostic.Detail, StringComparison.Ordinal);
        }

        [Fact]
        public async Task HtmlToWord_RemoteStylesheet_TransportFailure_EmitsSpecificDiagnosticAndSkipsStylesheet() {
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ =>
                Task.FromException<HttpResponseMessage>(new HttpRequestException("DNS resolution failed."))));
            string html = "<link rel=\"stylesheet\" href=\"https://styles.example.test/transport.css\" /><p>Remote CSS</p>";
            var options = new HtmlToWordOptions {
                AllowDocumentStylesheetLinks = true,
                HttpClient = httpClient
            };

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.NotEqual("778899", run.ColorHex);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("StylesheetTransportFailed", diagnostic.Code);
            Assert.Equal("https://styles.example.test/transport.css", diagnostic.Source);
            Assert.Contains("DNS resolution failed", diagnostic.Detail, StringComparison.Ordinal);
        }

        [Fact]
        public async Task HtmlToWord_RemoteStylesheet_ResourceTimeout_EmitsSpecificDiagnosticAndSkipsStylesheet() {
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ =>
                Task.FromException<HttpResponseMessage>(new TaskCanceledException("The stylesheet request timed out."))));
            string html = "<link rel=\"stylesheet\" href=\"https://styles.example.test/timeout.css\" /><p>Remote CSS</p>";
            var options = new HtmlToWordOptions {
                AllowDocumentStylesheetLinks = true,
                HttpClient = httpClient,
                ResourceTimeout = TimeSpan.FromMilliseconds(1)
            };

            HtmlToWordResult conversion = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResultAsync(options);
            var doc = conversion.Value;

            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.NotEqual("AABBCC", run.ColorHex);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("StylesheetLoadTimedOut", diagnostic.Code);
            Assert.Equal("https://styles.example.test/timeout.css", diagnostic.Source);
            Assert.Contains("timed out", diagnostic.Detail, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public async Task HtmlToWord_RemoteStylesheet_CallerCancellationStillThrows() {
            using var cts = new CancellationTokenSource();
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ => {
                cts.Cancel();
                return Task.FromException<HttpResponseMessage>(new OperationCanceledException(cts.Token));
            }));
            string html = "<link rel=\"stylesheet\" href=\"https://styles.example.test/caller-cancel.css\" /><p>Remote CSS</p>";
            var options = new HtmlToWordOptions {
                AllowDocumentStylesheetLinks = true,
                HttpClient = httpClient
            };

            await Assert.ThrowsAsync<OperationCanceledException>(() => OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentAsync(options, cts.Token));
        }

        [Fact]
        public void HtmlToWord_RelativeStylesheet_UsesBaseUrl() {
            var dir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(dir);
            var cssPath = Path.Combine(dir, "style.css");
            File.WriteAllText(cssPath, "p { color:#654321; }");
            try {
                var baseHref = new Uri(new Uri(Path.Combine(dir, "dummy"), UriKind.Absolute), ".").AbsoluteUri;
                Assert.EndsWith("/", baseHref);
                string html = $"<base href=\"{baseHref}\"><link rel=\"stylesheet\" href=\"style.css\" /><p>Test</p>";
                var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions { AllowDocumentStylesheetLinks = true });
                var run = doc.Paragraphs[0].GetRuns().First();
                Assert.Equal("654321", run.ColorHex);
            } finally {
                File.Delete(cssPath);
                Directory.Delete(dir);
            }
        }

        [Fact]
        public void HtmlToWord_FileStylesheet_DisallowedScheme_EmitsDiagnosticAndSkipsStylesheet() {
            var path = Path.GetTempFileName();
            File.WriteAllText(path, "p { color:#246810; }");
            string html = $"<link rel=\"stylesheet\" href=\"{path}\" /><p>Local CSS</p>";
            try {
                var options = new HtmlToWordOptions { AllowDocumentStylesheetLinks = true };
                options.AllowedStylesheetUriSchemes.Remove(Uri.UriSchemeFile);

                HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
                var doc = conversion.Value;

                var run = doc.Paragraphs[0].GetRuns().First();
                Assert.NotEqual("246810", run.ColorHex);
                var diagnostic = Assert.Single(conversion.Report.Diagnostics);
                Assert.Equal("StylesheetResourceRejectedByPolicy", diagnostic.Code);
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void HtmlToWord_FileStylesheet_MaxCssBytes_StopsBeforeParsingStylesheet() {
            var path = Path.GetTempFileName();
            File.WriteAllText(path, "p { color:#135790; }");
            string html = $"<link rel=\"stylesheet\" href=\"{path}\" /><p>Local CSS</p>";
            try {
                var options = new HtmlToWordOptions {
                    AllowDocumentStylesheetLinks = true,
                    MaxCssBytes = 8
                };

                var exception = Assert.Throws<HtmlConversionLimitException>(() => OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(options));

                Assert.Equal("CssSizeLimitExceeded", exception.Code);
                Assert.Equal(path, exception.LimitSource);
                Assert.Equal(8, exception.Limit);
                Assert.True(exception.Actual > exception.Limit);
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void HtmlToWord_FileStylesheet_MaxTotalCssBytes_StopsBeforeParsingSecondStylesheet() {
            var path = Path.GetTempFileName();
            File.WriteAllText(path, "p { color:#abcdef; }");
            string seededCss = "body { color:#111111; }";
            string html = $"<link rel=\"stylesheet\" href=\"{path}\" /><p>Local CSS</p>";
            try {
                var options = new HtmlToWordOptions {
                    AllowDocumentStylesheetLinks = true,
                    MaxTotalCssBytes = Encoding.UTF8.GetByteCount(seededCss) + 8
                };
                options.StylesheetContents.Add(seededCss);

                var exception = Assert.Throws<HtmlConversionLimitException>(() => OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(options));

                Assert.Equal("CssTotalSizeLimitExceeded", exception.Code);
                Assert.Equal(path, exception.LimitSource);
                Assert.Equal(options.MaxTotalCssBytes, exception.Limit);
                Assert.True(exception.Actual > exception.Limit);
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void HtmlToWord_BodyStylesheetLink_AppliesAcrossDocument() {
            var path = Path.GetTempFileName();
            File.WriteAllText(path, "p { color:#456789; }");
            string html = $"<p>Before</p><link rel=\"stylesheet\" href=\"{path}\" /><p>After</p>";
            try {
                var doc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions { AllowDocumentStylesheetLinks = true });

                var beforeRun = doc.Paragraphs[0].GetRuns().First();
                var afterRun = doc.Paragraphs[1].GetRuns().First();
                Assert.Equal("456789", beforeRun.ColorHex);
                Assert.Equal("456789", afterRun.ColorHex);
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void HtmlToWord_BodyStylesheetLink_Disabled_EmitsDiagnosticAndSkipsStylesheet() {
            var path = Path.GetTempFileName();
            File.WriteAllText(path, "p { color:#765432; }");
            string html = $"<p>Before</p><link rel=\"stylesheet\" href=\"{path}\" /><p>After</p>";
            try {
                var options = new HtmlToWordOptions();
                HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
                var doc = conversion.Value;

                var afterRun = doc.Paragraphs[1].GetRuns().First();
                Assert.NotEqual("765432", afterRun.ColorHex);
                var diagnostic = Assert.Single(conversion.Report.Diagnostics);
                Assert.Equal("HtmlStylesheetLinkSkipped", diagnostic.Code);
                Assert.Equal(path, diagnostic.Source, ignoreCase: true);
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void HtmlToWord_StylesheetLinkWithoutHref_EmitsDiagnostic() {
            string html = "<link rel=\"stylesheet\" /><p>Missing href</p>";
            var options = new HtmlToWordOptions { AllowDocumentStylesheetLinks = true };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            Assert.Equal("Missing href", doc.Paragraphs[0].Text);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("HtmlStylesheetLinkMissingHref", diagnostic.Code);
            Assert.Equal("link", diagnostic.Source);
        }

        [Fact]
        public void HtmlToWord_StylesheetLinkUnsupportedScheme_EmitsPolicyDiagnostic() {
            string html = "<link rel=\"stylesheet\" href=\"data:text/css,p%7Bcolor:%23999999%7D\" /><p>Unsupported scheme</p>";
            var options = new HtmlToWordOptions { AllowDocumentStylesheetLinks = true };

            HtmlToWordResult conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentResult(options);
            var doc = conversion.Value;

            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.NotEqual("999999", run.ColorHex);
            var diagnostic = Assert.Single(conversion.Report.Diagnostics);
            Assert.Equal("StylesheetResourceRejectedByPolicy", diagnostic.Code);
            Assert.Equal("data:text/css,p%7Bcolor:%23999999%7D", diagnostic.Source);
            Assert.Contains("data", diagnostic.Detail, StringComparison.OrdinalIgnoreCase);
        }
    }
}
