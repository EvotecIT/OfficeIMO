using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_StylesheetCache_LoadsPathContent() {
            var path = Path.GetTempFileName();
            const string css = "p { color:#111111; }";
            File.WriteAllText(path, css);
            try {
                var html = $"<link rel=\"stylesheet\" href=\"{path}\" /><p>Test</p>";
                using var document = OfficeIMO.Html.HtmlConversionDocument
                    .Parse(html).ToWordDocument(new HtmlToWordOptions {
                        AllowDocumentStylesheetLinks = true
                    });
                Assert.Equal("111111",
                    document.Paragraphs[0].GetRuns().First().ColorHex);
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void HtmlToWord_StylesheetCache_DoesNotReuseStaleRules_WhenPathContentChanges() {
            var path = Path.GetTempFileName();
            string html = $"<link rel=\"stylesheet\" href=\"{path}\" /><p>Test</p>";
            try {
                File.WriteAllText(path, "p { color:#111111; }");
                var firstDoc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions { AllowDocumentStylesheetLinks = true });
                Assert.Equal("111111", firstDoc.Paragraphs[0].GetRuns().First().ColorHex);

                File.WriteAllText(path, "p { color:#222222; }");
                var secondDoc = OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocument(new HtmlToWordOptions { AllowDocumentStylesheetLinks = true });

                Assert.Equal("222222", secondDoc.Paragraphs[0].GetRuns().First().ColorHex);
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public async Task HtmlToWord_StylesheetCache_DoesNotReuseStaleRules_WhenRemoteContentChanges() {
            var call = 0;
            using var httpClient = new HttpClient(new FakeHtmlHttpMessageHandler(_ => {
                call++;
                var color = call == 1 ? "333333" : "444444";
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK) {
                    Content = new StringContent($"p {{ color:#{color}; }}", Encoding.UTF8, "text/css")
                });
            }));
            string html = "<link rel=\"stylesheet\" href=\"https://styles.example.test/live.css\" /><p>Test</p>";
            var options = new HtmlToWordOptions {
                AllowDocumentStylesheetLinks = true,
                HttpClient = httpClient
            };

            var firstDoc = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentAsync(options);
            var secondDoc = await OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToWordDocumentAsync(options);

            Assert.Equal("333333", firstDoc.Paragraphs[0].GetRuns().First().ColorHex);
            Assert.Equal("444444", secondDoc.Paragraphs[0].GetRuns().First().ColorHex);
            Assert.Equal(2, call);
        }

        [Fact]
        public void HtmlToWord_StylesheetCache_LoadsInlineContent() {
            const string css = "p { color:#222222; }";
            var html = $"<style>{css}</style><p>Test</p>";
            using var document = OfficeIMO.Html.HtmlConversionDocument
                .Parse(html).ToWordDocument(new HtmlToWordOptions());
            Assert.Equal("222222",
                document.Paragraphs[0].GetRuns().First().ColorHex);
        }
    }
}
