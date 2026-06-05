using System;
using System.Collections;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        private static IDictionary GetCache() {
            var assembly = typeof(HtmlToWordOptions).Assembly;
            var converterType = assembly.GetType("OfficeIMO.Word.Html.HtmlToWordConverter", true)!;
            var field = converterType.GetField("_stylesheetCache", BindingFlags.NonPublic | BindingFlags.Static)!;
            return (IDictionary)field.GetValue(null)!;
        }

        private static string ComputeHash(string css) {
            using var sha = SHA256.Create();
            var bytes = sha.ComputeHash(Encoding.UTF8.GetBytes(css));
            return BitConverter.ToString(bytes).Replace("-", "");
        }

        [Fact]
        public void HtmlToWord_StylesheetCache_Reused_ForPathContent() {
            var path = Path.GetTempFileName();
            const string css = "p { color:#111111; }";
            File.WriteAllText(path, css);
            try {
                var cache = GetCache();
                var key = ComputeHash(css);
                cache.Remove(key);
                var html = $"<link rel=\"stylesheet\" href=\"{path}\" /><p>Test</p>";
                html.LoadFromHtml(new HtmlToWordOptions { AllowDocumentStylesheetLinks = true });
                var first = cache[key];
                html.LoadFromHtml(new HtmlToWordOptions { AllowDocumentStylesheetLinks = true });
                Assert.Same(first, cache[key]);
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
                var firstDoc = html.LoadFromHtml(new HtmlToWordOptions { AllowDocumentStylesheetLinks = true });
                Assert.Equal("111111", firstDoc.Paragraphs[0].GetRuns().First().ColorHex);

                File.WriteAllText(path, "p { color:#222222; }");
                var secondDoc = html.LoadFromHtml(new HtmlToWordOptions { AllowDocumentStylesheetLinks = true });

                Assert.Equal("222222", secondDoc.Paragraphs[0].GetRuns().First().ColorHex);
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void HtmlToWord_StylesheetCache_DoesNotReuseStaleRules_WhenRemoteContentChanges() {
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

            var firstDoc = html.LoadFromHtml(options);
            var secondDoc = html.LoadFromHtml(options);

            Assert.Equal("333333", firstDoc.Paragraphs[0].GetRuns().First().ColorHex);
            Assert.Equal("444444", secondDoc.Paragraphs[0].GetRuns().First().ColorHex);
            Assert.Equal(2, call);
        }

        [Fact]
        public void HtmlToWord_StylesheetCache_Reused_ForContent() {
            const string css = "p { color:#222222; }";
            var key = ComputeHash(css);
            var cache = GetCache();
            cache.Remove(key);
            var html = $"<style>{css}</style><p>Test</p>";
            html.LoadFromHtml(new HtmlToWordOptions());
            var first = cache[key];
            html.LoadFromHtml(new HtmlToWordOptions());
            Assert.Same(first, cache[key]);
        }
    }
}
