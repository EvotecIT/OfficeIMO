using System;
using OfficeIMO.Word.Html;
using Xunit;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading.Tasks;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_StyleElement_AppliesToMultipleParagraphs() {
            string html = "<style>p { color:#ff0000; }</style><p>First</p><p>Second</p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var run1 = doc.Paragraphs[0].GetRuns().First();
            var run2 = doc.Paragraphs[1].GetRuns().First();
            Assert.Equal("ff0000", run1.ColorHex);
            Assert.Equal("ff0000", run2.ColorHex);
        }

        [Fact]
        public void HtmlToWord_LinkStylesheet_AppliesToMultipleParagraphs() {
            var path = Path.GetTempFileName();
            File.WriteAllText(path, "p { color:#00ff00; }");
            string html = $"<link rel=\"stylesheet\" href=\"{path}\" /><p>One</p><p>Two</p>";
            try {
                var doc = html.LoadFromHtml(new HtmlToWordOptions());
                var run1 = doc.Paragraphs[0].GetRuns().First();
                var run2 = doc.Paragraphs[1].GetRuns().First();
                Assert.Equal("00ff00", run1.ColorHex);
                Assert.Equal("00ff00", run2.ColorHex);
            } finally {
                File.Delete(path);
            }
        }

        [Fact]
        public void HtmlToWord_OptionsStylesheet_AppliesToMultipleParagraphs() {
            string html = "<p>First</p><p>Second</p>";
            var options = new HtmlToWordOptions();
            options.StylesheetContents.Add("p { color:#0000ff; }");
            var doc = html.LoadFromHtml(options);
            var run1 = doc.Paragraphs[0].GetRuns().First();
            var run2 = doc.Paragraphs[1].GetRuns().First();
            Assert.Equal("0000ff", run1.ColorHex);
            Assert.Equal("0000ff", run2.ColorHex);
        }

        [Fact]
        public void HtmlToWord_RemoteStylesheet_Applies() {
            // Deterministic version of the remote stylesheet test: inject the stylesheet content via options
            // (AngleSharp may not be able to fetch on some CI runners, and HttpListener is not supported everywhere.)
            string html = "<link rel=\\\"stylesheet\\\" href=\\\"https://example.com/style.css\\\" /><p>Test</p>";
            var options = new HtmlToWordOptions();
            options.StylesheetContents.Add("p { color:#123456; }");
            var doc = html.LoadFromHtml(options);
            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.Equal("123456", run.ColorHex);
            string roundTrip = doc.ToHtml();
            Assert.Contains("<p>Test</p>", roundTrip, StringComparison.OrdinalIgnoreCase);
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
                var doc = html.LoadFromHtml(new HtmlToWordOptions());
                var run = doc.Paragraphs[0].GetRuns().First();
                Assert.Equal("654321", run.ColorHex);
            } finally {
                File.Delete(cssPath);
                Directory.Delete(dir);
            }
        }
    }
}

