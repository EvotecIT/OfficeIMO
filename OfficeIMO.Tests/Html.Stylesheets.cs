using OfficeIMO.Word.Html;
using Xunit;

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
    }
}

