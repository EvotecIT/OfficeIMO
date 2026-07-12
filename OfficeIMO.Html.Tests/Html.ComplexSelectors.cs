using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_ComplexSelector_Specificity() {
            string html = "<style>.highlight { color:#0000ff; } div .highlight { color:#ff0000; }</style><div><p class=\"highlight\">Test</p></div>";
            var doc = html.ToWordDocument(new HtmlToWordOptions());
            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.Equal("FF0000", run.ColorHex);
        }

        [Fact]
        public void HtmlToWord_InvalidCss_Ignored() {
            string html = "<style>p { color:#00ff00 } .invalid {</style><p>Test</p>";
            var doc = html.ToWordDocument(new HtmlToWordOptions());
            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.Equal("00FF00", run.ColorHex);
        }
    }
}
