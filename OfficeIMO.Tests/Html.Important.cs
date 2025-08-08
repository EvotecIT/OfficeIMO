using System.Linq;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_StylesheetImportant_OverridesInline() {
            string html = "<style>p { color:#0000ff !important; }</style><p style=\"color:#ff0000\">Test</p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.Equal("0000ff", run.ColorHex);
        }

        [Fact]
        public void HtmlToWord_ImportantBeatsSpecificity() {
            string html = "<style>p { color:#0000ff !important; } div p { color:#ff0000; }</style><div><p>Test</p></div>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.Equal("0000ff", run.ColorHex);
        }
    }
}
