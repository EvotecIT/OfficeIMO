using OfficeIMO.Word.Html;
using OfficeIMO.Word;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_SpanStyles() {
            string html = "<p><span style=\"color:#ff0000;font-family:Arial;font-size:24px\">Styled</span></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var run = doc.Paragraphs[0].GetRuns().First();

            Assert.Equal("ff0000", run.ColorHex);
            Assert.Equal("Arial", run.FontFamily);
            Assert.Equal(24, run.FontSize);
        }

        [Fact]
        public void HtmlToWord_SpanStyles_Decorations() {
            string html = "<p><span style=\"text-decoration:line-through\">strike</span><span style=\"text-decoration:underline\">under</span><span style=\"background-color:#ffff00\">mark</span></p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var runs = doc.Paragraphs;

            var strikeRun = runs.First(r => r.Text == "strike");
            Assert.True(strikeRun.Strike);

            var underRun = runs.First(r => r.Text == "under");
            Assert.Equal(UnderlineValues.Single, underRun.Underline);

            var markRun = runs.First(r => r.Text == "mark");
            Assert.Equal(HighlightColorValues.Yellow, markRun.Highlight);
        }
    }
}
