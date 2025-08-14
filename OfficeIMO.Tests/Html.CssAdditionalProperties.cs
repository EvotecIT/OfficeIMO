using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_Css_MarginShorthand() {
            string html = "<p style=\"margin:10pt 20pt\">Test</p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var paragraph = doc.Paragraphs[0];
            Assert.Equal(10d, paragraph.LineSpacingBeforePoints);
            Assert.Equal(10d, paragraph.LineSpacingAfterPoints);
            Assert.Equal(20d, paragraph.IndentationBeforePoints);
            Assert.Equal(20d, paragraph.IndentationAfterPoints);
        }

        [Fact]
        public void HtmlToWord_Css_LineHeight() {
            string html = "<p style=\"line-height:1.5\">Line</p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var paragraph = doc.Paragraphs[0];
            Assert.Equal(360, paragraph.LineSpacing);
            Assert.Equal(LineSpacingRuleValues.Auto, paragraph.LineSpacingRule);
        }

        [Fact]
        public void HtmlToWord_Css_BackgroundColor() {
            string html = "<p style=\"background-color:#ffff00\">Mark</p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var paragraph = doc.Paragraphs[0];
            Assert.Equal(HighlightColorValues.Yellow, paragraph.Highlight);
        }

        [Fact]
        public void HtmlToWord_Css_TextDecoration() {
            string html = "<p><span style=\"text-decoration:underline line-through\">styled</span></p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.Equal(UnderlineValues.Single, run.Underline);
            Assert.True(run.Strike);
        }
    }
}
