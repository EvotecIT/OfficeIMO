using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_Css_MarginShorthand() {
            string html = "<p style=\"margin:10pt 20pt\">Test</p>";
            var doc = html.ToWordDocument(new HtmlToWordOptions());
            var paragraph = doc.Paragraphs[0];
            Assert.Equal(10d, paragraph.LineSpacingBeforePoints);
            Assert.Equal(10d, paragraph.LineSpacingAfterPoints);
            Assert.Equal(20d, paragraph.IndentationBeforePoints);
            Assert.Equal(20d, paragraph.IndentationAfterPoints);
        }

        [Fact]
        public void HtmlToWord_Css_LineHeight() {
            string html = "<p style=\"line-height:1.5\">Line</p>";
            var doc = html.ToWordDocument(new HtmlToWordOptions());
            var paragraph = doc.Paragraphs[0];
            Assert.Equal(360, paragraph.LineSpacing);
            Assert.Equal(LineSpacingRuleValues.Auto, paragraph.LineSpacingRule);
        }

        [Fact]
        public void HtmlToWord_Css_BackgroundColor() {
            string html = "<p style=\"background-color:#ffff00\">Mark</p>";
            var doc = html.ToWordDocument(new HtmlToWordOptions());
            var paragraph = doc.Paragraphs[0];
            Assert.Equal(HighlightColorValues.Yellow, paragraph.Highlight);
        }

        [Fact]
        public void HtmlToWord_Css_TextDecoration() {
            string html = "<p><span style=\"text-decoration:underline line-through\">styled</span></p>";
            var doc = html.ToWordDocument(new HtmlToWordOptions());
            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.Equal(UnderlineValues.Single, run.Underline);
            Assert.True(run.Strike);
        }

        [Fact]
        public void HtmlToWord_Css_TextDecorationStyle_MapsUnderlineVariants() {
            string html = "<p><span style=\"text-decoration:underline dotted\">dotted</span><span style=\"text-decoration-line:underline;text-decoration-style:wavy\">wavy</span></p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());
            var runs = doc.Paragraphs[0].GetRuns().ToArray();

            Assert.Equal(UnderlineValues.Dotted, runs[0].Underline);
            Assert.Equal(UnderlineValues.Wave, runs[1].Underline);
        }

        [Fact]
        public void HtmlToWord_Css_TextDecorationNone_ClearsInheritedDecoration() {
            string html = "<p style=\"text-decoration-line:underline;text-decoration-style:double\">under <span style=\"text-decoration:none\">plain</span></p>";

            var doc = html.ToWordDocument(new HtmlToWordOptions());
            var runs = doc.Paragraphs[0].GetRuns().ToArray();

            Assert.Equal(UnderlineValues.Double, runs[0].Underline);
            Assert.Null(runs[1].Underline);
        }
    }
}
