using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Html;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_Paragraph_MixedUnits() {
            string html = "<p style=\"margin-left:1.5em;padding-top:10px;text-align:right\"><span style=\"font-size:24px;color:#123456\">Test</span></p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var paragraph = doc.Paragraphs[0];
            Assert.Equal(JustificationValues.Right, paragraph.ParagraphAlignment);
            Assert.Equal(18d, paragraph.IndentationBeforePoints);
            Assert.Equal(7.5d, paragraph.LineSpacingBeforePoints);
            var run = paragraph.GetRuns().First();
            Assert.Equal(18, run.FontSize);
            Assert.Equal("123456", run.ColorHex);
        }

        [Fact]
        public void HtmlToWord_SpanStyles_MultipleDeclarations() {
            string html = "<p><span style=\"font-weight:bold;font-style:italic;text-decoration:underline line-through;font-size:16pt;color:rgb(0,128,0)\">Styled</span></p>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var run = doc.Paragraphs[0].GetRuns().First();
            Assert.True(run.Bold);
            Assert.True(run.Italic);
            Assert.Equal(UnderlineValues.Single, run.Underline);
            Assert.True(run.Strike);
            Assert.Equal(16, run.FontSize);
            Assert.Equal("008000", run.ColorHex);
        }

        [Fact]
        public void HtmlToWord_NestedInheritance() {
            string html = "<div style=\"color:#ff0000;font-size:20px;\">A<span style=\"font-size:10px;\">B</span><span>C</span></div>";
            var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var runs = doc.Paragraphs[0].GetRuns().ToArray();
            Assert.Equal("ff0000", runs[0].ColorHex);
            Assert.Equal(15, runs[0].FontSize);
            Assert.Equal("ff0000", runs[1].ColorHex);
            Assert.Equal(8, runs[1].FontSize);
            Assert.Equal("ff0000", runs[2].ColorHex);
            Assert.Equal(15, runs[2].FontSize);
        }
    }
}
