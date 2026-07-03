using System.Linq;

using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Html {
        [Fact]
        public void HtmlToWord_ParagraphDirection_DirAttribute() {
            string html = "<div dir=\"rtl\"><p>Alpha</p></div>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var paragraph = doc.Paragraphs.First(p => p.Text.Contains("Alpha"));
            Assert.True(paragraph.BiDi);
        }

        [Fact]
        public void HtmlToWord_ParagraphDirection_CssDirection() {
            string html = "<p style=\"direction: rtl\">Bravo</p><p style=\"direction:ltr\">Charlie</p>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var rtlParagraph = doc.Paragraphs.First(p => p.Text.Contains("Bravo"));
            var ltrParagraph = doc.Paragraphs.First(p => p.Text.Contains("Charlie"));

            Assert.True(rtlParagraph.BiDi);
            Assert.False(ltrParagraph.BiDi);
        }

        [Fact]
        public void HtmlToWord_BlockDirection_TextNode() {
            string html = "<div dir=\"rtl\">Delta</div>";

            var doc = html.LoadFromHtml(new HtmlToWordOptions());

            var paragraph = doc.Paragraphs.First(p => p.Text.Contains("Delta"));
            Assert.True(paragraph.BiDi);
        }

        [Fact]
        public void HtmlToWord_LogicalTextAlign_UsesDirection() {
            string html = "<p style=\"text-align:start\">Left</p><p dir=\"rtl\" style=\"text-align:start\">Right</p><p dir=\"rtl\" style=\"text-align:end\">LeftAgain</p>";

            using var doc = html.LoadFromHtml(new HtmlToWordOptions());

            Assert.Equal(JustificationValues.Left, doc.Paragraphs.First(p => p.Text.Contains("Left")).ParagraphAlignment);
            Assert.Equal(JustificationValues.Right, doc.Paragraphs.First(p => p.Text.Contains("Right")).ParagraphAlignment);
            Assert.Equal(JustificationValues.Left, doc.Paragraphs.First(p => p.Text.Contains("LeftAgain")).ParagraphAlignment);
        }

        [Fact]
        public void HtmlToWord_TableCellLogicalTextAlign_UsesInheritedDirection() {
            string html = "<table dir=\"rtl\"><tr><td style=\"text-align:start\">Start</td><td style=\"text-align:end\">End</td></tr></table>";

            using var doc = html.LoadFromHtml(new HtmlToWordOptions());
            var row = doc.Tables[0].Rows[0];

            Assert.Equal(JustificationValues.Right, row.Cells[0].Paragraphs[0].ParagraphAlignment);
            Assert.Equal(JustificationValues.Left, row.Cells[1].Paragraphs[0].ParagraphAlignment);
        }
    }
}
