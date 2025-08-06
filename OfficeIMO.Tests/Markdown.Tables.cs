using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word.Markdown;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void MarkdownToWord_TableCellAlignmentAndFormatting() {
            string md = "| Left | Center | Right |\n| :--- | :---: | ---: |\n| **Bold** | *Italic* | Normal |";
            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions { FontFamily = "Calibri" });
            var table = doc.Tables[0];

            var leftParagraph = table.Rows[1].Cells[0].Paragraphs[0];
            var centerParagraph = table.Rows[1].Cells[1].Paragraphs[0];
            var rightParagraph = table.Rows[1].Cells[2].Paragraphs[0];

            Assert.Equal(JustificationValues.Left, leftParagraph.ParagraphAlignment);
            Assert.Equal(JustificationValues.Center, centerParagraph.ParagraphAlignment);
            Assert.Equal(JustificationValues.Right, rightParagraph.ParagraphAlignment);

            Assert.True(leftParagraph.GetRuns().Any(r => r.Bold));
            Assert.True(centerParagraph.GetRuns().Any(r => r.Italic));
        }
    }
}

