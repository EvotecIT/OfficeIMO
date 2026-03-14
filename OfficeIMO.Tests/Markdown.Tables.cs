using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Word;
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

            Assert.Contains(leftParagraph.GetRuns(), r => r.Bold);
            Assert.Contains(centerParagraph.GetRuns(), r => r.Italic);
        }

        [Fact]
        public void WordToMarkdown_TableAlignmentMarkers() {
            using var doc = WordDocument.Create();
            var table = doc.AddTable(2, 3);

            var left = table.Rows[0].Cells[0].Paragraphs[0];
            left.Text = "Left";
            left.ParagraphAlignment = JustificationValues.Left;

            var center = table.Rows[0].Cells[1].Paragraphs[0];
            center.Text = "Center";
            center.ParagraphAlignment = JustificationValues.Center;

            var right = table.Rows[0].Cells[2].Paragraphs[0];
            right.Text = "Right";
            right.ParagraphAlignment = JustificationValues.Right;

            table.Rows[1].Cells[0].Paragraphs[0].Text = "A";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "B";
            table.Rows[1].Cells[2].Paragraphs[0].Text = "C";

            string markdown = doc.ToMarkdown(new WordToMarkdownOptions());

            var lines = markdown.Split('\n');
            Assert.Equal("| Left | Center | Right |", lines[0].TrimEnd('\r'));
            Assert.Equal("| :--- | :---: | ---: |", lines[1].TrimEnd('\r'));
        }

        [Fact]
        public void MarkdownDoc_ToWordDocument_Preserves_MultiParagraph_TableCell_From_HtmlAst() {
            const string html = """
                <table>
                  <tr>
                    <th>Status</th>
                  </tr>
                  <tr>
                    <td>
                      <p>Healthy</p>
                      <p>Observed from AST path</p>
                    </td>
                  </tr>
                </table>
                """;

            var markdown = html.LoadFromHtml();

            using var doc = markdown.ToWordDocument(new MarkdownToWordOptions { FontFamily = "Calibri" });
            var cellParagraphs = doc.Tables[0].Rows[1].Cells[0].Paragraphs
                .Select(p => p.Text)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();

            Assert.Equal(2, cellParagraphs.Count);
            Assert.Equal("Healthy", cellParagraphs[0]);
            Assert.Equal("Observed from AST path", cellParagraphs[1]);
        }

        [Fact]
        public void Html_LoadFromHtmlViaMarkdown_Preserves_MultiParagraph_TableCell_From_AstBridge() {
            const string html = """
                <table>
                  <tr>
                    <th>Status</th>
                  </tr>
                  <tr>
                    <td>
                      <p>Healthy</p>
                      <p>Observed from bridge</p>
                    </td>
                  </tr>
                </table>
                """;

            using var doc = html.LoadFromHtmlViaMarkdown(
                wordOptions: new MarkdownToWordOptions { FontFamily = "Calibri" });
            var cellParagraphs = doc.Tables[0].Rows[1].Cells[0].Paragraphs
                .Select(p => p.Text)
                .Where(text => !string.IsNullOrWhiteSpace(text))
                .ToList();

            Assert.Equal(2, cellParagraphs.Count);
            Assert.Equal("Healthy", cellParagraphs[0]);
            Assert.Equal("Observed from bridge", cellParagraphs[1]);
        }
    }
}

