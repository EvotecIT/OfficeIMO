using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public class MarkdownTableAndCodeLayoutTests {
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void MarkdownQuoteRuleRequiresATheme(bool themed) {
            using var document = OfficeIMO.Markdown.MarkdownDocument.Parse("> Source quote\n").ToWord(
                new MarkdownWordExportOptions { ApplyDefaultTheme = themed });
            var quote = Assert.Single(document.Paragraphs);
            Assert.True(quote.IndentationBefore > OfficeLength.Zero);
            Assert.Equal(themed ? WordBorderStyle.Single : (WordBorderStyle?)null, quote.Borders.LeftStyle);
        }

        [Fact]
        public void MarkdownTablesWithoutAThemeKeepPlainNativeFormatting() {
            const string markdown = "| Item | Default | Authored |\n| --- | --- | ---: |\n| A | 12 | 12 |\n";
            using var document = OfficeIMO.Markdown.MarkdownDocument.Parse(markdown).ToWord();
            var table = Assert.Single(document.Tables);
            Assert.Equal(WordTableStyle.TableNormal, table.Style);
            foreach (var cell in table.Rows.SelectMany(row => row.Cells)) {
                Assert.Null(cell.ShadingFillColor);
                Assert.True(cell.Borders.BottomStyle is null or WordBorderStyle.None);
                Assert.True(cell.Borders.LeftStyle is null or WordBorderStyle.None);
            }
            Assert.Null(table.Rows[1].Cells[1].Paragraphs[0].ParagraphAlignment);
            Assert.Equal(WordParagraphAlignment.Right, table.Rows[1].Cells[2].Paragraphs[0].ParagraphAlignment);
        }

        [Fact]
        public void MarkdownTable_UsesHorizontalRulesAndLeavesSpaceAfterIt() {
            string md = "| Item | Qty |\n|---|---:|\n| Licences | 12 |\n\nAfter the table.\n";
            using var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md).ToWord(new OfficeIMO.Word.Markdown.MarkdownWordExportOptions { ApplyDefaultTheme = true });

            var table = doc.Tables.Single();
            foreach (var cell in table.Rows.SelectMany(row => row.Cells)) {
                Assert.True(cell.Borders.LeftStyle is null or WordBorderStyle.None);
                Assert.True(cell.Borders.RightStyle is null or WordBorderStyle.None);
                Assert.Equal(WordBorderStyle.Single, cell.Borders.BottomStyle);
            }

            Paragraph after = Assert.IsType<Paragraph>(table._table.NextSibling());
            Assert.True(int.Parse(after.ParagraphProperties!.SpacingBetweenLines!.Before!.Value!) >= 200);
        }

        [Fact]
        public void MarkdownCodeBlock_KeepsItsRuleOffTheText() {
            string md = "```powershell\nGet-Process\n```\n";
            using var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md).ToWord(new OfficeIMO.Word.Markdown.MarkdownWordExportOptions { ApplyDefaultTheme = true });

            var code = doc.Paragraphs.First(paragraph => paragraph.Text.Contains("Get-Process"));
            Assert.Equal(WordBorderStyle.Single, code.Borders.LeftStyle);
            Assert.True(code.Borders.LeftSpace > OfficeLength.Zero);
        }
    }
}
