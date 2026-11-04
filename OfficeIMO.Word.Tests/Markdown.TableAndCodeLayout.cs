using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public class MarkdownTableAndCodeLayoutTests {
        [Fact]
        public void MarkdownTable_UsesHorizontalRulesAndLeavesSpaceAfterIt() {
            string md = "| Item | Qty |\n|---|---:|\n| Licences | 12 |\n\nAfter the table.\n";
            using var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md).ToWord();

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
            using var doc = OfficeIMO.Markdown.MarkdownDocument.Parse(md).ToWord();

            var code = doc.Paragraphs.First(paragraph => paragraph.Text.Contains("Get-Process"));
            Assert.Equal(WordBorderStyle.Single, code.Borders.LeftStyle);
            Assert.True(code.Borders.LeftSpace > OfficeLength.Zero);
        }
    }
}
