using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Markdown {
        [Fact]
        public void MarkdownToWord_TableRowAndColumnSpans() {
            string md = @"+---+---+---+
| AAAAA | B |
+ AAAAA +---+
| AAAAA | C |
+---+---+---+
| D | E | F |
+---+---+---+
";
            var doc = md.LoadFromMarkdown(new MarkdownToWordOptions());
            var table = doc.Tables[0];

            Assert.Equal(MergedCellValues.Restart, table.Rows[0].Cells[0].HorizontalMerge);
            Assert.Equal(MergedCellValues.Continue, table.Rows[0].Cells[1].HorizontalMerge);
            Assert.Equal(MergedCellValues.Restart, table.Rows[0].Cells[0].VerticalMerge);
            Assert.Equal(MergedCellValues.Continue, table.Rows[1].Cells[0].VerticalMerge);
        }
    }
}