using Markdig.Extensions.Tables;
using Markdig.Syntax;
using OfficeIMO.Word;
using System.Linq;

namespace OfficeIMO.Word.Markdown.Converters {
    internal partial class MarkdownToWordConverter {
        private static void ProcessTable(Table table, WordDocument document, MarkdownToWordOptions options) {
            int rows = table.Count();
            int cols = table.ColumnDefinitions.Count;
            var wordTable = document.AddTable(rows, cols);
            int r = 0;
            foreach (TableRow row in table) {
                int c = 0;
                foreach (TableCell cell in row) {
                    var target = wordTable.Rows[r].Cells[c].Paragraphs[0];
                    foreach (var cellBlock in cell) {
                        if (cellBlock is ParagraphBlock pb) {
                            ProcessInline(pb.Inline, target, options, document);
                        }
                    }
                    c++;
                }
                r++;
            }
        }
    }
}