using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Word.Markdown.Converters {
    internal partial class WordToMarkdownConverter {
        private string ConvertTable(WordTable table, WordToMarkdownOptions options) {
            var sb = new StringBuilder();
            var rows = table.Rows;
            if (rows.Count == 0) return string.Empty;

            sb.Append('|');
            foreach (var cell in rows[0].Cells) {
                sb.Append(' ').Append(GetCellText(cell, options)).Append(" |");
            }
            sb.AppendLine();

            var alignments = GetColumnAlignments(rows);

            sb.Append('|');
            for (int c = 0; c < alignments.Length; c++) {
                sb.Append(' ').Append(GetAlignmentMarker(alignments[c])).Append(' ').Append('|');
            }
            sb.AppendLine();

            for (int r = 1; r < rows.Count; r++) {
                sb.Append('|');
                foreach (var cell in rows[r].Cells) {
                    sb.Append(' ').Append(GetCellText(cell, options)).Append(" |");
                }
                sb.AppendLine();
            }

            return sb.ToString().TrimEnd();
        }

        private string GetCellText(WordTableCell cell, WordToMarkdownOptions options) {
            var sb = new StringBuilder();
            foreach (var p in cell.Paragraphs) {
                if (sb.Length > 0) sb.Append("<br>");
                sb.Append(RenderRuns(p, options));
            }
            return sb.ToString();
        }

        private static JustificationValues?[] GetColumnAlignments(IReadOnlyList<WordTableRow> rows) {
            int columnCount = rows[0].Cells.Count;
            var result = new JustificationValues?[columnCount];

            for (int c = 0; c < columnCount; c++) {
                foreach (var row in rows) {
                    var paragraph = row.Cells[c].Paragraphs.FirstOrDefault();
                    var alignment = paragraph?.ParagraphAlignment;
                    if (alignment != null) {
                        result[c] = alignment;
                        break;
                    }
                }
            }

            return result;
        }

        private static string GetAlignmentMarker(JustificationValues? alignment) {
            if (alignment == JustificationValues.Center) {
                return ":---:";
            }
            if (alignment == JustificationValues.Right || alignment == JustificationValues.End) {
                return "---:";
            }
            if (alignment == JustificationValues.Left || alignment == JustificationValues.Start) {
                return ":---";
            }
            return "---";
        }
    }
}

