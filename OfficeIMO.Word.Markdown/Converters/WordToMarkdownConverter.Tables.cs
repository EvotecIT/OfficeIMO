using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System.Linq;
using System.Text;

namespace OfficeIMO.Word.Markdown {
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

            sb.Append('|');
            foreach (var cell in rows[0].Cells) {
                sb.Append(' ').Append(GetAlignmentMarker(cell)).Append(' ').Append('|');
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

        private static string GetAlignmentMarker(WordTableCell cell) {
            var alignment = cell.Paragraphs.FirstOrDefault()?.ParagraphAlignment;
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

