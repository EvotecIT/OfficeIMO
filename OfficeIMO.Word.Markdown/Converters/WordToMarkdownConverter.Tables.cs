using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Markdown {
    internal partial class WordToMarkdownConverter {
        internal string ConvertTable(WordTable table, WordToMarkdownOptions options) {
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

        internal string GetCellText(WordTableCell cell, WordToMarkdownOptions options) {
            var sb = new StringBuilder();
            bool first = true;
            foreach (var p in cell.Paragraphs) {
                bool hasRuns = false;
                try {
                    hasRuns = p.GetRuns().Any();
                } catch (InvalidOperationException ex) {
                    System.Diagnostics.Debug.WriteLine($"GetRuns() failed for table cell paragraph: {ex.Message}");
                    hasRuns = false;
                }
                // Render only once per underlying OpenXml paragraph:
                // - If there are runs, render the first-run wrapper only
                // - If there are no runs, render this paragraph once
                if (hasRuns && !p.IsFirstRun) continue;

                var text = RenderRuns(p, options);
                // Guard against accidental newlines inside a cell which would break Markdown tables
                if (!string.IsNullOrEmpty(text)) {
                    text = text.Replace("\r\n", " ").Replace('\n', ' ').Replace('\r', ' ');
                    // Escape pipes to retain column boundaries
                    text = text.Replace("|", "\\|");
                    if (!first) sb.Append("<br/>"); // explicit line break between distinct paragraphs
                    sb.Append(text);
                    first = false;
                }
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
