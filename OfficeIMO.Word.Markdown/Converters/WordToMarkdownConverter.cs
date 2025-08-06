using OfficeIMO.Word;
using System;
using System.IO;
using System.Text;

namespace OfficeIMO.Word.Markdown.Converters {
    /// <summary>
    /// IMPLEMENTATION GUIDELINES:
    /// 1. Read document content using OfficeIMO.Word API:
    ///    - document.Paragraphs for text content
    ///    - paragraph.Style to determine heading levels
    ///    - document.Lists for bullet/numbered lists
    ///    - document.Tables for tables
    /// 2. Convert OfficeIMO.Word elements to Markdown syntax:
    ///    - WordParagraphStyles.Heading1 -> # Heading
    ///    - WordParagraphStyles.Heading2 -> ## Heading
    ///    - Bold text -> **text**
    ///    - Italic text -> *text*
    ///    - Lists -> - item or 1. item
    ///    - Tables -> | col1 | col2 |
    /// 3. Check paragraph.IsListItem to identify list items
    /// 4. Use paragraph.Bold, paragraph.Italic for inline formatting
    /// </summary>
    internal class WordToMarkdownConverter {
        private readonly StringBuilder _output = new StringBuilder();

        public string Convert(WordDocument document, WordToMarkdownOptions options) {
            if (document == null) throw new ArgumentNullException(nameof(document));
            options ??= new WordToMarkdownOptions();

            foreach (var section in DocumentTraversal.EnumerateSections(document)) {
                foreach (var paragraph in section.Paragraphs) {
                    var text = ConvertParagraph(paragraph);
                    if (!string.IsNullOrEmpty(text)) {
                        _output.AppendLine(text);
                    }
                }

                foreach (var table in section.Tables) {
                    var tableText = ConvertTable(table);
                    if (!string.IsNullOrEmpty(tableText)) {
                        _output.AppendLine(tableText);
                    }
                }
            }

            return _output.ToString().TrimEnd();
        }

        private string ConvertParagraph(WordParagraph paragraph) {
            var sb = new StringBuilder();

            int? headingLevel = paragraph.Style.HasValue
                ? HeadingStyleMapper.GetLevelForHeadingStyle(paragraph.Style.Value)
                : (int?)null;
            if (headingLevel.HasValue && headingLevel.Value > 0) {
                sb.Append(new string('#', headingLevel.Value)).Append(' ');
            }

            var listInfo = DocumentTraversal.GetListInfo(paragraph);
            if (listInfo != null) {
                int level = listInfo.Value.Level;
                sb.Append(new string(' ', level * 2));
                sb.Append(listInfo.Value.Ordered ? "1. " : "- ");
            }

            sb.Append(RenderRuns(paragraph));

            return sb.ToString();
        }

        private string RenderRuns(WordParagraph paragraph) {
            var sb = new StringBuilder();
            foreach (var run in FormattingHelper.GetFormattedRuns(paragraph)) {
                if (run.Image != null) {
                    sb.Append(RenderImage(run.Image));
                    continue;
                }

                string text = run.Text;
                if (string.IsNullOrEmpty(text)) {
                    continue;
                }

                if (run.Bold && run.Italic) {
                    text = $"***{text}***";
                } else if (run.Bold) {
                    text = $"**{text}**";
                } else if (run.Italic) {
                    text = $"*{text}*";
                }

                if (run.Strike) {
                    text = $"~~{text}~~";
                }

                if (run.Code) {
                    text = $"`{text}`";
                }

                if (!string.IsNullOrEmpty(run.Hyperlink)) {
                    text = $"[{text}]({run.Hyperlink})";
                }

                sb.Append(text);
            }

            return sb.ToString();
        }

        private string RenderImage(WordImage image) {
            if (image == null) {
                return string.Empty;
            }

            string alt = !string.IsNullOrEmpty(image.Description)
                ? image.Description
                : (string.IsNullOrEmpty(image.FilePath) ? "" : Path.GetFileName(image.FilePath));

            if (!string.IsNullOrEmpty(image.FilePath)) {
                return $"![{alt}]({image.FilePath})";
            }

            byte[] bytes = image.GetBytes();
            string base64 = System.Convert.ToBase64String(bytes);
            string extension = string.IsNullOrEmpty(image.FilePath) ? null : Path.GetExtension(image.FilePath).ToLower();
            string mime = extension switch {
                ".jpg" => "image/jpeg",
                ".jpeg" => "image/jpeg",
                ".gif" => "image/gif",
                ".bmp" => "image/bmp",
                _ => "image/png"
            };

            return $"![{alt}](data:{mime};base64,{base64})";
        }

        private string ConvertTable(WordTable table) {
            var sb = new StringBuilder();
            var rows = table.Rows;
            if (rows.Count == 0) return string.Empty;

            sb.Append('|');
            foreach (var cell in rows[0].Cells) {
                sb.Append(' ').Append(GetCellText(cell)).Append(" |");
            }
            sb.AppendLine();

            sb.Append('|');
            for (int i = 0; i < rows[0].CellsCount; i++) {
                sb.Append("---|");
            }
            sb.AppendLine();

            for (int r = 1; r < rows.Count; r++) {
                sb.Append('|');
                foreach (var cell in rows[r].Cells) {
                    sb.Append(' ').Append(GetCellText(cell)).Append(" |");
                }
                sb.AppendLine();
            }

            return sb.ToString().TrimEnd();
        }

        private string GetCellText(WordTableCell cell) {
            var sb = new StringBuilder();
            foreach (var p in cell.Paragraphs) {
                if (sb.Length > 0) sb.Append("<br>");
                sb.Append(RenderRuns(p));
            }
            return sb.ToString();
        }

    }
}