using AngleSharp.Dom;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private void ProcessPreformattedElement(IElement element, WordDocument doc, WordSection section, HtmlToWordOptions options, WordParagraph? currentParagraph, WordTableCell? cell, WordHeaderFooter? headerFooter) {
            var textContent = element.TextContent;
            var lines = textContent.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            int start = 0;
            int end = lines.Length;
            while (start < end && string.IsNullOrEmpty(lines[start])) start++;
            while (end > start && string.IsNullOrEmpty(lines[end - 1])) end--;

            var mono = FontResolver.Resolve("monospace");
            bool bookmarkAdded = false;
            if (options.RenderPreAsTable) {
                WordTable preTable;
                if (cell != null) {
                    preTable = cell.AddTable(1, 1);
                } else if (currentParagraph != null) {
                    preTable = currentParagraph.AddTableAfter(1, 1);
                } else if (headerFooter != null) {
                    preTable = headerFooter.AddTable(1, 1);
                } else {
                    var placeholder = section.AddParagraph("");
                    preTable = placeholder.AddTableAfter(1, 1);
                }

                var preCell = preTable.Rows[0].Cells[0];
                for (int i = start; i < end; i++) {
                    var line = lines[i];
                    var paragraph = i == start ? preCell.AddParagraph("", true) : preCell.AddParagraph("");
                    AddPreformattedLine(element, paragraph, line, mono, ref bookmarkAdded, options);
                }
            } else {
                for (int i = start; i < end; i++) {
                    var line = lines[i];
                    var paragraph = cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");
                    AddPreformattedLine(element, paragraph, line, mono, ref bookmarkAdded, options);
                }
            }
        }

        private void AddPreformattedLine(IElement element, WordParagraph paragraph, string line, string? mono, ref bool bookmarkAdded, HtmlToWordOptions options) {
            paragraph.SetStyleId("HTMLPreformatted");
            if (!string.IsNullOrEmpty(mono)) {
                paragraph.SetFontFamily(mono!);
            }

            ApplyBidiIfPresent(element, paragraph);
            if (!bookmarkAdded) {
                AddBookmarkIfPresent(element, paragraph);
                bookmarkAdded = true;
            }

            var fmt = new TextFormatting(false, false, false, null, mono);
            AddTextRun(paragraph, line, fmt, options);
        }

        private void ProcessInlineCodeElement(IElement element, WordDocument doc, WordSection section, HtmlToWordOptions options, WordParagraph? currentParagraph, Stack<WordList> listStack, TextFormatting formatting, WordTableCell? cell, WordHeaderFooter? headerFooter, WordList? headingList) {
            currentParagraph ??= cell != null ? cell.AddParagraph("", true) : headerFooter != null ? headerFooter.AddParagraph("") : section.AddParagraph("");

            var fmt = formatting;
            ApplySpanStyles(element, ref fmt);
            var mono = FontResolver.Resolve("monospace");
            if (!string.IsNullOrEmpty(mono)) {
                fmt.FontFamily = mono;
            }

            int startRuns = currentParagraph.GetRuns().Count();
            foreach (var child in element.ChildNodes) {
                ProcessNode(child, doc, section, options, currentParagraph, listStack, fmt, cell, headerFooter, headingList);
            }

            var runs = currentParagraph.GetRuns().ToList();
            for (int i = startRuns; i < runs.Count; i++) {
                runs[i].SetCharacterStyleId("HtmlCode");
            }
        }
    }
}
