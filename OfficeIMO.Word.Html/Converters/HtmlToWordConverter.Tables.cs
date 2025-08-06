using AngleSharp.Html.Dom;
using OfficeIMO.Word;
using System;
using System.Collections.Generic;

namespace OfficeIMO.Word.Html.Converters {
    internal partial class HtmlToWordConverter {
        private void ProcessTable(IHtmlTableElement tableElem, WordDocument doc, WordSection section, HtmlToWordOptions options,
            Stack<WordList> listStack, WordTableCell? cell, WordParagraph? currentParagraph) {
            int rows = tableElem.Rows.Length;
            int cols = 0;
            foreach (var r in tableElem.Rows) {
                cols = Math.Max(cols, r.Cells.Length);
            }
            WordTable wordTable;
            if (cell != null) {
                wordTable = cell.AddTable(rows, cols);
            } else if (currentParagraph != null) {
                wordTable = currentParagraph.AddTableAfter(rows, cols);
            } else {
                var placeholder = section.AddParagraph("");
                wordTable = placeholder.AddTableAfter(rows, cols);
            }
            for (int r = 0; r < rows; r++) {
                var htmlRow = tableElem.Rows[r];
                for (int c = 0; c < htmlRow.Cells.Length; c++) {
                    var htmlCell = htmlRow.Cells[c];
                    var wordCell = wordTable.Rows[r].Cells[c];
                    if (wordCell.Paragraphs.Count == 1 && string.IsNullOrEmpty(wordCell.Paragraphs[0].Text)) {
                        wordCell.Paragraphs[0].Remove();
                    }
                    foreach (var child in htmlCell.ChildNodes) {
                        ProcessNode(child, doc, section, options, null, listStack, new TextFormatting(), wordCell);
                    }
                }
            }
        }
    }
}