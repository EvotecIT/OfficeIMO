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
            foreach (var row in tableElem.Rows) {
                int count = 0;
                foreach (var cellElem in row.Cells) {
                    int span = 1;
                    if (cellElem is IHtmlTableCellElement cellElement) {
                        span = Math.Max(1, cellElement.ColumnSpan);
                    }
                    count += span;
                }
                cols = Math.Max(cols, count);
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
            var occupied = new bool[rows, cols];
            for (int r = 0; r < rows; r++) {
                var htmlRow = tableElem.Rows[r];
                int cIndex = 0;
                for (int c = 0; c < htmlRow.Cells.Length; c++) {
                    while (cIndex < cols && occupied[r, cIndex]) {
                        cIndex++;
                    }

                    var htmlCell = htmlRow.Cells[c];
                    var wordCell = wordTable.Rows[r].Cells[cIndex];
                    if (wordCell.Paragraphs.Count == 1 && string.IsNullOrEmpty(wordCell.Paragraphs[0].Text)) {
                        wordCell.Paragraphs[0].Remove();
                    }
                    foreach (var child in htmlCell.ChildNodes) {
                        ProcessNode(child, doc, section, options, null, listStack, new TextFormatting(), wordCell);
                    }

                    int rowSpan = 1;
                    int colSpan = 1;
                    if (htmlCell is IHtmlTableCellElement htmlTableCell) {
                        rowSpan = Math.Max(1, htmlTableCell.RowSpan);
                        colSpan = Math.Max(1, htmlTableCell.ColumnSpan);
                    }

                    if (rowSpan > 1 || colSpan > 1) {
                        wordTable.MergeCells(r, cIndex, rowSpan, colSpan);
                        for (int rr = r; rr < r + rowSpan; rr++) {
                            for (int cc = cIndex; cc < cIndex + colSpan; cc++) {
                                if (rr == r && cc == cIndex) {
                                    continue;
                                }
                                occupied[rr, cc] = true;
                            }
                        }
                    }

                    cIndex++;
                }
            }
        }
    }
}