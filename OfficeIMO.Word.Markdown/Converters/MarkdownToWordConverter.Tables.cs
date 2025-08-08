using Markdig.Extensions.Tables;
using Markdig.Syntax;
using OfficeIMO.Word;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using JustificationValues = DocumentFormat.OpenXml.Wordprocessing.JustificationValues;


namespace OfficeIMO.Word.Markdown.Converters {

    internal partial class MarkdownToWordConverter {

        private static void ProcessTable(Table table, WordDocument document, MarkdownToWordOptions options) {
            int rows = table.Count();
            int cols = 0;
            foreach (TableRow row in table) {
                int count = 0;
                foreach (TableCell cell in row) {
                    int span = cell.ColumnSpan > 0 ? cell.ColumnSpan : 1;
                    count += span;
                }
                cols = Math.Max(cols, count);
            }
            var wordTable = document.AddTable(rows, cols);
            var occupied = new bool[rows, cols];
            int r = 0;
            foreach (TableRow row in table) {
                var rowAlignments = GetRowAlignments(row);
                int cIndex = 0;
                foreach (TableCell cell in row) {
                    while (cIndex < cols && occupied[r, cIndex]) {
                        cIndex++;
                    }
                    var wordCell = wordTable.Rows[r].Cells[cIndex];
                    JustificationValues? justification = null;
                    if (rowAlignments != null && cIndex < rowAlignments.Length) {
                        justification = ToJustification(rowAlignments[cIndex]);
                    }
                    if (justification == null && cIndex < table.ColumnDefinitions.Count) {
                        justification = ToJustification(table.ColumnDefinitions[cIndex].Alignment);
                    }
                    int blockIndex = 0;
                    foreach (var cellBlock in cell) {
                        var paragraph = blockIndex == 0 ? wordCell.Paragraphs[0] : wordCell.AddParagraph();
                        if (justification != null) {
                            paragraph.SetAlignment(justification.Value);
                        }
                        if (cellBlock is ParagraphBlock pb) {
                            ProcessInline(pb.Inline, paragraph, options, document);
                        }
                        blockIndex++;
                    }
                    if (blockIndex == 0) {
                        var paragraph = wordCell.Paragraphs[0];
                        if (justification != null) {
                            paragraph.SetAlignment(justification.Value);
                        }
                    }

                    int rowSpan = cell.RowSpan > 0 ? cell.RowSpan : 1;
                    int colSpan = cell.ColumnSpan > 0 ? cell.ColumnSpan : 1;
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
                r++;
            }
        }


        private static TableColumnAlign?[]? GetRowAlignments(TableRow row) {

            object data = row.GetData("alignment") ?? row.GetData("alignments");

            if (data is IEnumerable enumerable) {

                List<TableColumnAlign?> list = new List<TableColumnAlign?>();

                foreach (var item in enumerable) {

                    if (item is TableColumnAlign align) {

                        list.Add(align);

                    } else if (item is string str) {

                        if (Enum.TryParse<TableColumnAlign>(str, true, out var parsed)) {

                            list.Add(parsed);

                        } else {

                            list.Add(null);

                        }

                    } else {

                        list.Add(null);

                    }

                }

                return list.ToArray();

            }

            return null;

        }



        private static JustificationValues? ToJustification(TableColumnAlign? align) {

            if (align == null) {

                return null;

            }



            switch (align.Value) {

                case TableColumnAlign.Left:

                    return JustificationValues.Left;

                case TableColumnAlign.Center:

                    return JustificationValues.Center;

                case TableColumnAlign.Right:

                    return JustificationValues.Right;

                default:

                    return null;

            }

        }

    }

        }

        private static TableColumnAlign?[]? GetRowAlignments(TableRow row) {
            object data = row.GetData("alignment") ?? row.GetData("alignments");
            if (data is IEnumerable enumerable) {
                List<TableColumnAlign?> list = new List<TableColumnAlign?>();
                foreach (var item in enumerable) {
                    if (item is TableColumnAlign align) {
                        list.Add(align);
                    } else if (item is string str) {
                        if (Enum.TryParse<TableColumnAlign>(str, true, out var parsed)) {
                            list.Add(parsed);
                        } else {
                            list.Add(null);
                        }
                    } else {
                        list.Add(null);
                    }
                }
                return list.ToArray();
            }
            return null;
        }

        private static JustificationValues? ToJustification(TableColumnAlign? align) {
            if (align == null) {
                return null;
            }

            switch (align.Value) {
                case TableColumnAlign.Left:
                    return JustificationValues.Left;
                case TableColumnAlign.Center:
                    return JustificationValues.Center;
                case TableColumnAlign.Right:
                    return JustificationValues.Right;
                default:
                    return null;
            }
        }
    }
}