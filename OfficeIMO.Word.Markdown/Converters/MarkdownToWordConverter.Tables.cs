using Markdig.Extensions.Tables;
using Markdig.Syntax;
using OfficeIMO.Word;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using JustificationValues = DocumentFormat.OpenXml.Wordprocessing.JustificationValues;

namespace OfficeIMO.Word.Markdown.Converters {
    internal partial class MarkdownToWordConverter {
        private static void ProcessTable(Table table, WordDocument document, MarkdownToWordOptions options) {
            int rows = table.Count();
            int cols = table.ColumnDefinitions.Count;
            var wordTable = document.AddTable(rows, cols);
            int r = 0;
            foreach (TableRow row in table) {
                var rowAlignments = GetRowAlignments(row);
                int c = 0;
                foreach (TableCell cell in row) {
                    var wordCell = wordTable.Rows[r].Cells[c];
                    JustificationValues? justification = null;
                    if (rowAlignments != null && c < rowAlignments.Length) {
                        justification = ToJustification(rowAlignments[c]);
                    }
                    if (justification == null && c < table.ColumnDefinitions.Count) {
                        justification = ToJustification(table.ColumnDefinitions[c].Alignment);
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
                    c++;
                }
                r++;
            }
        }

        private static TableColumnAlign?[]? GetRowAlignments(TableRow row) {
            object? data = row.GetData("alignment") ?? row.GetData("alignments");
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