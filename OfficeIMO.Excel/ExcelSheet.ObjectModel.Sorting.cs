using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Sorts a rectangular range by a 1-based column offset while moving whole row cell nodes.
        /// </summary>
        public void SortRangeByColumn(string a1Range, int columnOffset, bool ascending = true, bool hasHeader = true) {
            if (columnOffset < 1) throw new ArgumentOutOfRangeException(nameof(columnOffset));
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            int targetColumn = c1 + columnOffset - 1;
            if (targetColumn > c2) throw new ArgumentOutOfRangeException(nameof(columnOffset));

            int firstDataRow = hasHeader ? r1 + 1 : r1;
            if (firstDataRow >= r2) {
                return;
            }

            WriteLock(() => {
                var rows = new List<RowSnapshot>();
                for (int row = firstDataRow; row <= r2; row++) {
                    rows.Add(CaptureRow(row, c1, c2, targetColumn));
                }

                rows.Sort((left, right) => {
                    int result = CompareSortValues(left.SortValue, right.SortValue);
                    if (result == 0) {
                        result = left.OriginalRow.CompareTo(right.OriginalRow);
                    }
                    return ascending ? result : -result;
                });

                var rowMap = BuildSortedRowMap(rows, firstDataRow);
                for (int index = 0; index < rows.Count; index++) {
                    WriteRowSnapshot(firstDataRow + index, c1, c2, rows[index], rowMap);
                }

                RemapSortedRangeMetadata(rowMap, firstDataRow, r2, c1, c2);
                WorksheetRoot.Save();
                ClearHeaderCache();
            });
        }


        private static Dictionary<int, int> BuildSortedRowMap(IReadOnlyList<RowSnapshot> rows, int firstRow) {
            var rowMap = new Dictionary<int, int>();
            for (int index = 0; index < rows.Count; index++) {
                int targetRow = firstRow + index;
                if (rows[index].OriginalRow != targetRow) {
                    rowMap[rows[index].OriginalRow] = targetRow;
                }
            }

            return rowMap;
        }

        private void RemapSortedRangeMetadata(IReadOnlyDictionary<int, int> rowMap, int firstRow, int lastRow, int firstColumn, int lastColumn) {
            if (rowMap.Count == 0) {
                return;
            }

            RemapSortedComments(rowMap, firstRow, lastRow, firstColumn, lastColumn);
            RemapSortedHyperlinks(rowMap, firstRow, lastRow, firstColumn, lastColumn);
            RemapSortedDataValidations(rowMap, firstRow, lastRow, firstColumn, lastColumn);
            RemapSortedConditionalFormatting(rowMap, firstRow, lastRow, firstColumn, lastColumn);
        }

        private void RemapSortedComments(IReadOnlyDictionary<int, int> rowMap, int firstRow, int lastRow, int firstColumn, int lastColumn) {
            var commentsPart = WorksheetCommentsPartRoot;
            if (commentsPart?.Comments?.CommentList == null) {
                return;
            }

            var moved = new List<((int Row, int Col) OldCell, (int Row, int Col) NewCell)>();
            foreach (var comment in commentsPart.Comments.CommentList.Elements<Comment>()) {
                if (comment.Reference?.Value is not string reference) {
                    continue;
                }

                var cell = A1.ParseCellRef(reference);
                if (cell.Row < firstRow || cell.Row > lastRow || cell.Col < firstColumn || cell.Col > lastColumn) {
                    continue;
                }

                if (rowMap.TryGetValue(cell.Row, out int targetRow)) {
                    comment.Reference = A1.CellReference(targetRow, cell.Col);
                    moved.Add((cell, (targetRow, cell.Col)));
                }
            }

            if (moved.Count == 0) {
                return;
            }

            commentsPart.Comments.Save();
            foreach (var pair in moved) {
                RemoveCommentVmlShape(pair.OldCell.Row, pair.OldCell.Col);
            }

            foreach (var pair in moved) {
                EnsureCommentVmlShape(pair.NewCell.Row, pair.NewCell.Col);
            }
        }

        private void RemapSortedHyperlinks(IReadOnlyDictionary<int, int> rowMap, int firstRow, int lastRow, int firstColumn, int lastColumn) {
            var hyperlinks = WorksheetRoot.GetFirstChild<Hyperlinks>();
            if (hyperlinks == null) {
                return;
            }

            foreach (var link in hyperlinks.Elements<Hyperlink>().ToList()) {
                if (link.Reference?.Value is not string reference
                    || !TryRemapReferenceForSortedRange(reference, rowMap, firstRow, lastRow, firstColumn, lastColumn, out string remapped)) {
                    continue;
                }

                bool firstReference = true;
                var insertAfter = link;
                foreach (ReferenceListPart remappedReference in SplitReferenceList(remapped)) {
                    if (firstReference) {
                        link.Reference = remappedReference.ToString();
                        firstReference = false;
                        continue;
                    }

                    var clone = (Hyperlink)link.CloneNode(true);
                    clone.Reference = remappedReference.ToString();
                    hyperlinks.InsertAfter(clone, insertAfter);
                    insertAfter = clone;
                }
            }
        }

        private void RemapSortedDataValidations(IReadOnlyDictionary<int, int> rowMap, int firstRow, int lastRow, int firstColumn, int lastColumn) {
            var validations = WorksheetRoot.GetFirstChild<DataValidations>();
            if (validations == null) {
                return;
            }

            foreach (var validation in validations.Elements<DataValidation>()) {
                if (validation.SequenceOfReferences?.InnerText is string references
                    && TryRemapReferenceListForSortedRange(references, rowMap, firstRow, lastRow, firstColumn, lastColumn, out string remapped)) {
                    validation.SequenceOfReferences = new ListValue<StringValue> { InnerText = remapped };
                }
            }
        }

        private void RemapSortedConditionalFormatting(IReadOnlyDictionary<int, int> rowMap, int firstRow, int lastRow, int firstColumn, int lastColumn) {
            foreach (var conditional in WorksheetRoot.Elements<ConditionalFormatting>()) {
                if (conditional.SequenceOfReferences?.InnerText is string references
                    && TryRemapReferenceListForSortedRange(references, rowMap, firstRow, lastRow, firstColumn, lastColumn, out string remapped)) {
                    conditional.SequenceOfReferences = new ListValue<StringValue> { InnerText = remapped };
                }
            }
        }

        private RowSnapshot CaptureRow(int rowIndex, int firstColumn, int lastColumn, int sortColumn) {
            var cells = new List<CellSnapshot>();
            object? sortValue = null;
            var rowElement = WorksheetRoot.GetFirstChild<SheetData>()?
                .Elements<Row>()
                .FirstOrDefault(row => row.RowIndex?.Value == (uint)rowIndex);
            var rowClone = rowElement == null ? null : (Row)rowElement.CloneNode(false);
            for (int column = firstColumn; column <= lastColumn; column++) {
                var cell = TryGetExistingCell(rowIndex, column);
                var clone = cell == null ? null : (Cell)cell.CloneNode(true);
                cells.Add(new CellSnapshot(column - firstColumn, clone));
                if (column == sortColumn) {
                    sortValue = GetCellValueSnapshot(cell).Value;
                }
            }

            return new RowSnapshot(rowIndex, rowClone, cells, sortValue);
        }

        private void WriteRowSnapshot(int targetRow, int firstColumn, int lastColumn, RowSnapshot snapshot, IReadOnlyDictionary<int, int> rowMap, int formulaRowOffset = 0) {
            for (int column = firstColumn; column <= lastColumn; column++) {
                var source = snapshot.Cells[column - firstColumn].Cell;
                var cell = TryGetExistingCell(targetRow, column);
                if (source == null) {
                    cell?.Remove();
                    continue;
                }

                cell ??= GetCell(targetRow, column);
                cell.RemoveAllChildren();
                cell.CellFormula = null;
                cell.CellValue = null;
                cell.DataType = null;
                cell.StyleIndex = null;

                cell.DataType = source.DataType;
                cell.StyleIndex = source.StyleIndex;
                if (source.CellFormula != null) {
                    var formula = (CellFormula)source.CellFormula.CloneNode(true);
                    if (!string.IsNullOrEmpty(formula.Text)) {
                        formula.Text = formulaRowOffset == 0
                            ? RewriteSortedFormulaReferences(formula.Text, rowMap, firstColumn, lastColumn)
                            : RewriteCopiedFormulaReferences(formula.Text, formulaRowOffset, Name);
                    }
                    cell.CellFormula = formula;
                }
                if (source.CellValue != null) cell.CellValue = (CellValue)source.CellValue.CloneNode(true);
                if (source.InlineString != null) cell.InlineString = (InlineString)source.InlineString.CloneNode(true);
                foreach (var child in source.ChildElements.Where(c =>
                    !(c is DocumentFormat.OpenXml.Spreadsheet.CellFormula)
                    && !(c is DocumentFormat.OpenXml.Spreadsheet.CellValue)
                    && !(c is InlineString))) {
                    cell.Append(child.CloneNode(true));
                }
            }

            CopyRowMetadata(targetRow, snapshot.Row);
        }

        private void CopyRowMetadata(int targetRow, Row? source) {
            if (source == null) {
                return;
            }

            var row = TryGetExistingCell(targetRow, 1)?.Parent as Row;
            if (row == null) {
                row = GetCell(targetRow, 1).Parent as Row;
                row?.Elements<Cell>().FirstOrDefault(cell => cell.CellReference?.Value == BuildCellReference(targetRow, 1) && cell.CellValue == null && cell.CellFormula == null && cell.InlineString == null)?.Remove();
            }

            if (row == null) {
                return;
            }

            var attributes = source.GetAttributes()
                .Where(attribute => !(attribute.LocalName == "r" && attribute.NamespaceUri.Length == 0))
                .ToList();
            row.ClearAllAttributes();
            row.RowIndex = (uint)targetRow;
            row.SetAttributes(attributes);
            row.RowIndex = (uint)targetRow;
        }


        private static int CompareSortValues(object? left, object? right) {
            if (left == null && right == null) return 0;
            if (left == null) return 1;
            if (right == null) return -1;
            if (left is double ld && right is double rd) return ld.CompareTo(rd);
            if (left is IComparable comparable && left.GetType() == right.GetType()) return comparable.CompareTo(right);
            return string.Compare(Convert.ToString(left, CultureInfo.InvariantCulture), Convert.ToString(right, CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase);
        }

        private static bool RangesOverlapInclusive((int r1, int c1, int r2, int c2) left, (int r1, int c1, int r2, int c2) right) {
            return left.r1 <= right.r2 && left.r2 >= right.r1 && left.c1 <= right.c2 && left.c2 >= right.c1;
        }

        private static bool TryRemapReferenceListForSortedRange(string referenceList, IReadOnlyDictionary<int, int> rowMap, int firstRow, int lastRow, int firstColumn, int lastColumn, out string remapped) {
            foreach (ReferenceListPart part in SplitReferenceList(referenceList)) {
                if (TryRemapReferenceForSortedRange(part, rowMap, firstRow, lastRow, firstColumn, lastColumn, out _)) {
                    return BuildRemappedReferenceList(referenceList, rowMap, firstRow, lastRow, firstColumn, lastColumn, out remapped);
                }
            }

            remapped = referenceList;
            return false;
        }

        private static bool BuildRemappedReferenceList(string referenceList, IReadOnlyDictionary<int, int> rowMap, int firstRow, int lastRow, int firstColumn, int lastColumn, out string remapped) {
            var builder = new StringBuilder(referenceList.Length);
            bool first = true;
            foreach (ReferenceListPart part in SplitReferenceList(referenceList)) {
                if (!first) {
                    builder.Append(' ');
                }

                if (TryRemapReferenceForSortedRange(part, rowMap, firstRow, lastRow, firstColumn, lastColumn, out string remappedPart)) {
                    builder.Append(remappedPart);
                } else {
                    part.AppendTo(builder);
                }

                first = false;
            }

            remapped = builder.ToString();
            return true;
        }

        private static bool TryRemapReferenceForSortedRange(string reference, IReadOnlyDictionary<int, int> rowMap, int firstRow, int lastRow, int firstColumn, int lastColumn, out string remapped) {
            return TryRemapReferenceForSortedRange(new ReferenceListPart(reference, 0, reference.Length), rowMap, firstRow, lastRow, firstColumn, lastColumn, out remapped);
        }

        private static bool TryRemapReferenceForSortedRange(ReferenceListPart reference, IReadOnlyDictionary<int, int> rowMap, int firstRow, int lastRow, int firstColumn, int lastColumn, out string remapped) {
            var bounds = TryParseReference(reference, out var parsed) ? parsed : default;
            if (bounds == default
                || bounds.r1 < firstRow
                || bounds.r2 > lastRow
                || bounds.c1 < firstColumn
                || bounds.c2 > lastColumn) {
                remapped = string.Empty;
                return false;
            }

            var remappedRows = new List<int>();
            bool changed = false;
            for (int row = bounds.r1; row <= bounds.r2; row++) {
                if (rowMap.TryGetValue(row, out int targetRow)) {
                    remappedRows.Add(targetRow);
                    changed = true;
                } else {
                    remappedRows.Add(row);
                }
            }

            if (!changed) {
                remapped = string.Empty;
                return false;
            }

            remappedRows.Sort();
            var parts = new List<string>();
            int runStart = remappedRows[0];
            int previous = remappedRows[0];
            for (int index = 1; index < remappedRows.Count; index++) {
                if (remappedRows[index] == previous + 1) {
                    previous = remappedRows[index];
                    continue;
                }

                parts.Add(ToReference(runStart, bounds.c1, previous, bounds.c2));
                runStart = previous = remappedRows[index];
            }

            parts.Add(ToReference(runStart, bounds.c1, previous, bounds.c2));
            remapped = string.Join(" ", parts);
            return true;
        }


        private sealed class RowSnapshot {
            internal RowSnapshot(int originalRow, Row? row, List<CellSnapshot> cells, object? sortValue) {
                OriginalRow = originalRow;
                Row = row;
                Cells = cells;
                SortValue = sortValue;
            }

            internal int OriginalRow { get; }
            internal Row? Row { get; }
            internal List<CellSnapshot> Cells { get; }
            internal object? SortValue { get; }
        }

        private sealed class CellSnapshot {
            internal CellSnapshot(int offset, Cell? cell) {
                Offset = offset;
                Cell = cell;
            }

            internal int Offset { get; }
            internal Cell? Cell { get; }
        }
    }
}
