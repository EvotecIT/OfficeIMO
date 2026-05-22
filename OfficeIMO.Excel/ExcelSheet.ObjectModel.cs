using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Returns a lightweight object wrapper for a single cell.
        /// </summary>
        public ExcelCell CellAt(int row, int column) => new ExcelCell(this, row, column);

        /// <summary>
        /// Returns a lightweight object wrapper for an A1 range.
        /// </summary>
        public ExcelRange Range(string a1Range) => new ExcelRange(this, a1Range);

        /// <summary>
        /// Returns a lightweight object wrapper for a table by name, display name, or range.
        /// </summary>
        public ExcelTable Table(string nameOrRange) => new ExcelTable(this, nameOrRange);

        internal ExcelCellData GetCellValueSnapshot(int row, int column) {
            var cell = TryGetExistingCell(row, column);
            return GetCellValueSnapshot(cell);
        }

        private ExcelCellData GetCellValueSnapshot(Cell? cell) {
            if (cell == null) {
                return new ExcelCellData(ExcelCellDataKind.Blank, null);
            }

            string? cached = cell.CellValue?.Text;
            if (cell.CellFormula != null) {
                object? formulaValue = double.TryParse(cached, NumberStyles.Float, CultureInfo.InvariantCulture, out double cachedNumber)
                    ? cachedNumber
                    : cached;
                return new ExcelCellData(ExcelCellDataKind.Formula, formulaValue, cell.CellFormula.Text, cached);
            }

            if (cell.CellValue == null && cell.InlineString == null) {
                return new ExcelCellData(ExcelCellDataKind.Blank, null);
            }

            var type = cell.DataType?.Value;
            if (type == DocumentFormat.OpenXml.Spreadsheet.CellValues.Boolean) {
                return new ExcelCellData(ExcelCellDataKind.Boolean, cached == "1" || string.Equals(cached, "true", StringComparison.OrdinalIgnoreCase), cachedText: cached);
            }

            if (type == DocumentFormat.OpenXml.Spreadsheet.CellValues.Error) {
                return new ExcelCellData(ExcelCellDataKind.Error, cached, cachedText: cached);
            }

            string text = GetCellText(cell);
            if (type == DocumentFormat.OpenXml.Spreadsheet.CellValues.String
                || type == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString
                || type == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString) {
                return new ExcelCellData(ExcelCellDataKind.Text, text, cachedText: text);
            }

            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out double number)) {
                return new ExcelCellData(ExcelCellDataKind.Number, number, cachedText: text);
            }

            return string.IsNullOrEmpty(text)
                ? new ExcelCellData(ExcelCellDataKind.Blank, null)
                : new ExcelCellData(ExcelCellDataKind.Text, text, cachedText: text);
        }

        internal string GetCellFormattedText(int row, int column, IFormatProvider? provider) {
            var value = GetCellValueSnapshot(row, column);
            if (value.Value is IFormattable formattable) {
                return formattable.ToString(null, provider ?? CultureInfo.CurrentCulture) ?? string.Empty;
            }

            return Convert.ToString(value.Value, provider as CultureInfo ?? CultureInfo.CurrentCulture) ?? value.CachedText ?? string.Empty;
        }

        /// <summary>
        /// Applies a number format to every cell in the range.
        /// </summary>
        public void FormatRange(string a1Range, string numberFormat) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);

            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            WriteLock(() => FormatRangeCore(r1, c1, r2, c2, numberFormat));
        }

        /// <summary>
        /// Applies a solid fill to every cell in the range.
        /// </summary>
        public void FillRange(string a1Range, string hexColor) {
            if (string.IsNullOrWhiteSpace(hexColor)) return;
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);

            if (!_excelDocument.IsMaterializingDeferredDataSetImport) {
                MaterializeDeferredDataSetImportIfNeeded();
            }

            WriteLock(() => FillRangeCore(r1, c1, r2, c2, hexColor));
        }

        /// <summary>
        /// Clears selected parts of every cell and attached worksheet metadata in the range.
        /// </summary>
        public void ClearRange(string a1Range, ExcelClearOptions options = ExcelClearOptions.All) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            if (options == ExcelClearOptions.None) {
                return;
            }

            WriteLock(() => {
                var ws = WorksheetRoot;
                bool worksheetChanged = false;
                bool clearCellFields = options.HasFlag(ExcelClearOptions.Values)
                    || options.HasFlag(ExcelClearOptions.Formulas)
                    || options.HasFlag(ExcelClearOptions.Styles);

                if (clearCellFields) {
                    worksheetChanged |= ClearExistingCellFieldsInRange((r1, c1, r2, c2), options);
                }

                if (options.HasFlag(ExcelClearOptions.Comments)) {
                    worksheetChanged |= ClearCommentsInRange(r1, c1, r2, c2);
                }

                if (options.HasFlag(ExcelClearOptions.Hyperlinks)) {
                    worksheetChanged |= ClearHyperlinksInRange(ws, (r1, c1, r2, c2));
                }

                if (options.HasFlag(ExcelClearOptions.DataValidations)) {
                    RemoveDataValidationsCore(a1Range);
                }

                if (options.HasFlag(ExcelClearOptions.ConditionalFormatting)) {
                    ClearConditionalFormattingCore(a1Range);
                }

                if (options.HasFlag(ExcelClearOptions.Merges)) {
                    UnmergeRangeCore((r1, c1, r2, c2));
                }

                if (options.HasFlag(ExcelClearOptions.Sparklines)) {
                    worksheetChanged |= ClearSparklinesInRange((r1, c1, r2, c2));
                }

                if (worksheetChanged) {
                    ws.Save();
                    ClearHeaderCache();
                }
            });
        }

        private bool ClearExistingCellFieldsInRange((int r1, int c1, int r2, int c2) bounds, ExcelClearOptions options) {
            var sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return false;
            }

            bool clearValues = options.HasFlag(ExcelClearOptions.Values);
            bool clearFormulas = options.HasFlag(ExcelClearOptions.Formulas);
            bool clearStyles = options.HasFlag(ExcelClearOptions.Styles);
            bool changed = false;

            foreach (var row in sheetData.Elements<Row>()) {
                uint rowIndex = row.RowIndex?.Value ?? 0U;
                if (rowIndex < (uint)bounds.r1 || rowIndex > (uint)bounds.r2) {
                    continue;
                }

                foreach (var cell in row.Elements<Cell>()) {
                    if (cell.CellReference?.Value is not string reference) {
                        continue;
                    }

                    int columnIndex = GetColumnIndex(reference);
                    if (columnIndex < bounds.c1 || columnIndex > bounds.c2) {
                        continue;
                    }

                    if (clearValues && (cell.CellValue != null || cell.DataType != null || cell.InlineString != null)) {
                        cell.CellValue = null;
                        cell.DataType = null;
                        cell.InlineString = null;
                        changed = true;
                    }

                    if (clearFormulas && cell.CellFormula != null) {
                        cell.CellFormula = null;
                        changed = true;
                    }

                    if (clearStyles && cell.StyleIndex != null) {
                        cell.StyleIndex = null;
                        changed = true;
                    }
                }
            }

            return changed;
        }

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

        /// <summary>
        /// Merges the specified A1 range.
        /// </summary>
        public void MergeRange(string a1Range) {
            A1.ParseRange(a1Range);
            WriteLock(() => {
                var ws = WorksheetRoot;
                var merges = ws.GetFirstChild<MergeCells>();
                uint mergeCount = 0;

                if (merges == null) {
                    var customSheetViews = ws.GetFirstChild<CustomSheetViews>();
                    merges = new MergeCells();
                    if (customSheetViews != null) {
                        ws.InsertBefore(merges, customSheetViews);
                    } else {
                        ws.Append(merges);
                    }
                } else if (MergeCellsContainReference(merges, a1Range, out mergeCount)) {
                    return;
                }

                merges.Append(new MergeCell { Reference = a1Range });
                merges.Count = mergeCount + 1U;
                ws.Save();
            });
        }

        private static bool MergeCellsContainReference(MergeCells merges, string reference, out uint count) {
            count = 0;
            foreach (var merge in merges.Elements<MergeCell>()) {
                count++;
                if (string.Equals(merge.Reference?.Value, reference, StringComparison.OrdinalIgnoreCase)) {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Removes merge definitions that overlap the supplied A1 range.
        /// </summary>
        public void UnmergeRange(string a1Range) {
            var bounds = A1.ParseRange(a1Range);
            WriteLock(() => UnmergeRangeCore(bounds));
        }

        private void UnmergeRangeCore((int r1, int c1, int r2, int c2) bounds) {
            var merges = WorksheetRoot.GetFirstChild<MergeCells>();
            if (merges == null) return;
            if (!MergeCellsOverlap(merges, bounds)) return;

            bool changed = false;
            uint remainingCount = 0;
            foreach (var merge in merges.Elements<MergeCell>().ToList()) {
                if (merge.Reference?.Value is string reference
                    && TryParseReference(reference, out var mergeBounds)
                    && RangesOverlapInclusive(bounds, mergeBounds)) {
                    merge.Remove();
                    changed = true;
                } else {
                    remainingCount++;
                }
            }

            if (changed) {
                merges.Count = remainingCount;
                WorksheetRoot.Save();
            }
        }

        private static bool MergeCellsOverlap(MergeCells merges, (int r1, int c1, int r2, int c2) bounds) {
            foreach (var merge in merges.Elements<MergeCell>()) {
                if (merge.Reference?.Value is string reference
                    && TryParseReference(reference, out var mergeBounds)
                    && RangesOverlapInclusive(bounds, mergeBounds)) {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Writes rich inline text runs into a cell.
        /// </summary>
        public void SetRichText(int row, int column, IEnumerable<ExcelRichTextRun> runs) {
            if (runs == null) throw new ArgumentNullException(nameof(runs));
            WriteLock(() => {
                var cell = GetCell(row, column);
                var inline = new InlineString();
                foreach (var run in runs) {
                    var text = new Text(run.Text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve };
                    var properties = new RunProperties();
                    if (run.Bold) properties.Append(new Bold());
                    if (run.Italic) properties.Append(new Italic());
                    if (run.Underline) properties.Append(new Underline());
                    if (!string.IsNullOrWhiteSpace(run.FontColor)) properties.Append(new Color { Rgb = NormalizeHexColor(run.FontColor!) });
                    if (!string.IsNullOrWhiteSpace(run.FontName)) properties.Append(new RunFont { Val = run.FontName });
                    if (run.FontSize.HasValue) properties.Append(new FontSize { Val = run.FontSize.Value });

                    var openXmlRun = new Run();
                    if (properties.HasChildren) {
                        openXmlRun.Append(properties);
                    }
                    openXmlRun.Append(text);
                    inline.Append(openXmlRun);
                }

                cell.CellFormula = null;
                cell.CellValue = null;
                cell.DataType = DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString;
                cell.InlineString = inline;
                ClearHeaderCache();
            });
        }

        /// <summary>
        /// Reads rich inline text runs from a cell.
        /// </summary>
        public IReadOnlyList<ExcelRichTextRun> GetRichText(int row, int column) {
            var cell = TryGetExistingCell(row, column);
            if (cell?.InlineString == null) {
                return Array.Empty<ExcelRichTextRun>();
            }

            var runs = new List<ExcelRichTextRun>();
            foreach (var run in cell.InlineString.Elements<Run>()) {
                var properties = run.RunProperties;
                var text = run.Text?.Text ?? string.Empty;
                runs.Add(new ExcelRichTextRun(text) {
                    Bold = properties?.GetFirstChild<Bold>() != null,
                    Italic = properties?.GetFirstChild<Italic>() != null,
                    Underline = properties?.GetFirstChild<Underline>() != null,
                    FontColor = properties?.GetFirstChild<Color>()?.Rgb?.Value,
                    FontName = properties?.GetFirstChild<RunFont>()?.Val?.Value,
                    FontSize = properties?.GetFirstChild<FontSize>()?.Val?.Value
                });
            }

            return runs;
        }

        private bool ClearCommentsInRange(int firstRow, int firstColumn, int lastRow, int lastColumn) {
            bool changed = false;
            var commentsPart = WorksheetCommentsPartRoot;
            if (commentsPart?.Comments?.CommentList != null) {
                bool removedComment = false;
                var commentList = commentsPart.Comments.CommentList;
                if (CommentListOverlapsRange(commentList, firstRow, firstColumn, lastRow, lastColumn)) {
                    foreach (var comment in commentList.Elements<Comment>().ToList()) {
                        if (comment.Reference?.Value is not string reference) {
                            continue;
                        }

                        var (row, col) = A1.ParseCellRef(reference);
                        if (row >= firstRow && row <= lastRow && col >= firstColumn && col <= lastColumn) {
                            comment.Remove();
                            removedComment = true;
                        }
                    }
                }

                if (removedComment) {
                    commentsPart.Comments.Save();
                    changed = true;
                }
            }

            changed |= RemoveCommentVmlShapesInRange(firstRow, firstColumn, lastRow, lastColumn);
            changed |= CleanupCommentArtifacts();
            return changed;
        }

        private static bool CommentListOverlapsRange(CommentList commentList, int firstRow, int firstColumn, int lastRow, int lastColumn) {
            foreach (var comment in commentList.Elements<Comment>()) {
                if (comment.Reference?.Value is not string reference) {
                    continue;
                }

                var (row, col) = A1.ParseCellRef(reference);
                if (row >= firstRow && row <= lastRow && col >= firstColumn && col <= lastColumn) {
                    return true;
                }
            }

            return false;
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

        private void WriteRowSnapshot(int targetRow, int firstColumn, int lastColumn, RowSnapshot snapshot, IReadOnlyDictionary<int, int> rowMap) {
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
                        formula.Text = RewriteSortedFormulaReferences(formula.Text, rowMap, firstColumn, lastColumn);
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

        private void RewriteWorksheetFormulaReferences(int firstAffectedRow, int rowDelta) {
            foreach (var cell in WorksheetRoot.Descendants<Cell>()) {
                if (cell.CellFormula?.Text is string formulaText && formulaText.Length > 0) {
                    cell.CellFormula.Text = RewriteShiftedFormulaReferences(formulaText, firstAffectedRow, rowDelta, Name);
                }
            }
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

        private bool ClearHyperlinksInRange(Worksheet ws, (int r1, int c1, int r2, int c2) bounds) {
            var hyperlinks = ws.GetFirstChild<Hyperlinks>();
            if (hyperlinks == null) return false;
            if (!HyperlinksOverlapRange(hyperlinks, bounds)) return false;

            bool changed = false;
            foreach (var link in hyperlinks.Elements<Hyperlink>().ToList()) {
                if (link.Reference?.Value is string reference) {
                    if (!TryRemoveReferenceOverlap(reference, bounds, out var remaining)) {
                        continue;
                    }

                    if (remaining.Count == 0) {
                        link.Remove();
                        changed = true;
                        continue;
                    }

                    link.Reference = remaining[0];
                    var insertAfter = link;
                    for (int index = 1; index < remaining.Count; index++) {
                        var clone = (Hyperlink)link.CloneNode(true);
                        clone.Reference = remaining[index];
                        hyperlinks.InsertAfter(clone, insertAfter);
                        insertAfter = clone;
                    }

                    changed = true;
                }
            }

            return changed;
        }

        private static bool HyperlinksOverlapRange(Hyperlinks hyperlinks, (int r1, int c1, int r2, int c2) bounds) {
            foreach (var link in hyperlinks.Elements<Hyperlink>()) {
                if (link.Reference?.Value is string reference && ReferenceListOverlaps(reference, bounds)) {
                    return true;
                }
            }

            return false;
        }

        private bool ClearSparklinesInRange((int r1, int c1, int r2, int c2) bounds) {
            if (!SparklinesOverlap(bounds)) return false;

            bool changed = false;
            foreach (var sparkline in WorksheetRoot.Descendants<DocumentFormat.OpenXml.Office2010.Excel.Sparkline>().ToList()) {
                var reference = sparkline.ReferenceSequence?.Text;
                if (!string.IsNullOrWhiteSpace(reference) && TryParseReference(reference!, out var sparklineBounds)) {
                    if (RangesOverlapInclusive(bounds, sparklineBounds)) {
                        sparkline.Remove();
                        changed = true;
                    }
                }
            }

            return changed;
        }

        private bool SparklinesOverlap((int r1, int c1, int r2, int c2) bounds) {
            foreach (var sparkline in WorksheetRoot.Descendants<DocumentFormat.OpenXml.Office2010.Excel.Sparkline>()) {
                var reference = sparkline.ReferenceSequence?.Text;
                if (!string.IsNullOrWhiteSpace(reference)
                    && TryParseReference(reference!, out var sparklineBounds)
                    && RangesOverlapInclusive(bounds, sparklineBounds)) {
                    return true;
                }
            }

            return false;
        }

        private static (int r1, int c1, int r2, int c2) CellAsRange(string cellRef) {
            var parsed = A1.ParseCellRef(cellRef);
            return (parsed.Row, parsed.Col, parsed.Row, parsed.Col);
        }

        private static bool TryParseReference(string reference, out (int r1, int c1, int r2, int c2) bounds) {
            return TryParseReference(new ReferenceListPart(reference, 0, reference.Length), out bounds);
        }

        private static bool TryParseReference(ReferenceListPart reference, out (int r1, int c1, int r2, int c2) bounds) {
            int start = reference.Start;
            int length = reference.Length;
            if (!TrimReferenceBounds(reference.Text, ref start, ref length)) {
                bounds = default;
                return false;
            }

            int end = start + length;
            int separator = -1;
            for (int index = start; index < end; index++) {
                if (reference.Text[index] == ':') {
                    separator = index;
                    break;
                }
            }

            if (separator >= 0) {
                if (!TryParseCellReferencePart(reference.Text, start, separator - start, out int r1, out int c1)
                    || !TryParseCellReferencePart(reference.Text, separator + 1, end - separator - 1, out int r2, out int c2)) {
                    bounds = default;
                    return false;
                }

                if (c1 > c2) (c1, c2) = (c2, c1);
                if (r1 > r2) (r1, r2) = (r2, r1);
                bounds = (r1, c1, r2, c2);
                return true;
            }

            if (!TryParseCellReferencePart(reference.Text, start, length, out int row, out int col)) {
                bounds = default;
                return false;
            }

            bounds = (row, col, row, col);
            return true;
        }

        private static bool TrimReferenceBounds(string text, ref int start, ref int length) {
            if (string.IsNullOrEmpty(text) || length <= 0 || start < 0 || start > text.Length || length > text.Length - start) {
                return false;
            }

            int end = start + length;
            while (start < end && char.IsWhiteSpace(text[start])) {
                start++;
            }

            while (end > start && char.IsWhiteSpace(text[end - 1])) {
                end--;
            }

            length = end - start;
            return length > 0;
        }

        private static bool TryParseCellReferencePart(string text, int start, int length, out int row, out int col) {
            row = 0;
            col = 0;
            if (!TrimReferenceBounds(text, ref start, ref length)) {
                return false;
            }

            int end = start + length;
            int index = start;
            if (index < end && text[index] == '$') {
                index++;
            }

            int letterStart = index;
            for (; index < end; index++) {
                char ch = ToUpperAscii(text[index]);
                if (ch < 'A' || ch > 'Z') {
                    break;
                }

                int value = ch - 'A' + 1;
                if (col > (int.MaxValue - value) / 26) {
                    row = 0;
                    col = 0;
                    return false;
                }

                col = (col * 26) + value;
            }

            if (index == letterStart || index == end) {
                row = 0;
                col = 0;
                return false;
            }

            if (text[index] == '$') {
                index++;
            }

            int digitStart = index;
            for (; index < end; index++) {
                char ch = text[index];
                if (ch < '0' || ch > '9') {
                    row = 0;
                    col = 0;
                    return false;
                }

                int digit = ch - '0';
                if (row > (int.MaxValue - digit) / 10) {
                    row = 0;
                    col = 0;
                    return false;
                }

                row = (row * 10) + digit;
            }

            if (index == digitStart || row <= 0 || col <= 0) {
                row = 0;
                col = 0;
                return false;
            }

            return true;
        }

        private static char ToUpperAscii(char character) {
            return character >= 'a' && character <= 'z' ? (char)(character - 32) : character;
        }

        private static string ToReference(int r1, int c1, int r2, int c2) {
            string start = A1.CellReference(r1, c1);
            string end = A1.CellReference(r2, c2);
            return string.Equals(start, end, StringComparison.OrdinalIgnoreCase) ? start : $"{start}:{end}";
        }

        private Cell? TryGetExistingCell(int row, int column) {
            return TryGetCell(row, column);
        }

        private static string RewriteSortedFormulaReferences(string formula, IReadOnlyDictionary<int, int> rowMap, int firstColumn, int lastColumn) {
            if (rowMap.Count == 0 || string.IsNullOrEmpty(formula)) {
                return formula;
            }

            return Regex.Replace(
                formula,
                @"(?<![A-Za-z0-9_\.!])(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})(?=[:),+\-*/^&=<> \t\r\n]|$)",
                match => {
                    bool rowAbsolute = match.Groups[3].Value == "$";
                    if (rowAbsolute || !int.TryParse(match.Groups[4].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int row)) {
                        return match.Value;
                    }

                    var cell = A1.ParseCellRef(match.Groups[2].Value + row.ToString(CultureInfo.InvariantCulture));
                    if (cell.Col < firstColumn || cell.Col > lastColumn || !rowMap.TryGetValue(row, out int targetRow)) {
                        return match.Value;
                    }

                    return match.Groups[1].Value + match.Groups[2].Value + match.Groups[3].Value + targetRow.ToString(CultureInfo.InvariantCulture);
                },
                RegexOptions.CultureInvariant,
                TimeSpan.FromMilliseconds(200));
        }

        private static string RewriteShiftedFormulaReferences(string formula, int firstAffectedRow, int rowDelta, string? sheetName = null) {
            if (rowDelta == 0 || firstAffectedRow <= 0 || string.IsNullOrEmpty(formula)) {
                return formula;
            }

            return Regex.Replace(
                formula,
                @"(?<![A-Za-z0-9_\.])(?:(?<sheet>'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_\.]*)!)?(?<colAbs>\$?)(?<col>[A-Za-z]{1,3})(?<rowAbs>\$?)(?<row>\d{1,7})(?=[:),+\-*/^&=<> \t\r\n]|$)",
                match => {
                    string sheetQualifier = match.Groups["sheet"].Value;
                    if (sheetQualifier.Length > 0 && !IsCurrentSheetQualifier(sheetQualifier, sheetName)) {
                        return match.Value;
                    }

                    bool rowAbsolute = match.Groups["rowAbs"].Value == "$";
                    if (rowAbsolute || !int.TryParse(match.Groups["row"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int row)) {
                        return match.Value;
                    }

                    if (row < firstAffectedRow) {
                        return match.Value;
                    }

                    int targetRow = row + rowDelta;
                    if (targetRow <= 0 || targetRow > A1.MaxRows) {
                        return match.Value;
                    }

                    return sheetQualifier
                        + (sheetQualifier.Length > 0 ? "!" : string.Empty)
                        + match.Groups["colAbs"].Value
                        + match.Groups["col"].Value
                        + match.Groups["rowAbs"].Value
                        + targetRow.ToString(CultureInfo.InvariantCulture);
                },
                RegexOptions.CultureInvariant,
                TimeSpan.FromMilliseconds(200));
        }

        private static bool IsCurrentSheetQualifier(string qualifier, string? sheetName) {
            if (string.IsNullOrEmpty(sheetName)) {
                return false;
            }

            string value = qualifier;
            if (value.Length >= 2 && value[0] == '\'' && value[value.Length - 1] == '\'') {
                value = value.Substring(1, value.Length - 2).Replace("''", "'");
            }

            return string.Equals(value, sheetName, StringComparison.OrdinalIgnoreCase);
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
