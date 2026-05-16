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
            for (int row = r1; row <= r2; row++) {
                for (int column = c1; column <= c2; column++) {
                    FormatCell(row, column, numberFormat);
                }
            }
        }

        /// <summary>
        /// Applies a solid fill to every cell in the range.
        /// </summary>
        public void FillRange(string a1Range, string hexColor) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            for (int row = r1; row <= r2; row++) {
                for (int column = c1; column <= c2; column++) {
                    CellBackground(row, column, hexColor);
                }
            }
        }

        /// <summary>
        /// Clears selected parts of every cell and attached worksheet metadata in the range.
        /// </summary>
        public void ClearRange(string a1Range, ExcelClearOptions options = ExcelClearOptions.All) {
            var (r1, c1, r2, c2) = A1.ParseRange(a1Range);
            WriteLock(() => {
                var ws = WorksheetRoot;
                bool clearCellFields = options.HasFlag(ExcelClearOptions.Values)
                    || options.HasFlag(ExcelClearOptions.Formulas)
                    || options.HasFlag(ExcelClearOptions.Styles);

                if (clearCellFields) {
                    for (int row = r1; row <= r2; row++) {
                        for (int column = c1; column <= c2; column++) {
                            var cell = GetCell(row, column);
                            if (options.HasFlag(ExcelClearOptions.Values)) {
                                cell.CellValue = null;
                                cell.DataType = null;
                                cell.InlineString = null;
                            }

                            if (options.HasFlag(ExcelClearOptions.Formulas)) {
                                cell.CellFormula = null;
                            }

                            if (options.HasFlag(ExcelClearOptions.Styles)) {
                                cell.StyleIndex = null;
                            }
                        }
                    }
                }

                if (options.HasFlag(ExcelClearOptions.Comments)) {
                    ClearCommentsInRange(r1, c1, r2, c2);
                }

                if (options.HasFlag(ExcelClearOptions.Hyperlinks)) {
                    ClearHyperlinksInRange(ws, a1Range);
                }

                if (options.HasFlag(ExcelClearOptions.DataValidations)) {
                    RemoveDataValidationsCore(a1Range);
                }

                if (options.HasFlag(ExcelClearOptions.ConditionalFormatting)) {
                    ClearConditionalFormattingCore(a1Range);
                }

                if (options.HasFlag(ExcelClearOptions.Merges)) {
                    UnmergeRangeCore(a1Range);
                }

                if (options.HasFlag(ExcelClearOptions.Sparklines)) {
                    ClearSparklinesInRange(a1Range);
                }

                ws.Save();
                ClearHeaderCache();
            });
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
                if (merges == null) {
                    var customSheetViews = ws.GetFirstChild<CustomSheetViews>();
                    merges = new MergeCells();
                    if (customSheetViews != null) {
                        ws.InsertBefore(merges, customSheetViews);
                    } else {
                        ws.Append(merges);
                    }
                }

                if (!merges.Elements<MergeCell>().Any(m => string.Equals(m.Reference?.Value, a1Range, StringComparison.OrdinalIgnoreCase))) {
                    merges.Append(new MergeCell { Reference = a1Range });
                    merges.Count = (uint)merges.Elements<MergeCell>().Count();
                }

                ws.Save();
            });
        }

        /// <summary>
        /// Removes merge definitions that overlap the supplied A1 range.
        /// </summary>
        public void UnmergeRange(string a1Range) {
            WriteLock(() => UnmergeRangeCore(a1Range));
        }

        private void UnmergeRangeCore(string a1Range) {
            var bounds = A1.ParseRange(a1Range);
            var merges = WorksheetRoot.GetFirstChild<MergeCells>();
            if (merges == null) return;

            foreach (var merge in merges.Elements<MergeCell>().ToList()) {
                if (merge.Reference?.Value is string reference && RangesOverlapInclusive(bounds, A1.ParseRange(reference))) {
                    merge.Remove();
                }
            }

            merges.Count = (uint)merges.Elements<MergeCell>().Count();
            WorksheetRoot.Save();
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

        private void ClearCommentsInRange(int firstRow, int firstColumn, int lastRow, int lastColumn) {
            var commentsPart = WorksheetCommentsPartRoot;
            if (commentsPart?.Comments?.CommentList != null) {
                foreach (var comment in commentsPart.Comments.CommentList.Elements<Comment>().ToList()) {
                    if (comment.Reference?.Value is not string reference) {
                        continue;
                    }

                    var (row, col) = A1.ParseCellRef(reference);
                    if (row >= firstRow && row <= lastRow && col >= firstColumn && col <= lastColumn) {
                        comment.Remove();
                    }
                }

                commentsPart.Comments.Save();
            }

            for (int row = firstRow; row <= lastRow; row++) {
                for (int column = firstColumn; column <= lastColumn; column++) {
                    RemoveCommentVmlShape(row, column);
                }
            }

            CleanupCommentArtifacts();
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

                var references = SplitReferenceList(remapped);
                if (references.Length == 0) {
                    continue;
                }

                link.Reference = references[0];
                var insertAfter = link;
                for (int index = 1; index < references.Length; index++) {
                    var clone = (Hyperlink)link.CloneNode(true);
                    clone.Reference = references[index];
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
            for (int column = firstColumn; column <= lastColumn; column++) {
                var cell = TryGetExistingCell(rowIndex, column);
                var clone = cell == null ? null : (Cell)cell.CloneNode(true);
                cells.Add(new CellSnapshot(column - firstColumn, clone));
                if (column == sortColumn) {
                    sortValue = GetCellValueSnapshot(cell).Value;
                }
            }

            return new RowSnapshot(rowIndex, cells, sortValue);
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
            bool changed = false;
            var parts = new List<string>();
            foreach (string part in SplitReferenceList(referenceList)) {
                if (TryRemapReferenceForSortedRange(part, rowMap, firstRow, lastRow, firstColumn, lastColumn, out string remappedPart)) {
                    parts.Add(remappedPart);
                    changed = true;
                } else {
                    parts.Add(part);
                }
            }

            remapped = string.Join(" ", parts);
            return changed;
        }

        private static bool TryRemapReferenceForSortedRange(string reference, IReadOnlyDictionary<int, int> rowMap, int firstRow, int lastRow, int firstColumn, int lastColumn, out string remapped) {
            remapped = reference;
            var bounds = TryParseReference(reference, out var parsed) ? parsed : default;
            if (bounds == default
                || bounds.r1 < firstRow
                || bounds.r2 > lastRow
                || bounds.c1 < firstColumn
                || bounds.c2 > lastColumn) {
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

        private void ClearHyperlinksInRange(Worksheet ws, string a1Range) {
            var bounds = A1.ParseRange(a1Range);
            var hyperlinks = ws.GetFirstChild<Hyperlinks>();
            if (hyperlinks == null) return;

            foreach (var link in hyperlinks.Elements<Hyperlink>().ToList()) {
                if (link.Reference?.Value is string reference) {
                    var remaining = RemoveReferenceOverlap(reference, bounds);
                    if (remaining.Count == 0) {
                        link.Remove();
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
                }
            }
        }

        private void ClearSparklinesInRange(string a1Range) {
            var bounds = A1.ParseRange(a1Range);
            foreach (var sparkline in WorksheetRoot.Descendants<DocumentFormat.OpenXml.Office2010.Excel.Sparkline>().ToList()) {
                var reference = sparkline.ReferenceSequence?.Text;
                if (!string.IsNullOrWhiteSpace(reference)) {
                    var sparklineBounds = CellAsRange(reference!);
                    if (RangesOverlapInclusive(bounds, sparklineBounds)) {
                        sparkline.Remove();
                    }
                }
            }
        }

        private static (int r1, int c1, int r2, int c2) CellAsRange(string cellRef) {
            var parsed = A1.ParseCellRef(cellRef);
            return (parsed.Row, parsed.Col, parsed.Row, parsed.Col);
        }

        private static bool TryParseReference(string reference, out (int r1, int c1, int r2, int c2) bounds) {
            string normalized = reference.Replace("$", string.Empty);
            if (normalized.IndexOf(':') >= 0) {
                return A1.TryParseRange(normalized, out bounds.r1, out bounds.c1, out bounds.r2, out bounds.c2);
            }

            var cell = A1.ParseCellRef(normalized);
            if (cell.Row <= 0 || cell.Col <= 0) {
                bounds = default;
                return false;
            }

            bounds = (cell.Row, cell.Col, cell.Row, cell.Col);
            return true;
        }

        private static string ToReference(int r1, int c1, int r2, int c2) {
            string start = A1.CellReference(r1, c1);
            string end = A1.CellReference(r2, c2);
            return string.Equals(start, end, StringComparison.OrdinalIgnoreCase) ? start : $"{start}:{end}";
        }

        private Cell? TryGetExistingCell(int row, int column) {
            if (row <= 0) throw new ArgumentOutOfRangeException(nameof(row));
            if (column <= 0) throw new ArgumentOutOfRangeException(nameof(column));

            var sheetData = WorksheetRoot.GetFirstChild<SheetData>();
            if (sheetData == null) {
                return null;
            }

            var rowElement = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == (uint)row);
            if (rowElement == null) {
                return null;
            }

            foreach (Cell cell in rowElement.Elements<Cell>()) {
                if (cell.CellReference?.Value is string reference
                    && GetColumnIndex(reference) == column) {
                    return cell;
                }
            }

            return null;
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

        private sealed class RowSnapshot {
            internal RowSnapshot(int originalRow, List<CellSnapshot> cells, object? sortValue) {
                OriginalRow = originalRow;
                Cells = cells;
                SortValue = sortValue;
            }

            internal int OriginalRow { get; }
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
