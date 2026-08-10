using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
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
            return ExcelFormulaReferenceRewriter.RewriteReferences(formula, match => {
                ExcelReference reference = match.Reference;
                if (reference.IsQualified || reference.Kind is not (ExcelReferenceKind.Cell or ExcelReferenceKind.Range)) {
                    return match.Text;
                }
                int startRow = MapSortedRow(reference.Start, rowMap, firstColumn, lastColumn, out bool startChanged);
                int endRow = MapSortedRow(reference.End, rowMap, firstColumn, lastColumn, out bool endChanged);
                return startChanged || endChanged
                    ? FormatFormulaReference(match, startRow, endRow)
                    : match.Text;
            });
        }

        private string RewriteCopiedFormulaReferences(string formula, int rowOffset, string? sheetName) {
            if (rowOffset == 0 || string.IsNullOrEmpty(formula)) {
                return formula;
            }

            return ExcelFormulaReferenceRewriter.RewriteReferences(formula, match => {
                if (IsFormulaFunctionReferenceToken(formula, match)
                    || match.Reference.Kind is not (ExcelReferenceKind.Cell or ExcelReferenceKind.Range)) return match.Text;
                int startRow = CopyRow(match.Reference.Start, rowOffset, out bool startChanged);
                int endRow = CopyRow(match.Reference.End, rowOffset, out bool endChanged);
                return startChanged || endChanged
                    ? FormatFormulaReference(match, startRow, endRow)
                    : match.Text;
            });
        }

        private string RewriteShiftedFormulaReferences(
            string formula,
            int firstAffectedRow,
            int rowDelta,
            string? sheetName = null,
            bool rewriteUnqualifiedReferences = true) {
            if (rowDelta == 0 || firstAffectedRow <= 0 || string.IsNullOrEmpty(formula)) {
                return formula;
            }

            return ExcelFormulaReferenceRewriter.RewriteReferences(formula, match => {
                if (IsFormulaFunctionReferenceToken(formula, match)
                    || !CanRewriteFormulaReference(match.Reference, sheetName, rewriteUnqualifiedReferences)) return match.Text;
                if (match.Reference.Kind == ExcelReferenceKind.Cell) {
                    if (match.Reference.Start.Row < firstAffectedRow) return match.Text;
                    int targetRow = match.Reference.Start.Row + rowDelta;
                    return targetRow <= 0 || targetRow > A1.MaxRows
                        ? "#REF!"
                        : FormatFormulaReference(match, targetRow, targetRow);
                }
                if (match.Reference.Kind is ExcelReferenceKind.Range or ExcelReferenceKind.WholeRow) {
                    return RewriteShiftedFormulaRangeReference(match, firstAffectedRow, rowDelta);
                }
                return match.Text;
            });
        }

        private string RewriteDeletedFormulaReferences(
            string formula,
            int firstDeletedRow,
            int lastDeletedRow,
            int rowDelta,
            string? sheetName,
            bool rewriteUnqualifiedReferences = true) {
            if (rowDelta == 0 || firstDeletedRow <= 0 || lastDeletedRow < firstDeletedRow || string.IsNullOrEmpty(formula)) {
                return formula;
            }

            return ExcelFormulaReferenceRewriter.RewriteReferences(formula, match => {
                if (IsFormulaFunctionReferenceToken(formula, match)
                    || !CanRewriteFormulaReference(match.Reference, sheetName, rewriteUnqualifiedReferences)) return match.Text;
                if (match.Reference.Kind == ExcelReferenceKind.Cell) {
                    int row = match.Reference.Start.Row;
                    if (row >= firstDeletedRow && row <= lastDeletedRow) return "#REF!";
                    if (row <= lastDeletedRow) return match.Text;
                    int targetRow = row + rowDelta;
                    return targetRow <= 0 || targetRow > A1.MaxRows
                        ? match.Text
                        : FormatFormulaReference(match, targetRow, targetRow);
                }
                if (match.Reference.Kind is ExcelReferenceKind.Range or ExcelReferenceKind.WholeRow) {
                    return RewriteDeletedFormulaRangeReference(match, firstDeletedRow, lastDeletedRow, rowDelta);
                }
                return match.Text;
            });
        }

        /// <summary>
        /// Rewrites non-literal formula segments in one pass while preserving escaped string
        /// quotes and double quotes that occur inside single-quoted sheet qualifiers.
        /// </summary>
        internal static string RewriteFormulaReferencesOutsideStrings(string formula, Func<string, string> rewriteSegment) {
            return ExcelFormulaReferenceRewriter.RewriteOutsideStrings(formula, rewriteSegment);
        }

        private bool IsFormulaFunctionReferenceToken(string formula, ExcelFormulaReferenceCandidate match) {
            if (match.Reference.IsQualified
                || match.Reference.Kind != ExcelReferenceKind.Cell
                || match.Reference.Start.ColumnAbsolute
                || match.Reference.Start.RowAbsolute
                || match.HasSpill) {
                return false;
            }

            int cursor = match.Index + match.Length;
            int whitespaceStart = cursor;
            while (cursor < formula.Length && char.IsWhiteSpace(formula[cursor])) {
                cursor++;
            }

            if (cursor == whitespaceStart || cursor >= formula.Length || formula[cursor] != '(') {
                return false;
            }

            string token = match.Text;
            return ExcelFormulaCapabilities.IsBuiltInFunction(token)
                || _excelDocument.Calculation.TryGetCustomFunction(token, out _);
        }

        private static string RewriteDeletedFormulaRangeReference(
            ExcelFormulaReferenceCandidate match,
            int firstDeletedRow,
            int lastDeletedRow,
            int rowDelta) {
            int startRow = match.Reference.Start.Row;
            int endRow = match.Reference.End.Row;

            bool reversed = startRow > endRow;
            int lowRow = Math.Min(startRow, endRow);
            int highRow = Math.Max(startRow, endRow);
            if (highRow < firstDeletedRow) {
                return match.Text;
            }

            if (lowRow >= firstDeletedRow && highRow <= lastDeletedRow) {
                return "#REF!";
            }

            int targetLow = lowRow;
            int targetHigh = highRow;
            if (lowRow > lastDeletedRow) {
                targetLow += rowDelta;
            } else if (lowRow >= firstDeletedRow) {
                targetLow = firstDeletedRow;
            }

            if (highRow > lastDeletedRow) {
                targetHigh += rowDelta;
            } else if (highRow >= firstDeletedRow) {
                targetHigh = firstDeletedRow - 1;
            }

            if (targetLow <= 0 || targetHigh <= 0 || targetHigh < targetLow || targetHigh > A1.MaxRows) {
                return "#REF!";
            }
            int targetStart = reversed ? targetHigh : targetLow;
            int targetEnd = reversed ? targetLow : targetHigh;

            return FormatFormulaReference(match, targetStart, targetEnd);
        }

        private static string RewriteShiftedFormulaRangeReference(
            ExcelFormulaReferenceCandidate match,
            int firstAffectedRow,
            int rowDelta) {
            int startRow = match.Reference.Start.Row;
            int endRow = match.Reference.End.Row;

            bool reversed = startRow > endRow;
            int lowRow = Math.Min(startRow, endRow);
            int highRow = Math.Max(startRow, endRow);
            if (highRow < firstAffectedRow) {
                return match.Text;
            }
            int targetLow = lowRow < firstAffectedRow ? lowRow : lowRow + rowDelta;
            int targetHigh = highRow + rowDelta;
            if (targetLow <= 0 || targetHigh <= 0 || targetHigh < targetLow || targetHigh > A1.MaxRows) {
                return "#REF!";
            }
            int targetStart = reversed ? targetHigh : targetLow;
            int targetEnd = reversed ? targetLow : targetHigh;

            return FormatFormulaReference(match, targetStart, targetEnd);
        }

        private static bool CanRewriteFormulaReference(
            ExcelReference reference,
            string? sheetName,
            bool rewriteUnqualifiedReferences = true) {
            return reference.IsQualified
                ? IsCurrentSheetQualifier(reference.Qualifier!, sheetName)
                : rewriteUnqualifiedReferences;
        }

        private static int MapSortedRow(
            ExcelReferencePoint point,
            IReadOnlyDictionary<int, int> rowMap,
            int firstColumn,
            int lastColumn,
            out bool changed) {
            changed = !point.RowAbsolute
                && point.Column >= firstColumn
                && point.Column <= lastColumn
                && rowMap.TryGetValue(point.Row, out int targetRow)
                && targetRow != point.Row;
            return changed ? rowMap[point.Row] : point.Row;
        }

        private static int CopyRow(ExcelReferencePoint point, int rowOffset, out bool changed) {
            changed = false;
            if (point.RowAbsolute) return point.Row;
            int target = point.Row + rowOffset;
            if (target <= 0 || target > A1.MaxRows) return point.Row;
            changed = target != point.Row;
            return target;
        }

        private static string FormatFormulaReference(ExcelFormulaReferenceCandidate match, int startRow, int endRow) =>
            FormatFormulaReference(
                match.Reference,
                startRow,
                endRow,
                match.HasSpill,
                match.Text.LastIndexOf('!') > match.Text.LastIndexOf(':'));

        private static string FormatFormulaReference(
            ExcelReference reference,
            int startRow,
            int endRow,
            bool spill,
            bool repeatEndQualifier = false) {
            string qualifier = reference.IsQualified ? reference.Qualifier + "!" : string.Empty;
            string start = FormatFormulaReferencePoint(reference.Start, reference.Kind, startRow);
            string result = qualifier + start;
            if (reference.Kind != ExcelReferenceKind.Cell) {
                result += ":" + (repeatEndQualifier ? qualifier : string.Empty)
                    + FormatFormulaReferencePoint(reference.End, reference.Kind, endRow);
            }
            return spill ? result + "#" : result;
        }

        private static string FormatFormulaReferencePoint(ExcelReferencePoint point, ExcelReferenceKind kind, int row) {
            if (kind == ExcelReferenceKind.WholeRow) {
                return (point.RowAbsolute ? "$" : string.Empty) + row.ToString(CultureInfo.InvariantCulture);
            }
            string column = (point.ColumnAbsolute ? "$" : string.Empty) + A1.ColumnIndexToLetters(point.Column);
            if (kind == ExcelReferenceKind.WholeColumn) return column;
            return column + (point.RowAbsolute ? "$" : string.Empty) + row.ToString(CultureInfo.InvariantCulture);
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

        private static bool TryRemapShiftedReferenceListRows(string referenceList, int firstAffectedRow, int rowDelta, int? lastDeletedRow, out List<string> remapped) {
            remapped = new List<string>();
            bool changed = false;
            foreach (ReferenceListPart part in SplitReferenceList(referenceList)) {
                if (!TryParseReference(part, out var bounds)) {
                    remapped.Add(part.ToString());
                    continue;
                }

                if (!TryRemapShiftedReferenceRows(bounds, firstAffectedRow, rowDelta, lastDeletedRow, out var remappedBounds)) {
                    remapped.Add(part.ToString());
                    continue;
                }

                changed = true;
                if (remappedBounds != null) {
                    remapped.Add(ToReference(remappedBounds.Value.r1, remappedBounds.Value.c1, remappedBounds.Value.r2, remappedBounds.Value.c2));
                }
            }

            return changed;
        }

        private static bool TryRemapShiftedReferenceRows((int r1, int c1, int r2, int c2) bounds, int firstAffectedRow, int rowDelta, int? lastDeletedRow, out (int r1, int c1, int r2, int c2)? remapped) {
            remapped = null;
            if (rowDelta == 0 || firstAffectedRow <= 0 || bounds.r2 < firstAffectedRow) {
                return false;
            }

            if (!lastDeletedRow.HasValue) {
                int targetFirstRow = bounds.r1 < firstAffectedRow ? bounds.r1 : bounds.r1 + rowDelta;
                int targetLastRow = bounds.r2 + rowDelta;
                if (targetFirstRow <= 0
                    || targetLastRow <= 0
                    || targetLastRow < targetFirstRow
                    || targetLastRow > A1.MaxRows) {
                    remapped = null;
                    return true;
                }

                remapped = (targetFirstRow, bounds.c1, targetLastRow, bounds.c2);
                return true;
            }

            int deletedLast = lastDeletedRow.Value;
            if (bounds.r1 >= firstAffectedRow && bounds.r2 <= deletedLast) {
                remapped = null;
                return true;
            }

            int newFirst = bounds.r1 > deletedLast ? bounds.r1 + rowDelta : bounds.r1;
            int newLast = bounds.r2 > deletedLast ? bounds.r2 + rowDelta : firstAffectedRow - 1;
            if (bounds.r1 >= firstAffectedRow && bounds.r1 <= deletedLast) {
                newFirst = firstAffectedRow;
            }

            if (newFirst <= 0 || newLast <= 0 || newLast < newFirst) {
                remapped = null;
                return true;
            }

            remapped = (newFirst, bounds.c1, newLast, bounds.c2);
            return true;
        }
    }
}
