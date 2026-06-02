using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;
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

        private static string RewriteCopiedFormulaReferences(string formula, int rowOffset, string? sheetName) {
            if (rowOffset == 0 || string.IsNullOrEmpty(formula)) {
                return formula;
            }

            return RewriteFormulaReferencesOutsideStrings(formula, segment => ReplaceFormulaReferences(segment, match => {
                if (!CanRewriteFormulaReference(match, sheetName, allowAbsoluteRows: false, allowOtherSheets: true, out int row)) {
                    return match.Value;
                }

                int targetRow = row + rowOffset;
                if (targetRow <= 0 || targetRow > A1.MaxRows) {
                    return match.Value;
                }

                return BuildFormulaReference(match, targetRow);
            }));
        }

        private static string RewriteShiftedFormulaReferences(string formula, int firstAffectedRow, int rowDelta, string? sheetName = null) {
            if (rowDelta == 0 || firstAffectedRow <= 0 || string.IsNullOrEmpty(formula)) {
                return formula;
            }

            return RewriteFormulaReferencesOutsideStrings(formula, segment => ReplaceFormulaReferences(segment, match => {
                if (!CanRewriteFormulaReference(match, sheetName, allowAbsoluteRows: true, allowOtherSheets: false, out int row) || row < firstAffectedRow) {
                    return match.Value;
                }

                int targetRow = row + rowDelta;
                if (targetRow <= 0 || targetRow > A1.MaxRows) {
                    return match.Value;
                }

                return BuildFormulaReference(match, targetRow);
            }));
        }

        private static string RewriteDeletedFormulaReferences(string formula, int firstDeletedRow, int lastDeletedRow, int rowDelta, string? sheetName) {
            if (rowDelta == 0 || firstDeletedRow <= 0 || lastDeletedRow < firstDeletedRow || string.IsNullOrEmpty(formula)) {
                return formula;
            }

            return RewriteFormulaReferencesOutsideStrings(formula, segment => {
                var protectedRanges = new List<string>();
                string rewrittenRanges = ReplaceFormulaRanges(segment, match => {
                    string replacement = RewriteDeletedFormulaRangeReference(match, firstDeletedRow, lastDeletedRow, rowDelta, sheetName);
                    if (string.Equals(replacement, match.Value, StringComparison.Ordinal)) {
                        return match.Value;
                    }

                    string placeholder = "\u0001R" + protectedRanges.Count.ToString(CultureInfo.InvariantCulture) + "\u0002";
                    protectedRanges.Add(replacement);
                    return placeholder;
                });

                string rewritten = ReplaceFormulaReferences(rewrittenRanges, match => {
                    if (!CanRewriteFormulaReference(match, sheetName, allowAbsoluteRows: true, allowOtherSheets: false, out int row)) {
                        return match.Value;
                    }

                    if (row >= firstDeletedRow && row <= lastDeletedRow) {
                        return "#REF!";
                    }

                    if (row <= lastDeletedRow) {
                        return match.Value;
                    }

                    int targetRow = row + rowDelta;
                    if (targetRow <= 0 || targetRow > A1.MaxRows) {
                        return match.Value;
                    }

                    return BuildFormulaReference(match, targetRow);
                });

                for (int i = 0; i < protectedRanges.Count; i++) {
                    rewritten = rewritten.Replace("\u0001R" + i.ToString(CultureInfo.InvariantCulture) + "\u0002", protectedRanges[i]);
                }

                return rewritten;
            });
        }

        private static string RewriteFormulaReferencesOutsideStrings(string formula, Func<string, string> rewriteSegment) {
            var builder = new StringBuilder(formula.Length);
            int index = 0;
            while (index < formula.Length) {
                int quote = formula.IndexOf('"', index);
                if (quote < 0) {
                    builder.Append(rewriteSegment(formula.Substring(index)));
                    break;
                }

                if (quote > index) {
                    builder.Append(rewriteSegment(formula.Substring(index, quote - index)));
                }

                int literalStart = quote;
                index = quote + 1;
                while (index < formula.Length) {
                    if (formula[index] == '"') {
                        if (index + 1 < formula.Length && formula[index + 1] == '"') {
                            index += 2;
                            continue;
                        }

                        index++;
                        break;
                    }

                    index++;
                }

                builder.Append(formula, literalStart, index - literalStart);
            }

            return builder.ToString();
        }

        private static string ReplaceFormulaReferences(string segment, MatchEvaluator evaluator) {
            return Regex.Replace(
                segment,
                @"(?<![A-Za-z0-9_\.])(?:(?<sheet>'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_\.]*)!)?(?<colAbs>\$?)(?<col>[A-Za-z]{1,3})(?<rowAbs>\$?)(?<row>\d{1,7})(?=[:),+\-*/^&=<> \t\r\n]|$)",
                evaluator,
                RegexOptions.CultureInvariant,
                TimeSpan.FromMilliseconds(200));
        }

        private static string ReplaceFormulaRanges(string segment, MatchEvaluator evaluator) {
            return Regex.Replace(
                segment,
                @"(?<![A-Za-z0-9_\.])(?:(?<sheet>'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_\.]*)!)?(?<startColAbs>\$?)(?<startCol>[A-Za-z]{1,3})(?<startRowAbs>\$?)(?<startRow>\d{1,7}):(?<endColAbs>\$?)(?<endCol>[A-Za-z]{1,3})(?<endRowAbs>\$?)(?<endRow>\d{1,7})(?=[:),+\-*/^&=<> \t\r\n]|$)",
                evaluator,
                RegexOptions.CultureInvariant,
                TimeSpan.FromMilliseconds(200));
        }

        private static string RewriteDeletedFormulaRangeReference(Match match, int firstDeletedRow, int lastDeletedRow, int rowDelta, string? sheetName) {
            string sheetQualifier = match.Groups["sheet"].Value;
            if (sheetQualifier.Length > 0 && !IsCurrentSheetQualifier(sheetQualifier, sheetName)) {
                return match.Value;
            }

            if (!int.TryParse(match.Groups["startRow"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int startRow)
                || !int.TryParse(match.Groups["endRow"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int endRow)) {
                return match.Value;
            }

            if (startRow > endRow || endRow < firstDeletedRow) {
                return match.Value;
            }

            if (startRow >= firstDeletedRow && endRow <= lastDeletedRow) {
                return "#REF!";
            }

            int targetStart = startRow;
            int targetEnd = endRow;
            if (startRow > lastDeletedRow) {
                targetStart += rowDelta;
            } else if (startRow >= firstDeletedRow) {
                targetStart = firstDeletedRow;
            }

            if (endRow > lastDeletedRow) {
                targetEnd += rowDelta;
            } else if (endRow >= firstDeletedRow) {
                targetEnd = firstDeletedRow - 1;
            }

            if (targetStart <= 0 || targetEnd <= 0 || targetEnd < targetStart || targetEnd > A1.MaxRows) {
                return "#REF!";
            }

            return sheetQualifier
                + (sheetQualifier.Length > 0 ? "!" : string.Empty)
                + match.Groups["startColAbs"].Value
                + match.Groups["startCol"].Value
                + match.Groups["startRowAbs"].Value
                + targetStart.ToString(CultureInfo.InvariantCulture)
                + ":"
                + match.Groups["endColAbs"].Value
                + match.Groups["endCol"].Value
                + match.Groups["endRowAbs"].Value
                + targetEnd.ToString(CultureInfo.InvariantCulture);
        }

        private static bool CanRewriteFormulaReference(Match match, string? sheetName, bool allowAbsoluteRows, bool allowOtherSheets, out int row) {
            row = 0;
            string sheetQualifier = match.Groups["sheet"].Value;
            if (sheetQualifier.Length > 0 && !allowOtherSheets && !IsCurrentSheetQualifier(sheetQualifier, sheetName)) {
                return false;
            }

            if (!allowAbsoluteRows && match.Groups["rowAbs"].Value == "$") {
                return false;
            }

            return int.TryParse(match.Groups["row"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out row);
        }

        private static string BuildFormulaReference(Match match, int targetRow) {
            string sheetQualifier = match.Groups["sheet"].Value;
            return sheetQualifier
                + (sheetQualifier.Length > 0 ? "!" : string.Empty)
                + match.Groups["colAbs"].Value
                + match.Groups["col"].Value
                + match.Groups["rowAbs"].Value
                + targetRow.ToString(CultureInfo.InvariantCulture);
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
                if (targetFirstRow <= 0 || targetLastRow <= 0 || targetLastRow < targetFirstRow) {
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
