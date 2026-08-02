using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Excel {
    /// <summary>Reference notation used by an Excel formula or range address.</summary>
    public enum ExcelReferenceStyle {
        /// <summary>Column letters followed by row numbers, for example <c>$B$4</c>.</summary>
        A1,

        /// <summary>Row/column notation, for example <c>R4C2</c> or <c>R[-1]C</c>.</summary>
        R1C1
    }

    /// <summary>Shape represented by an <see cref="ExcelReference"/>.</summary>
    public enum ExcelReferenceKind {
        /// <summary>One worksheet cell.</summary>
        Cell,

        /// <summary>A rectangular worksheet range.</summary>
        Range,

        /// <summary>One or more complete worksheet rows.</summary>
        WholeRow,

        /// <summary>One or more complete worksheet columns.</summary>
        WholeColumn
    }

    /// <summary>One resolved endpoint in an Excel reference.</summary>
    public readonly struct ExcelReferencePoint : IEquatable<ExcelReferencePoint> {
        internal ExcelReferencePoint(int row, int column, bool rowAbsolute, bool columnAbsolute) {
            Row = row;
            Column = column;
            RowAbsolute = rowAbsolute;
            ColumnAbsolute = columnAbsolute;
        }

        /// <summary>Resolved 1-based row, or zero for a whole-column endpoint.</summary>
        public int Row { get; }

        /// <summary>Resolved 1-based column, or zero for a whole-row endpoint.</summary>
        public int Column { get; }

        /// <summary>Whether the row is absolute in the authored notation.</summary>
        public bool RowAbsolute { get; }

        /// <summary>Whether the column is absolute in the authored notation.</summary>
        public bool ColumnAbsolute { get; }

        /// <inheritdoc />
        public bool Equals(ExcelReferencePoint other) =>
            Row == other.Row
            && Column == other.Column
            && RowAbsolute == other.RowAbsolute
            && ColumnAbsolute == other.ColumnAbsolute;

        /// <inheritdoc />
        public override bool Equals(object? obj) => obj is ExcelReferencePoint other && Equals(other);

        /// <inheritdoc />
        public override int GetHashCode() {
            unchecked {
                int hash = Row;
                hash = (hash * 397) ^ Column;
                hash = (hash * 397) ^ RowAbsolute.GetHashCode();
                return (hash * 397) ^ ColumnAbsolute.GetHashCode();
            }
        }
    }

    /// <summary>
    /// Parsed, format-neutral Excel cell or range reference. Relative R1C1 endpoints are resolved
    /// against the anchor supplied to <see cref="Parse(string,ExcelReferenceStyle,int,int)"/>.
    /// </summary>
    public sealed class ExcelReference : IEquatable<ExcelReference> {
        private ExcelReference(
            ExcelReferenceKind kind,
            string? qualifier,
            ExcelReferencePoint start,
            ExcelReferencePoint end) {
            Kind = kind;
            Qualifier = qualifier;
            Start = start;
            End = end;
        }

        /// <summary>Reference shape.</summary>
        public ExcelReferenceKind Kind { get; }

        /// <summary>Optional workbook/sheet qualifier without the trailing exclamation mark.</summary>
        public string? Qualifier { get; }

        /// <summary>First authored endpoint.</summary>
        public ExcelReferencePoint Start { get; }

        /// <summary>Second authored endpoint. Equals <see cref="Start"/> for a cell.</summary>
        public ExcelReferencePoint End { get; }

        /// <summary>Whether the reference has a workbook or worksheet qualifier.</summary>
        public bool IsQualified => !string.IsNullOrWhiteSpace(Qualifier);

        /// <summary>Parses a reference and throws when it is invalid.</summary>
        public static ExcelReference Parse(
            string text,
            ExcelReferenceStyle style = ExcelReferenceStyle.A1,
            int anchorRow = 1,
            int anchorColumn = 1) {
            if (!TryParse(text, out ExcelReference? reference, style, anchorRow, anchorColumn)) {
                throw new FormatException($"'{text}' is not a valid {style} Excel reference.");
            }

            return reference!;
        }

        /// <summary>Tries to parse an A1 or R1C1 cell/range reference.</summary>
        public static bool TryParse(
            string? text,
            out ExcelReference? reference,
            ExcelReferenceStyle style = ExcelReferenceStyle.A1,
            int anchorRow = 1,
            int anchorColumn = 1) {
            reference = null;
            if (string.IsNullOrWhiteSpace(text)
                || anchorRow < 1 || anchorRow > A1.MaxRows
                || anchorColumn < 1 || anchorColumn > A1.MaxColumns) {
                return false;
            }

            string value = text!.Trim();
            if (!TrySplitQualifier(value, out string? qualifier, out string address)) return false;
            if (address.Length == 0) return false;

            return style == ExcelReferenceStyle.A1
                ? TryParseA1(address, qualifier, out reference)
                : TryParseR1C1(address, qualifier, anchorRow, anchorColumn, out reference);
        }

        /// <summary>Formats the reference using A1 or R1C1 notation.</summary>
        public string ToString(
            ExcelReferenceStyle style,
            int anchorRow = 1,
            int anchorColumn = 1) {
            if (anchorRow < 1 || anchorRow > A1.MaxRows) throw new ArgumentOutOfRangeException(nameof(anchorRow));
            if (anchorColumn < 1 || anchorColumn > A1.MaxColumns) throw new ArgumentOutOfRangeException(nameof(anchorColumn));

            string start = style == ExcelReferenceStyle.A1
                ? FormatA1Point(Start, Kind)
                : FormatR1C1Point(Start, Kind, anchorRow, anchorColumn);
            string end = style == ExcelReferenceStyle.A1
                ? FormatA1Point(End, Kind)
                : FormatR1C1Point(End, Kind, anchorRow, anchorColumn);
            string address = Kind == ExcelReferenceKind.Cell ? start : start + ":" + end;
            return string.IsNullOrWhiteSpace(Qualifier) ? address : Qualifier + "!" + address;
        }

        /// <inheritdoc />
        public override string ToString() => ToString(ExcelReferenceStyle.A1);

        /// <summary>Returns whether this rectangular reference contains a cell.</summary>
        public bool Contains(int row, int column) {
            ValidateCell(row, column);
            GetBounds(out int firstRow, out int firstColumn, out int lastRow, out int lastColumn);
            return row >= firstRow && row <= lastRow && column >= firstColumn && column <= lastColumn;
        }

        /// <summary>Returns whether two references on the same qualifier intersect.</summary>
        public bool Intersects(ExcelReference other) {
            if (other == null) throw new ArgumentNullException(nameof(other));
            if (!SameQualifier(other)) return false;
            GetBounds(out int ar1, out int ac1, out int ar2, out int ac2);
            other.GetBounds(out int br1, out int bc1, out int br2, out int bc2);
            return ar1 <= br2 && ar2 >= br1 && ac1 <= bc2 && ac2 >= bc1;
        }

        /// <summary>Returns the rectangular intersection, or null when the references do not overlap.</summary>
        public ExcelReference? Intersect(ExcelReference other) {
            if (other == null) throw new ArgumentNullException(nameof(other));
            if (!SameQualifier(other)) return null;
            GetBounds(out int ar1, out int ac1, out int ar2, out int ac2);
            other.GetBounds(out int br1, out int bc1, out int br2, out int bc2);
            int r1 = Math.Max(ar1, br1);
            int c1 = Math.Max(ac1, bc1);
            int r2 = Math.Min(ar2, br2);
            int c2 = Math.Min(ac2, bc2);
            return r1 > r2 || c1 > c2 ? null : FromBounds(r1, c1, r2, c2, Qualifier);
        }

        /// <summary>Returns the smallest rectangle containing both references.</summary>
        public ExcelReference BoundingUnion(ExcelReference other) {
            if (other == null) throw new ArgumentNullException(nameof(other));
            if (!SameQualifier(other)) {
                throw new InvalidOperationException("Range union requires matching workbook and worksheet qualifiers.");
            }
            GetBounds(out int ar1, out int ac1, out int ar2, out int ac2);
            other.GetBounds(out int br1, out int bc1, out int br2, out int bc2);
            return FromBounds(Math.Min(ar1, br1), Math.Min(ac1, bc1), Math.Max(ar2, br2), Math.Max(ac2, bc2), Qualifier);
        }

        /// <summary>Subtracts an overlapping rectangle and returns up to four non-overlapping rectangles.</summary>
        public IReadOnlyList<ExcelReference> Except(ExcelReference other) {
            ExcelReference? overlap = Intersect(other);
            if (overlap == null) return new[] { this };

            GetBounds(out int r1, out int c1, out int r2, out int c2);
            overlap.GetBounds(out int or1, out int oc1, out int or2, out int oc2);
            var result = new List<ExcelReference>(4);
            AddBounds(result, r1, c1, or1 - 1, c2, Qualifier);
            AddBounds(result, or2 + 1, c1, r2, c2, Qualifier);
            AddBounds(result, or1, c1, or2, oc1 - 1, Qualifier);
            AddBounds(result, or1, oc2 + 1, or2, c2, Qualifier);
            return new ReadOnlyCollection<ExcelReference>(result);
        }

        /// <summary>Returns a translated reference and validates worksheet boundaries.</summary>
        public ExcelReference Offset(int rowOffset, int columnOffset) {
            int startRow = Start.Row == 0 ? 0 : checked(Start.Row + rowOffset);
            int endRow = End.Row == 0 ? 0 : checked(End.Row + rowOffset);
            int startColumn = Start.Column == 0 ? 0 : checked(Start.Column + columnOffset);
            int endColumn = End.Column == 0 ? 0 : checked(End.Column + columnOffset);
            ValidateEndpoint(startRow, startColumn, Kind);
            ValidateEndpoint(endRow, endColumn, Kind);
            return new ExcelReference(
                Kind,
                Qualifier,
                new ExcelReferencePoint(startRow, startColumn, Start.RowAbsolute, Start.ColumnAbsolute),
                new ExcelReferencePoint(endRow, endColumn, End.RowAbsolute, End.ColumnAbsolute));
        }

        /// <inheritdoc />
        public bool Equals(ExcelReference? other) =>
            other != null
            && Kind == other.Kind
            && string.Equals(NormalizeQualifierForComparison(Qualifier), NormalizeQualifierForComparison(other.Qualifier), StringComparison.OrdinalIgnoreCase)
            && Start.Equals(other.Start)
            && End.Equals(other.End);

        /// <inheritdoc />
        public override bool Equals(object? obj) => Equals(obj as ExcelReference);

        /// <inheritdoc />
        public override int GetHashCode() {
            unchecked {
                int hash = (int)Kind;
                hash = (hash * 397) ^ StringComparer.OrdinalIgnoreCase.GetHashCode(NormalizeQualifierForComparison(Qualifier));
                hash = (hash * 397) ^ Start.GetHashCode();
                return (hash * 397) ^ End.GetHashCode();
            }
        }

        internal void GetBounds(out int firstRow, out int firstColumn, out int lastRow, out int lastColumn) {
            firstRow = Kind == ExcelReferenceKind.WholeColumn ? 1 : Math.Min(Start.Row, End.Row);
            lastRow = Kind == ExcelReferenceKind.WholeColumn ? A1.MaxRows : Math.Max(Start.Row, End.Row);
            firstColumn = Kind == ExcelReferenceKind.WholeRow ? 1 : Math.Min(Start.Column, End.Column);
            lastColumn = Kind == ExcelReferenceKind.WholeRow ? A1.MaxColumns : Math.Max(Start.Column, End.Column);
        }

        internal ExcelReference WithCoordinates(
            ExcelReferenceKind kind,
            int startRow,
            int startColumn,
            int endRow,
            int endColumn,
            bool? startRowAbsolute = null,
            bool? startColumnAbsolute = null,
            bool? endRowAbsolute = null,
            bool? endColumnAbsolute = null) {
            ValidateEndpoint(startRow, startColumn, kind);
            ValidateEndpoint(endRow, endColumn, kind);
            return new ExcelReference(
                kind,
                Qualifier,
                new ExcelReferencePoint(
                    startRow,
                    startColumn,
                    startRowAbsolute ?? Start.RowAbsolute,
                    startColumnAbsolute ?? Start.ColumnAbsolute),
                new ExcelReferencePoint(
                    endRow,
                    endColumn,
                    endRowAbsolute ?? End.RowAbsolute,
                    endColumnAbsolute ?? End.ColumnAbsolute));
        }

        private bool SameQualifier(ExcelReference other) =>
            string.Equals(NormalizeQualifierForComparison(Qualifier), NormalizeQualifierForComparison(other.Qualifier), StringComparison.OrdinalIgnoreCase);

        internal static bool TryGetThreeDimensionalSheetRange(
            string? qualifier,
            out string firstSheetName,
            out string lastSheetName) {
            firstSheetName = string.Empty;
            lastSheetName = string.Empty;
            string value = qualifier?.Trim() ?? string.Empty;
            if (value.Length == 0) return false;

            int quotedSeparator = value.IndexOf("':'", StringComparison.Ordinal);
            if (quotedSeparator >= 0) {
                firstSheetName = UnquoteQualifierToken(value.Substring(0, quotedSeparator + 1));
                lastSheetName = UnquoteQualifierToken(value.Substring(quotedSeparator + 2));
            } else {
                string normalized = UnquoteQualifierToken(value);
                int separator = normalized.IndexOf(':');
                if (separator <= 0 || separator != normalized.LastIndexOf(':')) return false;
                firstSheetName = normalized.Substring(0, separator);
                lastSheetName = normalized.Substring(separator + 1);
            }

            return firstSheetName.Length > 0
                && lastSheetName.Length > 0
                && !firstSheetName.StartsWith("[", StringComparison.Ordinal)
                && !lastSheetName.StartsWith("[", StringComparison.Ordinal);
        }

        private static string NormalizeQualifierForComparison(string? qualifier) {
            string value = qualifier?.Trim() ?? string.Empty;
            int separatelyQuotedSeparator = value.IndexOf("':'", StringComparison.Ordinal);
            if (separatelyQuotedSeparator >= 0) {
                return UnquoteQualifierToken(value.Substring(0, separatelyQuotedSeparator + 1))
                    + ":"
                    + UnquoteQualifierToken(value.Substring(separatelyQuotedSeparator + 2));
            }
            return UnquoteQualifierToken(value);
        }

        private static string UnquoteQualifierToken(string value) {
            string result = value.Trim();
            if (result.Length >= 2 && result[0] == '\'' && result[result.Length - 1] == '\'') {
                result = result.Substring(1, result.Length - 2).Replace("''", "'");
            }
            return result;
        }

        private static ExcelReference FromBounds(int r1, int c1, int r2, int c2, string? qualifier) {
            ValidateCell(r1, c1);
            ValidateCell(r2, c2);
            if (r1 == 1 && r2 == A1.MaxRows) {
                return new ExcelReference(
                    ExcelReferenceKind.WholeColumn,
                    qualifier,
                    new ExcelReferencePoint(0, c1, false, false),
                    new ExcelReferencePoint(0, c2, false, false));
            }
            if (c1 == 1 && c2 == A1.MaxColumns) {
                return new ExcelReference(
                    ExcelReferenceKind.WholeRow,
                    qualifier,
                    new ExcelReferencePoint(r1, 0, false, false),
                    new ExcelReferencePoint(r2, 0, false, false));
            }
            var start = new ExcelReferencePoint(r1, c1, false, false);
            var end = new ExcelReferencePoint(r2, c2, false, false);
            return new ExcelReference(r1 == r2 && c1 == c2 ? ExcelReferenceKind.Cell : ExcelReferenceKind.Range, qualifier, start, end);
        }

        private static void AddBounds(List<ExcelReference> ranges, int r1, int c1, int r2, int c2, string? qualifier) {
            if (r1 <= r2 && c1 <= c2) ranges.Add(FromBounds(r1, c1, r2, c2, qualifier));
        }

        private static bool TryParseA1(string address, string? qualifier, out ExcelReference? reference) {
            reference = null;
            int separator = address.IndexOf(':');
            if (separator != address.LastIndexOf(':')) return false;
            string first = separator < 0 ? address : address.Substring(0, separator);
            string second = separator < 0 ? first : address.Substring(separator + 1);

            if (TryParseA1Cell(first, out ExcelReferencePoint start)
                && TryParseA1Cell(second, out ExcelReferencePoint end)) {
                reference = new ExcelReference(separator < 0 ? ExcelReferenceKind.Cell : ExcelReferenceKind.Range, qualifier, start, end);
                return true;
            }

            if (separator >= 0
                && TryParseA1Column(first, out start)
                && TryParseA1Column(second, out end)) {
                reference = new ExcelReference(ExcelReferenceKind.WholeColumn, qualifier, start, end);
                return true;
            }

            if (separator >= 0
                && TryParseA1Row(first, out start)
                && TryParseA1Row(second, out end)) {
                reference = new ExcelReference(ExcelReferenceKind.WholeRow, qualifier, start, end);
                return true;
            }

            return false;
        }

        private static bool TryParseA1Cell(string value, out ExcelReferencePoint point) {
            point = default;
            int index = 0;
            bool columnAbsolute = ReadDollar(value, ref index);
            int columnStart = index;
            while (index < value.Length && IsAsciiLetter(value[index])) index++;
            if (index == columnStart || index - columnStart > 3) return false;
            int column = A1.ColumnLettersToIndex(value.Substring(columnStart, index - columnStart));
            bool rowAbsolute = ReadDollar(value, ref index);
            int rowStart = index;
            while (index < value.Length && char.IsDigit(value[index])) index++;
            if (index != value.Length || rowStart == index
                || !int.TryParse(value.Substring(rowStart), NumberStyles.None, CultureInfo.InvariantCulture, out int row)
                || row < 1 || row > A1.MaxRows || column < 1 || column > A1.MaxColumns) {
                return false;
            }

            point = new ExcelReferencePoint(row, column, rowAbsolute, columnAbsolute);
            return true;
        }

        private static bool TryParseA1Column(string value, out ExcelReferencePoint point) {
            point = default;
            int index = 0;
            bool absolute = ReadDollar(value, ref index);
            int start = index;
            while (index < value.Length && IsAsciiLetter(value[index])) index++;
            if (index != value.Length || start == index || index - start > 3) return false;
            int column = A1.ColumnLettersToIndex(value.Substring(start));
            if (column < 1 || column > A1.MaxColumns) return false;
            point = new ExcelReferencePoint(0, column, false, absolute);
            return true;
        }

        private static bool TryParseA1Row(string value, out ExcelReferencePoint point) {
            point = default;
            int index = 0;
            bool absolute = ReadDollar(value, ref index);
            if (!int.TryParse(value.Substring(index), NumberStyles.None, CultureInfo.InvariantCulture, out int row)
                || row < 1 || row > A1.MaxRows) return false;
            point = new ExcelReferencePoint(row, 0, absolute, false);
            return true;
        }

        private static bool TryParseR1C1(
            string address,
            string? qualifier,
            int anchorRow,
            int anchorColumn,
            out ExcelReference? reference) {
            reference = null;
            int separator = address.IndexOf(':');
            if (separator != address.LastIndexOf(':')) return false;
            string first = separator < 0 ? address : address.Substring(0, separator);
            string second = separator < 0 ? first : address.Substring(separator + 1);

            if (!TryParseR1C1Point(first, anchorRow, anchorColumn, out ExcelReferencePoint start, out bool hasRow, out bool hasColumn)
                || !TryParseR1C1Point(second, anchorRow, anchorColumn, out ExcelReferencePoint end, out bool endHasRow, out bool endHasColumn)
                || hasRow != endHasRow || hasColumn != endHasColumn) {
                return false;
            }

            ExcelReferenceKind kind = !hasRow
                ? ExcelReferenceKind.WholeColumn
                : !hasColumn
                    ? ExcelReferenceKind.WholeRow
                    : separator < 0 ? ExcelReferenceKind.Cell : ExcelReferenceKind.Range;
            reference = new ExcelReference(kind, qualifier, start, end);
            return true;
        }

        private static bool TryParseR1C1Point(
            string value,
            int anchorRow,
            int anchorColumn,
            out ExcelReferencePoint point,
            out bool hasRow,
            out bool hasColumn) {
            point = default;
            hasRow = hasColumn = false;
            int index = 0;
            int row = 0;
            int column = 0;
            bool rowAbsolute = false;
            bool columnAbsolute = false;

            if (index < value.Length && (value[index] == 'R' || value[index] == 'r')) {
                hasRow = true;
                index++;
                if (!TryParseR1C1Coordinate(value, ref index, anchorRow, out row, out rowAbsolute)) return false;
            }

            if (index < value.Length && (value[index] == 'C' || value[index] == 'c')) {
                hasColumn = true;
                index++;
                if (!TryParseR1C1Coordinate(value, ref index, anchorColumn, out column, out columnAbsolute)) return false;
            }

            if (index != value.Length || (!hasRow && !hasColumn)
                || (hasRow && (row < 1 || row > A1.MaxRows))
                || (hasColumn && (column < 1 || column > A1.MaxColumns))) {
                return false;
            }

            point = new ExcelReferencePoint(row, column, rowAbsolute, columnAbsolute);
            return true;
        }

        private static bool TryParseR1C1Coordinate(string value, ref int index, int anchor, out int resolved, out bool absolute) {
            resolved = anchor;
            absolute = false;
            if (index >= value.Length || value[index] == 'R' || value[index] == 'r' || value[index] == 'C' || value[index] == 'c') {
                return true;
            }

            if (value[index] == '[') {
                int close = value.IndexOf(']', index + 1);
                if (close < 0
                    || !int.TryParse(value.Substring(index + 1, close - index - 1), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int offset)) {
                    return false;
                }
                resolved = anchor + offset;
                index = close + 1;
                return true;
            }

            int start = index;
            while (index < value.Length && char.IsDigit(value[index])) index++;
            if (start == index
                || !int.TryParse(value.Substring(start, index - start), NumberStyles.None, CultureInfo.InvariantCulture, out resolved)) {
                return false;
            }
            absolute = true;
            return true;
        }

        private static string FormatA1Point(ExcelReferencePoint point, ExcelReferenceKind kind) {
            if (kind == ExcelReferenceKind.WholeRow) return (point.RowAbsolute ? "$" : string.Empty) + point.Row.ToString(CultureInfo.InvariantCulture);
            string column = (point.ColumnAbsolute ? "$" : string.Empty) + A1.ColumnIndexToLetters(point.Column);
            if (kind == ExcelReferenceKind.WholeColumn) return column;
            return column + (point.RowAbsolute ? "$" : string.Empty) + point.Row.ToString(CultureInfo.InvariantCulture);
        }

        private static string FormatR1C1Point(ExcelReferencePoint point, ExcelReferenceKind kind, int anchorRow, int anchorColumn) {
            string row = kind == ExcelReferenceKind.WholeColumn ? string.Empty : FormatR1C1Coordinate('R', point.Row, point.RowAbsolute, anchorRow);
            string column = kind == ExcelReferenceKind.WholeRow ? string.Empty : FormatR1C1Coordinate('C', point.Column, point.ColumnAbsolute, anchorColumn);
            return row + column;
        }

        private static string FormatR1C1Coordinate(char prefix, int value, bool absolute, int anchor) {
            if (absolute) return prefix + value.ToString(CultureInfo.InvariantCulture);
            int offset = value - anchor;
            return offset == 0 ? prefix.ToString() : prefix + "[" + offset.ToString(CultureInfo.InvariantCulture) + "]";
        }

        private static bool TrySplitQualifier(string value, out string? qualifier, out string address) {
            int separator = -1;
            bool quoted = false;
            for (int index = 0; index < value.Length; index++) {
                if (value[index] == '\'') {
                    if (quoted && index + 1 < value.Length && value[index + 1] == '\'') {
                        index++;
                        continue;
                    }
                    quoted = !quoted;
                } else if (value[index] == '!' && !quoted) {
                    if (separator >= 0) {
                        qualifier = null;
                        address = value;
                        return false;
                    }
                    separator = index;
                }
            }
            if (quoted) {
                qualifier = null;
                address = value;
                return false;
            }
            qualifier = separator < 0 ? null : value.Substring(0, separator);
            address = separator < 0 ? value : value.Substring(separator + 1);
            return qualifier == null || IsValidQualifier(qualifier);
        }

        private static bool IsValidQualifier(string qualifier) {
            if (qualifier.Length == 0) return false;
            if (qualifier[0] == '\'') {
                int quotedSpanSeparator = qualifier.IndexOf("':'", StringComparison.Ordinal);
                if (quotedSpanSeparator >= 0) {
                    return quotedSpanSeparator == qualifier.LastIndexOf("':'", StringComparison.Ordinal)
                        && IsValidQuotedQualifierToken(qualifier.Substring(0, quotedSpanSeparator + 1))
                        && IsValidQuotedQualifierToken(qualifier.Substring(quotedSpanSeparator + 2));
                }
                return IsValidQuotedQualifierToken(qualifier);
            }
            if (qualifier.IndexOf('\'') >= 0 || qualifier.Any(char.IsWhiteSpace)) return false;

            int workbookStart = qualifier.IndexOf('[');
            int workbookEnd = qualifier.IndexOf(']');
            if ((workbookStart < 0) != (workbookEnd < 0)) return false;
            if (workbookStart >= 0
                && (workbookStart != 0
                    || workbookEnd <= workbookStart + 1
                    || workbookEnd == qualifier.Length - 1
                    || qualifier.IndexOf('[', workbookStart + 1) >= 0
                    || qualifier.IndexOf(']', workbookEnd + 1) >= 0)) return false;

            string sheetNames = workbookEnd >= 0 ? qualifier.Substring(workbookEnd + 1) : qualifier;
            if (sheetNames.IndexOfAny(new[] { '!', '[', ']', '/', '\\', '?', '*' }) >= 0) return false;
            int rangeSeparator = sheetNames.IndexOf(':');
            return rangeSeparator < 0
                ? sheetNames.Length > 0
                : rangeSeparator > 0
                    && rangeSeparator < sheetNames.Length - 1
                    && sheetNames.IndexOf(':', rangeSeparator + 1) < 0;
        }

        private static bool IsValidQuotedQualifierToken(string token) {
            if (token.Length < 3 || token[0] != '\'' || token[token.Length - 1] != '\'') return false;
            for (int index = 1; index < token.Length - 1; index++) {
                if (token[index] != '\'') continue;
                if (index + 1 >= token.Length - 1 || token[index + 1] != '\'') return false;
                index++;
            }
            return true;
        }

        private static bool ReadDollar(string value, ref int index) {
            if (index < value.Length && value[index] == '$') {
                index++;
                return true;
            }
            return false;
        }

        private static bool IsAsciiLetter(char value) =>
            (value >= 'A' && value <= 'Z') || (value >= 'a' && value <= 'z');

        private static void ValidateCell(int row, int column) {
            if (row < 1 || row > A1.MaxRows) throw new ArgumentOutOfRangeException(nameof(row));
            if (column < 1 || column > A1.MaxColumns) throw new ArgumentOutOfRangeException(nameof(column));
        }

        private static void ValidateEndpoint(int row, int column, ExcelReferenceKind kind) {
            if (kind != ExcelReferenceKind.WholeColumn && (row < 1 || row > A1.MaxRows)) throw new ArgumentOutOfRangeException(nameof(row));
            if (kind != ExcelReferenceKind.WholeRow && (column < 1 || column > A1.MaxColumns)) throw new ArgumentOutOfRangeException(nameof(column));
        }
    }
}
