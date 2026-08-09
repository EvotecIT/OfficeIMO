using System;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Spreadsheet;

/// <summary>Identifies the address grammar used by a spreadsheet reference.</summary>
public enum SpreadsheetAddressDialect {
    /// <summary>Excel A1 notation, for example <c>'Data'!$A$1:$B$2</c>.</summary>
    ExcelA1 = 0,

    /// <summary>OpenDocument address notation, for example <c>$'Data'.$A$1:$'Data'.$B$2</c>.</summary>
    OpenDocument = 1,

    /// <summary>
    /// A1 notation without Excel's fixed row and column limits. This is intended for bounded
    /// selections over spreadsheet formats whose grids can extend beyond Excel's worksheet size.
    /// </summary>
    UnboundedA1 = 2
}

/// <summary>One parsed spreadsheet cell, whole-row, or whole-column reference endpoint.</summary>
public sealed class SpreadsheetCellReference {
    internal SpreadsheetCellReference(
        string? sheetName,
        bool isSheetAbsolute,
        int? column,
        bool isColumnAbsolute,
        long? row,
        bool isRowAbsolute) {
        if (!column.HasValue && !row.HasValue) {
            throw new ArgumentException("A spreadsheet reference endpoint requires a row, a column, or both.");
        }
        if (column.HasValue && column.Value < 1) throw new ArgumentOutOfRangeException(nameof(column));
        if (row.HasValue && row.Value < 1) throw new ArgumentOutOfRangeException(nameof(row));

        SheetName = sheetName;
        IsSheetAbsolute = isSheetAbsolute;
        Column = column;
        IsColumnAbsolute = isColumnAbsolute;
        Row = row;
        IsRowAbsolute = isRowAbsolute;
    }

    /// <summary>Gets the decoded worksheet name, or <see langword="null"/> for the current sheet.</summary>
    public string? SheetName { get; }

    /// <summary>Gets whether the authored sheet locator was absolute.</summary>
    public bool IsSheetAbsolute { get; }

    /// <summary>Gets the one-based column index, or <see langword="null"/> for a whole-row endpoint.</summary>
    public int? Column { get; }

    /// <summary>Gets whether the column component is absolute.</summary>
    public bool IsColumnAbsolute { get; }

    /// <summary>Gets the one-based row index, or <see langword="null"/> for a whole-column endpoint.</summary>
    public long? Row { get; }

    /// <summary>Gets whether the row component is absolute.</summary>
    public bool IsRowAbsolute { get; }

    /// <summary>Gets whether this endpoint identifies one cell rather than a whole row or column.</summary>
    public bool IsCell => Column.HasValue && Row.HasValue;
}

/// <summary>
/// Immutable, typed spreadsheet address syntax. The parser understands quoting, absolute markers,
/// whole-row and whole-column ranges without splitting semantic text on punctuation.
/// </summary>
public sealed class SpreadsheetRangeReference {
    private SpreadsheetRangeReference(SpreadsheetCellReference start, SpreadsheetCellReference? end) {
        Start = start;
        End = end;
    }

    /// <summary>Gets the first endpoint.</summary>
    public SpreadsheetCellReference Start { get; }

    /// <summary>Gets the optional second endpoint.</summary>
    public SpreadsheetCellReference? End { get; }

    /// <summary>Gets whether the syntax represents a range.</summary>
    public bool IsRange => End != null;

    /// <summary>Parses a complete address using the requested grammar.</summary>
    public static SpreadsheetRangeReference Parse(string text, SpreadsheetAddressDialect dialect) {
        if (text == null) throw new ArgumentNullException(nameof(text));
        if (!TryParse(text, dialect, out SpreadsheetRangeReference? reference)) {
            throw new FormatException($"'{text}' is not a valid {dialect} spreadsheet address.");
        }
        return reference!;
    }

    /// <summary>Attempts to parse a complete address using the requested grammar.</summary>
    public static bool TryParse(
        string? text,
        SpreadsheetAddressDialect dialect,
        out SpreadsheetRangeReference? reference) {
        reference = null;
        if (string.IsNullOrWhiteSpace(text)) return false;
        string value = text!.Trim();
        if (dialect == SpreadsheetAddressDialect.ExcelA1 && value[0] == '=') value = value.Substring(1).Trim();
        if (value.Length == 0) return false;

        int cursor = 0;
        if (!TryReadEndpoint(value, ref cursor, dialect, allowPartial: true, allowImplicitCurrentSheet: false,
                out SpreadsheetCellReference? start)) return false;
        SpreadsheetCellReference? end = null;
        if (cursor < value.Length && value[cursor] == ':') {
            cursor++;
            if (!TryReadEndpoint(value, ref cursor, dialect, allowPartial: true, allowImplicitCurrentSheet: true, out end)) return false;
            end = InheritA1SheetQualifier(dialect, start!, end!);
        }
        if (cursor != value.Length) return false;
        if (end == null && !start!.IsCell) return false;
        if (end != null && (start!.Column.HasValue != end.Column.HasValue || start.Row.HasValue != end.Row.HasValue)) return false;
        reference = new SpreadsheetRangeReference(start!, end);
        return true;
    }

    /// <summary>Formats this address using the requested grammar.</summary>
    public string Format(SpreadsheetAddressDialect dialect) {
        if (!TryFormat(dialect, out string formatted)) {
            throw new InvalidOperationException($"The address cannot be represented safely in {dialect} syntax.");
        }
        return formatted;
    }

    /// <summary>Attempts to format this address without changing relative-sheet or cross-sheet range semantics.</summary>
    public bool TryFormat(SpreadsheetAddressDialect dialect, out string formatted) {
        formatted = string.Empty;
        if (UsesA1Syntax(dialect) && !CanRepresentInA1Syntax()) return false;
        var output = new StringBuilder();
        AppendEndpoint(output, Start, dialect, includeSheet: true);
        if (End != null) {
            output.Append(':');
            bool sameSheet = string.Equals(Start.SheetName, End.SheetName, StringComparison.Ordinal)
                && Start.IsSheetAbsolute == End.IsSheetAbsolute;
            // OpenDocument's leading dot always means the current sheet; it does not inherit the
            // first range endpoint. Repeat an authored qualifier to preserve same-sheet ranges.
            bool includeSheet = dialect == SpreadsheetAddressDialect.OpenDocument && End.SheetName != null
                ? true
                : !sameSheet;
            AppendEndpoint(output, End, dialect, includeSheet);
        }
        formatted = output.ToString();
        return true;
    }

    /// <summary>Formats the first cell using the requested grammar.</summary>
    public string FormatBaseCell(SpreadsheetAddressDialect dialect) {
        if (!Start.IsCell) throw new InvalidOperationException("A whole-row or whole-column range has no base cell.");
        var output = new StringBuilder();
        AppendEndpoint(output, Start, dialect, includeSheet: true);
        return output.ToString();
    }

    /// <inheritdoc />
    public override string ToString() => Format(SpreadsheetAddressDialect.ExcelA1);

    internal static bool TryReadExcelAt(
        string text,
        int start,
        out SpreadsheetRangeReference? reference,
        out int consumed) {
        reference = null;
        consumed = 0;
        if (text == null || start < 0 || start >= text.Length) return false;
        if (start > 0 && IsIdentifierCharacter(text[start - 1])) return false;

        int cursor = start;
        if (!TryReadEndpoint(text, ref cursor, SpreadsheetAddressDialect.ExcelA1, allowPartial: true, allowImplicitCurrentSheet: false,
                out SpreadsheetCellReference? first)) return false;
        SpreadsheetCellReference? second = null;
        if (cursor < text.Length && text[cursor] == ':') {
            cursor++;
            if (!TryReadEndpoint(text, ref cursor, SpreadsheetAddressDialect.ExcelA1, allowPartial: true,
                    allowImplicitCurrentSheet: false, out second)) return false;
            second = InheritA1SheetQualifier(SpreadsheetAddressDialect.ExcelA1, first!, second!);
        }
        if (second == null && !first!.IsCell) return false;
        if (second != null && (first!.Column.HasValue != second.Column.HasValue || first.Row.HasValue != second.Row.HasValue)) return false;
        if (cursor < text.Length && IsIdentifierCharacter(text[cursor])) return false;

        // A token such as LOG10 is a valid A1 coordinate in isolation, but LOG10(...) is a function call.
        int next = cursor;
        while (next < text.Length && char.IsWhiteSpace(text[next])) next++;
        if (second == null && first!.SheetName == null && next < text.Length && text[next] == '(') return false;

        reference = new SpreadsheetRangeReference(first!, second);
        consumed = cursor - start;
        return consumed > 0;
    }

    private static bool TryReadEndpoint(
        string text,
        ref int cursor,
        SpreadsheetAddressDialect dialect,
        bool allowPartial,
        bool allowImplicitCurrentSheet,
        out SpreadsheetCellReference? endpoint) {
        endpoint = null;
        int original = cursor;
        string? sheetName = null;
        bool sheetAbsolute = false;

        if (UsesA1Syntax(dialect)) {
            int qualifierCursor = cursor;
            if (TryReadExcelSheetQualifier(text, ref qualifierCursor, out string? parsedSheet)) {
                sheetName = parsedSheet;
                // Excel sheet qualifiers remain fixed when a formula is copied, so their ODF projection is absolute.
                sheetAbsolute = true;
                cursor = qualifierCursor;
            }
        } else {
            int qualifierCursor = cursor;
            if (TryReadOpenDocumentSheetQualifier(text, ref qualifierCursor, out string? parsedSheet, out bool parsedAbsolute)) {
                sheetName = parsedSheet;
                sheetAbsolute = parsedAbsolute;
                cursor = qualifierCursor;
            } else if (cursor < text.Length && text[cursor] == '.') {
                cursor++;
            } else if (!allowImplicitCurrentSheet) {
                cursor = original;
                return false;
            }
        }

        // In a whole-row endpoint such as $1 the leading dollar belongs to the row.
        // Only consume it as a column marker when a column name actually follows.
        bool columnAbsolute = cursor + 1 < text.Length && text[cursor] == '$' && IsAsciiLetter(text[cursor + 1]);
        if (columnAbsolute) cursor++;
        int columnStart = cursor;
        while (cursor < text.Length && IsAsciiLetter(text[cursor])) cursor++;
        int? column = null;
        if (cursor > columnStart) {
            if (!TryColumnNumber(text, columnStart, cursor - columnStart, out int parsedColumn)) {
                cursor = original;
                return false;
            }
            if (dialect == SpreadsheetAddressDialect.ExcelA1 && parsedColumn > 16384) {
                cursor = original;
                return false;
            }
            column = parsedColumn;
        } else if (columnAbsolute) {
            cursor = original;
            return false;
        }

        bool rowAbsolute = cursor < text.Length && text[cursor] == '$';
        if (rowAbsolute) cursor++;
        int rowStart = cursor;
        while (cursor < text.Length && text[cursor] >= '0' && text[cursor] <= '9') cursor++;
        long? row = null;
        if (cursor > rowStart) {
            if (text[rowStart] == '0' || !long.TryParse(text.Substring(rowStart, cursor - rowStart),
                    NumberStyles.None, CultureInfo.InvariantCulture, out long parsedRow) || parsedRow < 1) {
                cursor = original;
                return false;
            }
            if (dialect == SpreadsheetAddressDialect.ExcelA1 && parsedRow > 1048576) {
                cursor = original;
                return false;
            }
            row = parsedRow;
        } else if (rowAbsolute) {
            cursor = original;
            return false;
        }

        if (!column.HasValue && !row.HasValue) {
            cursor = original;
            return false;
        }
        if (!allowPartial && (!column.HasValue || !row.HasValue)) {
            cursor = original;
            return false;
        }
        endpoint = new SpreadsheetCellReference(sheetName, sheetAbsolute, column, columnAbsolute, row, rowAbsolute);
        return true;
    }

    private static bool TryReadExcelSheetQualifier(string text, ref int cursor, out string? sheetName) {
        sheetName = null;
        int original = cursor;
        if (cursor >= text.Length) return false;
        if (text[cursor] == '[') return false; // External workbook references require a separate carrier contract.

        if (text[cursor] == '\'') {
            if (!TryReadQuoted(text, ref cursor, out string? quoted) || cursor >= text.Length || text[cursor] != '!') {
                cursor = original;
                return false;
            }
            cursor++;
            sheetName = quoted;
            return true;
        }

        int nameStart = cursor;
        while (cursor < text.Length && (IsIdentifierCharacter(text[cursor]) || text[cursor] == '.')) cursor++;
        if (cursor == nameStart || cursor >= text.Length || text[cursor] != '!') {
            cursor = original;
            return false;
        }
        sheetName = text.Substring(nameStart, cursor - nameStart);
        cursor++;
        return true;
    }

    private static bool TryReadOpenDocumentSheetQualifier(
        string text,
        ref int cursor,
        out string? sheetName,
        out bool isAbsolute) {
        sheetName = null;
        isAbsolute = false;
        int original = cursor;
        if (cursor < text.Length && text[cursor] == '$') {
            isAbsolute = true;
            cursor++;
        }

        if (cursor < text.Length && text[cursor] == '\'') {
            if (!TryReadQuoted(text, ref cursor, out sheetName)) {
                cursor = original;
                return false;
            }
        } else {
            int start = cursor;
            while (cursor < text.Length && text[cursor] != '.' && text[cursor] != ':' && text[cursor] != ']') {
                char character = text[cursor];
                if (char.IsWhiteSpace(character) || character == '#' || character == '$' || character == '\'') break;
                cursor++;
            }
            if (cursor > start) sheetName = text.Substring(start, cursor - start);
        }

        if (cursor >= text.Length || text[cursor] != '.') {
            cursor = original;
            sheetName = null;
            isAbsolute = false;
            return false;
        }
        cursor++;
        return true;
    }

    private static bool TryReadQuoted(string text, ref int cursor, out string? value) {
        value = null;
        if (cursor >= text.Length || text[cursor] != '\'') return false;
        cursor++;
        var output = new StringBuilder();
        while (cursor < text.Length) {
            if (text[cursor] != '\'') {
                output.Append(text[cursor++]);
                continue;
            }
            if (cursor + 1 < text.Length && text[cursor + 1] == '\'') {
                output.Append('\'');
                cursor += 2;
                continue;
            }
            cursor++;
            value = output.ToString();
            return true;
        }
        return false;
    }

    private static void AppendEndpoint(
        StringBuilder output,
        SpreadsheetCellReference endpoint,
        SpreadsheetAddressDialect dialect,
        bool includeSheet) {
        if (UsesA1Syntax(dialect)) {
            if (includeSheet && endpoint.SheetName != null) {
                output.Append('\'').Append(endpoint.SheetName.Replace("'", "''")).Append("'!");
            }
        } else {
            if (includeSheet && endpoint.SheetName != null) {
                if (endpoint.IsSheetAbsolute) output.Append('$');
                output.Append('\'').Append(endpoint.SheetName.Replace("'", "''")).Append("'.");
            } else {
                output.Append('.');
            }
        }

        if (endpoint.Column.HasValue) {
            if (endpoint.IsColumnAbsolute) output.Append('$');
            AppendColumnName(output, endpoint.Column.Value);
        }
        if (endpoint.Row.HasValue) {
            if (endpoint.IsRowAbsolute) output.Append('$');
            output.Append(endpoint.Row.Value.ToString(CultureInfo.InvariantCulture));
        }
    }

    private static bool TryColumnNumber(string text, int start, int length, out int column) {
        column = 0;
        for (int index = start; index < start + length; index++) {
            char character = char.ToUpperInvariant(text[index]);
            if (character < 'A' || character > 'Z') return false;
            if (column > (int.MaxValue - (character - 'A' + 1)) / 26) return false;
            column = column * 26 + character - 'A' + 1;
        }
        return column > 0;
    }

    private static void AppendColumnName(StringBuilder output, int column) {
        int value = column;
        int length = 0;
        var characters = new char[7];
        while (value > 0) {
            value--;
            characters[characters.Length - ++length] = (char)('A' + value % 26);
            value /= 26;
        }
        output.Append(characters, characters.Length - length, length);
    }

    private static bool IsAsciiLetter(char character) =>
        (character >= 'A' && character <= 'Z') || (character >= 'a' && character <= 'z');

    private static bool IsIdentifierCharacter(char character) =>
        char.IsLetterOrDigit(character) || character == '_' ||
        CharUnicodeInfo.GetUnicodeCategory(character) is UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark;

    private static bool UsesA1Syntax(SpreadsheetAddressDialect dialect) =>
        dialect == SpreadsheetAddressDialect.ExcelA1 || dialect == SpreadsheetAddressDialect.UnboundedA1;

    private bool CanRepresentInA1Syntax() {
        if (Start.SheetName != null && !Start.IsSheetAbsolute) return false;
        if (End == null) return true;
        if (End.SheetName != null && !End.IsSheetAbsolute) return false;
        return string.Equals(Start.SheetName, End.SheetName, StringComparison.Ordinal);
    }

    private static SpreadsheetCellReference InheritA1SheetQualifier(
        SpreadsheetAddressDialect dialect,
        SpreadsheetCellReference first,
        SpreadsheetCellReference second) {
        if (!UsesA1Syntax(dialect) || first.SheetName == null || second.SheetName != null) return second;
        return new SpreadsheetCellReference(
            first.SheetName,
            first.IsSheetAbsolute,
            second.Column,
            second.IsColumnAbsolute,
            second.Row,
            second.IsRowAbsolute);
    }
}