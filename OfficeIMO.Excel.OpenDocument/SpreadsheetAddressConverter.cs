using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Excel.OpenDocument;

internal static class SpreadsheetAddressConverter {
    private static readonly Regex ExcelReference = new Regex(
        @"(?<![A-Za-z0-9_.'!])(?<start>\$?[A-Z]{1,3}\$?[1-9][0-9]*)(?::(?<end>\$?[A-Z]{1,3}\$?[1-9][0-9]*))?(?![A-Za-z0-9_])",
        RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
    private static readonly Regex QualifiedExcelReference = new Regex(
        @"(?<![A-Za-z0-9_.])(?<sheet>'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!(?<start>\$?[A-Z]{1,3}\$?[1-9][0-9]*)(?::(?<end>\$?[A-Z]{1,3}\$?[1-9][0-9]*))?(?![A-Za-z0-9_])",
        RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    internal static string ExcelFormulaToOpenFormula(string formula) {
        if (string.IsNullOrWhiteSpace(formula)) return string.Empty;
        string body = formula.Trim();
        if (body.StartsWith("=", StringComparison.Ordinal)) body = body.Substring(1);
        body = ReplaceReferencesOutsideStrings(body);
        return "of:=" + body;
    }

    private static string ReplaceReferencesOutsideStrings(string body) {
        var output = new StringBuilder(body.Length);
        int segmentStart = 0;
        bool quoted = false;
        for (int index = 0; index < body.Length; index++) {
            if (body[index] != '"') continue;
            if (quoted && index + 1 < body.Length && body[index + 1] == '"') { index++; continue; }

            if (!quoted) {
                AppendConvertedReferences(output, body, segmentStart, index - segmentStart);
                segmentStart = index;
                quoted = true;
            } else {
                output.Append(body, segmentStart, index - segmentStart + 1);
                segmentStart = index + 1;
                quoted = false;
            }
        }
        if (segmentStart < body.Length) {
            if (quoted) output.Append(body, segmentStart, body.Length - segmentStart);
            else AppendConvertedReferences(output, body, segmentStart, body.Length - segmentStart);
        }
        return output.ToString();
    }

    private static void AppendConvertedReferences(StringBuilder output, string body, int start, int length) {
        if (length <= 0) return;
        string segment = body.Substring(start, length);
        string converted = QualifiedExcelReference.Replace(segment, match =>
            "[" + ExcelRangeToOpenAddress(match.Value) + "]");
        converted = ExcelReference.Replace(converted, match => {
            string referenceStart = match.Groups["start"].Value;
            string referenceEnd = match.Groups["end"].Value;
            return referenceEnd.Length == 0 ? "[." + referenceStart + "]" : "[." + referenceStart + ":." + referenceEnd + "]";
        });
        AppendOpenFormulaSeparators(output, converted);
    }

    private static void AppendOpenFormulaSeparators(StringBuilder output, string value) {
        bool quotedSheet = false;
        for (int index = 0; index < value.Length; index++) {
            char character = value[index];
            if (character == '\'') {
                output.Append(character);
                if (quotedSheet && index + 1 < value.Length && value[index + 1] == '\'') {
                    output.Append(value[++index]);
                } else {
                    quotedSheet = !quotedSheet;
                }
            } else {
                output.Append(character == ',' && !quotedSheet ? ';' : character);
            }
        }
    }

    internal static string OpenFormulaToExcel(string formula) {
        if (string.IsNullOrWhiteSpace(formula)) return string.Empty;
        string body = formula.Trim();
        int equals = body.IndexOf(":=", StringComparison.Ordinal);
        if (equals >= 0) body = body.Substring(equals + 2);
        else if (body.StartsWith("=", StringComparison.Ordinal)) body = body.Substring(1);

        var output = new StringBuilder(body.Length);
        bool quoted = false;
        for (int index = 0; index < body.Length; index++) {
            char character = body[index];
            if (character == '"') {
                output.Append(character);
                if (quoted && index + 1 < body.Length && body[index + 1] == '"') {
                    output.Append(body[++index]);
                } else {
                    quoted = !quoted;
                }
                continue;
            }
            if (!quoted && character == ';') {
                output.Append(',');
                continue;
            }
            if (quoted || character != '[') {
                output.Append(character);
                continue;
            }
            int close = body.IndexOf(']', index + 1);
            if (close < 0) {
                output.Append(body[index]);
                continue;
            }
            string address = body.Substring(index + 1, close - index - 1);
            output.Append(OpenAddressToExcel(address));
            index = close;
        }
        return output.ToString();
    }

    internal static string ExcelRangeToOpenAddress(string reference, string? defaultSheetName = null) {
        if (string.IsNullOrWhiteSpace(reference)) return string.Empty;
        string value = reference.Trim();
        int bang = FindUnquoted(value, '!');
        string sheet = bang >= 0 ? UnquoteExcelSheet(value.Substring(0, bang)) : (defaultSheetName ?? string.Empty);
        string range = bang >= 0 ? value.Substring(bang + 1) : value;
        if (sheet.Length == 0) return "." + range.Replace(":", ":.");
        string escaped = sheet.Replace("'", "''");
        return "$'" + escaped + "'." + range.Replace(":", ":.");
    }

    internal static string OpenAddressToExcel(string address) {
        if (string.IsNullOrWhiteSpace(address)) return string.Empty;
        string value = address.Trim();
        if (value.StartsWith(".", StringComparison.Ordinal)) return value.Substring(1).Replace(":.", ":");

        int dot = FindSheetSeparator(value);
        if (dot < 0) return value.Replace(":.", ":").TrimStart('.');
        string sheet = value.Substring(0, dot).TrimStart('$');
        if (sheet.Length >= 2 && sheet[0] == '\'' && sheet[sheet.Length - 1] == '\'') {
            sheet = sheet.Substring(1, sheet.Length - 2).Replace("''", "'");
        }
        string range = value.Substring(dot + 1).Replace(":.", ":");
        return "'" + sheet.Replace("'", "''") + "'!" + range;
    }

    internal static string ToA1(int row, int column) {
        if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));
        if (column < 1) throw new ArgumentOutOfRangeException(nameof(column));
        int value = column;
        var letters = new StringBuilder();
        while (value > 0) {
            value--;
            letters.Insert(0, (char)('A' + value % 26));
            value /= 26;
        }
        return letters.ToString() + row.ToString(CultureInfo.InvariantCulture);
    }

    private static int FindSheetSeparator(string value) {
        bool quoted = false;
        for (int index = 0; index < value.Length; index++) {
            if (value[index] == '\'') {
                if (quoted && index + 1 < value.Length && value[index + 1] == '\'') { index++; continue; }
                quoted = !quoted;
            } else if (value[index] == '.' && !quoted) return index;
        }
        return -1;
    }

    private static int FindUnquoted(string value, char character) {
        bool quoted = false;
        for (int index = 0; index < value.Length; index++) {
            if (value[index] == '\'') {
                if (quoted && index + 1 < value.Length && value[index + 1] == '\'') { index++; continue; }
                quoted = !quoted;
            } else if (value[index] == character && !quoted) return index;
        }
        return -1;
    }

    private static string UnquoteExcelSheet(string value) {
        string sheet = value.Trim();
        if (sheet.Length >= 2 && sheet[0] == '\'' && sheet[sheet.Length - 1] == '\'') {
            return sheet.Substring(1, sheet.Length - 2).Replace("''", "'");
        }
        return sheet;
    }
}
