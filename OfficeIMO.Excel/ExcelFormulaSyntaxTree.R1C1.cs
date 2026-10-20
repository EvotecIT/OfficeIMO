using System.Text.RegularExpressions;

namespace OfficeIMO.Excel;

public sealed partial class ExcelFormulaSyntaxTree {
    private static readonly Regex R1C1ReferencePattern = new(
        @"\G(?:(?:'(?:[^']|'')+'|(?:\[[^\]]+\])?[\p{L}_\\][\p{L}\p{N}_.\\]*)!)?(?:R(?:\[-?\d+\]|\d+)?C(?:\[-?\d+\]|\d+)?|R(?:\[-?\d+\]|\d+)?|C(?:\[-?\d+\]|\d+)?)(?::(?:R(?:\[-?\d+\]|\d+)?C(?:\[-?\d+\]|\d+)?|R(?:\[-?\d+\]|\d+)?|C(?:\[-?\d+\]|\d+)?))?",
        RegexOptions.CultureInvariant | RegexOptions.IgnoreCase, TimeSpan.FromSeconds(1));

    /// <summary>Parses references in the selected notation, resolving relative R1C1 coordinates at the supplied cell.</summary>
    public static ExcelFormulaSyntaxTree Parse(string formula, ExcelReferenceStyle inputStyle, int anchorRow = 1, int anchorColumn = 1) {
        if (inputStyle == ExcelReferenceStyle.A1) return Parse(formula);
        if (inputStyle != ExcelReferenceStyle.R1C1) throw new ArgumentOutOfRangeException(nameof(inputStyle));
        if (formula == null) throw new ArgumentNullException(nameof(formula));
        if (anchorRow < 1 || anchorRow > 1048576) throw new ArgumentOutOfRangeException(nameof(anchorRow));
        if (anchorColumn < 1 || anchorColumn > 16384) throw new ArgumentOutOfRangeException(nameof(anchorColumn));
        var nodes = new List<ExcelFormulaSyntaxNode>();
        int cursor = 0, textStart = 0;
        while (cursor < formula.Length) {
            if (formula[cursor] == '"') {
                AddText(nodes, formula, textStart, cursor - textStart);
                int start = cursor++;
                while (cursor < formula.Length) {
                    if (formula[cursor++] != '"') continue;
                    if (cursor < formula.Length && formula[cursor] == '"') { cursor++; continue; }
                    break;
                }
                nodes.Add(new ExcelFormulaStringSyntax(formula.Substring(start, cursor - start)));
                textStart = cursor;
                continue;
            }
            Match match = R1C1ReferencePattern.Match(formula, cursor);
            if (match.Success && match.Index == cursor
                && (cursor == 0 || !IsR1C1IdentifierCharacter(formula[cursor - 1]))
                && (cursor + match.Length == formula.Length || !IsR1C1IdentifierCharacter(formula[cursor + match.Length]))
                && !IsR1C1FunctionCall(formula, cursor + match.Length)) {
                if (!ExcelReference.TryParse(match.Value, out ExcelReference? reference, inputStyle, anchorRow, anchorColumn))
                    throw new ArgumentException("R1C1 reference is outside the supported worksheet grid.", nameof(formula));
                AddText(nodes, formula, textStart, cursor - textStart);
                nodes.Add(new ExcelFormulaReferenceSyntax(match.Value, reference!));
                cursor += match.Length;
                textStart = cursor;
            } else if (formula[cursor] == '\'') {
                // A quoted sheet qualifier may precede a defined name instead of a cell.
                // Keep the complete token opaque so reference-like text inside it is not rewritten.
                cursor++;
                while (cursor < formula.Length) {
                    if (formula[cursor++] != '\'') continue;
                    if (cursor < formula.Length && formula[cursor] == '\'') { cursor++; continue; }
                    break;
                }
            } else cursor++;
        }
        AddText(nodes, formula, textStart, formula.Length - textStart);
        return new ExcelFormulaSyntaxTree(formula, nodes.AsReadOnly());
    }

    private static bool IsR1C1IdentifierCharacter(char value) => char.IsLetterOrDigit(value) || value is '_' or '.' or '\\' or '[' or ']' or '!';

    private static bool IsR1C1FunctionCall(string formula, int offset) {
        while (offset < formula.Length && char.IsWhiteSpace(formula[offset])) offset++;
        return offset < formula.Length && formula[offset] == '(';
    }
}
