using OfficeIMO.Spreadsheet;

namespace OfficeIMO.OpenDocument;

/// <summary>The comparison subset that Excel and ODF validation rules can exchange without changing meaning.</summary>
internal static class OdsPortableValidationFormula {
    internal static bool IsSupported(SpreadsheetFormulaSyntaxTree syntax) {
        if (!syntax.IsValid) return false;
        var tokens = new List<SpreadsheetFormulaSyntaxNode>();
        if (!TryCollectTokens(syntax.Root, tokens, 0)) return false;
        int cursor = 0;
        return TryReadComparison(tokens, ref cursor, 0) && cursor == tokens.Count;
    }

    private static bool TryCollectTokens(SpreadsheetFormulaSyntaxNode node,
        ICollection<SpreadsheetFormulaSyntaxNode> tokens, int depth) {
        if (depth > 32 || node.Kind is SpreadsheetFormulaSyntaxKind.FunctionCall
            or SpreadsheetFormulaSyntaxKind.InlineArray) return false;
        if (node.Kind == SpreadsheetFormulaSyntaxKind.Token) {
            if (node.TokenKind is SpreadsheetFormulaTokenKind.Prefix or SpreadsheetFormulaTokenKind.Whitespace)
                return true;
            if (node.TokenKind == SpreadsheetFormulaTokenKind.Reference) {
                SpreadsheetRangeReference? reference = node.Reference;
                if (reference == null || !reference.Start.IsCell || reference.Start.SheetName != null
                    || reference.End != null) return false;
            } else if (node.TokenKind is not (SpreadsheetFormulaTokenKind.Operator
                or SpreadsheetFormulaTokenKind.NumberLiteral or SpreadsheetFormulaTokenKind.StringLiteral
                or SpreadsheetFormulaTokenKind.OpenDelimiter or SpreadsheetFormulaTokenKind.CloseDelimiter)) {
                return false;
            }
            tokens.Add(node);
            return true;
        }
        foreach (SpreadsheetFormulaSyntaxNode child in node.Children) {
            if (!TryCollectTokens(child, tokens, depth + 1)) return false;
        }
        return true;
    }

    private static bool TryReadComparison(IReadOnlyList<SpreadsheetFormulaSyntaxNode> tokens,
        ref int cursor, int depth) {
        if (depth > 32) return false;
        int start = cursor;
        if (TakeDelimiter(tokens, ref cursor, "(")
            && TryReadComparison(tokens, ref cursor, depth + 1)
            && TakeDelimiter(tokens, ref cursor, ")")) return true;
        cursor = start;
        if (!TryReadOperand(tokens, ref cursor, left: true, depth)) return false;
        if (cursor == tokens.Count || tokens[cursor].TokenKind != SpreadsheetFormulaTokenKind.Operator
            || tokens[cursor].Text is not ("=" or "<>" or "!=" or "<" or "<=" or ">" or ">=")) return false;
        cursor++;
        return TryReadOperand(tokens, ref cursor, left: false, depth);
    }

    private static bool TryReadOperand(IReadOnlyList<SpreadsheetFormulaSyntaxNode> tokens,
        ref int cursor, bool left, int depth) {
        if (depth > 32) return false;
        if (TakeDelimiter(tokens, ref cursor, "(")) {
            return TryReadOperand(tokens, ref cursor, left, depth + 1)
                && TakeDelimiter(tokens, ref cursor, ")");
        }
        if (cursor == tokens.Count) return false;
        if (left) return TakeKind(tokens, ref cursor, SpreadsheetFormulaTokenKind.Reference);
        bool signed = tokens[cursor].TokenKind == SpreadsheetFormulaTokenKind.Operator
            && tokens[cursor].Text is "+" or "-";
        if (signed) cursor++;
        if (TakeKind(tokens, ref cursor, SpreadsheetFormulaTokenKind.NumberLiteral)) return true;
        return !signed && (TakeKind(tokens, ref cursor, SpreadsheetFormulaTokenKind.StringLiteral)
            || TakeKind(tokens, ref cursor, SpreadsheetFormulaTokenKind.Reference));
    }

    private static bool TakeKind(IReadOnlyList<SpreadsheetFormulaSyntaxNode> tokens,
        ref int cursor, SpreadsheetFormulaTokenKind kind) {
        if (cursor == tokens.Count || tokens[cursor].TokenKind != kind) return false;
        cursor++;
        return true;
    }

    private static bool TakeDelimiter(IReadOnlyList<SpreadsheetFormulaSyntaxNode> tokens,
        ref int cursor, string text) {
        if (cursor == tokens.Count || tokens[cursor].TokenKind != (text == "("
                ? SpreadsheetFormulaTokenKind.OpenDelimiter : SpreadsheetFormulaTokenKind.CloseDelimiter)
            || tokens[cursor].Text != text) return false;
        cursor++;
        return true;
    }
}
