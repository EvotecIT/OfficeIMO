using System;
using System.Collections.Generic;

namespace OfficeIMO.Spreadsheet;

internal static class SpreadsheetFormulaParser {
    private const int MaximumNestingDepth = 128;

    internal static SpreadsheetFormulaSyntaxTree Parse(string text, SpreadsheetFormulaDialect dialect) {
        var diagnostics = new List<SpreadsheetFormulaDiagnostic>();
        int cursor = 0;
        var children = new List<SpreadsheetFormulaSyntaxNode>();
        ReadPrefix(text, dialect, ref cursor, children);
        ParseSequence(text, dialect, ref cursor, null, SpreadsheetFormulaSyntaxKind.Root, children, diagnostics, depth: 0);
        var root = new SpreadsheetFormulaSyntaxNode(
            SpreadsheetFormulaSyntaxKind.Root,
            null,
            text,
            0,
            children);
        return new SpreadsheetFormulaSyntaxTree(text, dialect, root, diagnostics);
    }

    private static void ReadPrefix(
        string text,
        SpreadsheetFormulaDialect dialect,
        ref int cursor,
        ICollection<SpreadsheetFormulaSyntaxNode> children) {
        if (text.Length == 0) return;
        if (dialect == SpreadsheetFormulaDialect.ExcelA1) {
            if (text[0] == '=') {
                children.Add(Token(SpreadsheetFormulaTokenKind.Prefix, "=", 0));
                cursor = 1;
            }
            return;
        }

        int marker = text.IndexOf(":=", StringComparison.Ordinal);
        if (marker >= 0 && marker <= 32 && IsOpenFormulaPrefix(text, marker)) {
            int length = marker + 2;
            children.Add(Token(SpreadsheetFormulaTokenKind.Prefix, text.Substring(0, length), 0));
            cursor = length;
        } else if (text[0] == '=') {
            children.Add(Token(SpreadsheetFormulaTokenKind.Prefix, "=", 0));
            cursor = 1;
        }
    }

    private static bool IsOpenFormulaPrefix(string text, int marker) {
        if (marker == 0) return false;
        for (int index = 0; index < marker; index++) {
            char character = text[index];
            if (!IsIdentifierStart(character) && character != '.' && character != '-') return false;
        }
        return true;
    }

    private static void ParseSequence(
        string text,
        SpreadsheetFormulaDialect dialect,
        ref int cursor,
        char? terminator,
        SpreadsheetFormulaSyntaxKind containerKind,
        ICollection<SpreadsheetFormulaSyntaxNode> children,
        ICollection<SpreadsheetFormulaDiagnostic> diagnostics,
        int depth) {
        if (depth > MaximumNestingDepth) {
            int start = cursor;
            children.Add(Token(SpreadsheetFormulaTokenKind.Unsupported, text.Substring(start), start));
            diagnostics.Add(Error(
                "FORMULA_NESTING_LIMIT",
                $"Formula nesting exceeds the supported limit of {MaximumNestingDepth}.",
                start,
                text.Length - start));
            cursor = text.Length;
            return;
        }
        while (cursor < text.Length) {
            char character = text[cursor];
            if (terminator.HasValue && character == terminator.Value) return;

            if (character == ')' || character == '}') {
                diagnostics.Add(Error(
                    "FORMULA_UNEXPECTED_CLOSING_DELIMITER",
                    $"Unexpected closing delimiter '{character}'.",
                    cursor,
                    1));
                children.Add(Token(SpreadsheetFormulaTokenKind.Unsupported, character.ToString(), cursor));
                cursor++;
                continue;
            }

            if (character == '"') {
                ReadString(text, ref cursor, children, diagnostics);
                continue;
            }

            if (dialect == SpreadsheetFormulaDialect.OpenFormula && character == '[') {
                ReadOpenFormulaReference(text, ref cursor, children, diagnostics);
                continue;
            }

            if (TryReadFunction(text, dialect, ref cursor, children, diagnostics, depth)) continue;

            if (dialect == SpreadsheetFormulaDialect.ExcelA1 &&
                SpreadsheetRangeReference.TryReadExcelAt(text, cursor, out SpreadsheetRangeReference? excelReference, out int consumed)) {
                children.Add(new SpreadsheetFormulaSyntaxNode(
                    SpreadsheetFormulaSyntaxKind.Token,
                    SpreadsheetFormulaTokenKind.Reference,
                    text.Substring(cursor, consumed),
                    cursor,
                    reference: excelReference));
                cursor += consumed;
                continue;
            }

            if (character == '(') {
                ReadGroup(text, dialect, ref cursor, SpreadsheetFormulaSyntaxKind.ParenthesizedExpression,
                    ')', null, children, diagnostics, depth);
                continue;
            }
            if (character == '{') {
                ReadGroup(text, dialect, ref cursor, SpreadsheetFormulaSyntaxKind.InlineArray,
                    '}', null, children, diagnostics, depth);
                continue;
            }

            if (char.IsWhiteSpace(character)) {
                ReadWhitespace(text, dialect, ref cursor, children);
                continue;
            }

            if (TryReadSeparator(text, dialect, ref cursor, containerKind, children, diagnostics)) continue;
            if (TryReadNumber(text, ref cursor, children)) continue;
            if (TryReadError(text, dialect, ref cursor, children, diagnostics)) continue;
            if (TryReadIdentifier(text, ref cursor, children)) continue;
            if (TryReadOperator(text, dialect, ref cursor, children, diagnostics)) continue;

            diagnostics.Add(Error(
                "FORMULA_UNSUPPORTED_SYNTAX",
                $"The token '{character}' is not supported by the {dialect} parser.",
                cursor,
                1));
            children.Add(Token(SpreadsheetFormulaTokenKind.Unsupported, character.ToString(), cursor));
            cursor++;
        }
    }

    private static bool TryReadFunction(
        string text,
        SpreadsheetFormulaDialect dialect,
        ref int cursor,
        ICollection<SpreadsheetFormulaSyntaxNode> destination,
        ICollection<SpreadsheetFormulaDiagnostic> diagnostics,
        int depth) {
        if (!IsIdentifierStart(text[cursor])) return false;
        int nameStart = cursor;
        int nameEnd = ScanIdentifier(text, cursor);
        int open = nameEnd;
        while (open < text.Length && char.IsWhiteSpace(text[open])) open++;
        if (open >= text.Length || text[open] != '(') return false;

        var children = new List<SpreadsheetFormulaSyntaxNode> {
            Token(SpreadsheetFormulaTokenKind.Identifier, text.Substring(nameStart, nameEnd - nameStart), nameStart)
        };
        if (open > nameEnd) {
            children.Add(Token(SpreadsheetFormulaTokenKind.Whitespace, text.Substring(nameEnd, open - nameEnd), nameEnd));
        }
        children.Add(Token(SpreadsheetFormulaTokenKind.OpenDelimiter, "(", open));
        cursor = open + 1;
        ParseSequence(text, dialect, ref cursor, ')', SpreadsheetFormulaSyntaxKind.FunctionCall, children, diagnostics, depth + 1);
        if (cursor < text.Length && text[cursor] == ')') {
            children.Add(Token(SpreadsheetFormulaTokenKind.CloseDelimiter, ")", cursor));
            cursor++;
        } else {
            diagnostics.Add(Error(
                "FORMULA_UNTERMINATED_FUNCTION",
                $"Function '{text.Substring(nameStart, nameEnd - nameStart)}' has no closing parenthesis.",
                nameStart,
                Math.Max(1, cursor - nameStart)));
        }

        destination.Add(new SpreadsheetFormulaSyntaxNode(
            SpreadsheetFormulaSyntaxKind.FunctionCall,
            null,
            text.Substring(nameStart, cursor - nameStart),
            nameStart,
            children,
            name: text.Substring(nameStart, nameEnd - nameStart)));
        return true;
    }

    private static void ReadGroup(
        string text,
        SpreadsheetFormulaDialect dialect,
        ref int cursor,
        SpreadsheetFormulaSyntaxKind kind,
        char close,
        string? name,
        ICollection<SpreadsheetFormulaSyntaxNode> destination,
        ICollection<SpreadsheetFormulaDiagnostic> diagnostics,
        int depth) {
        int start = cursor;
        char open = text[cursor++];
        var children = new List<SpreadsheetFormulaSyntaxNode> {
            Token(SpreadsheetFormulaTokenKind.OpenDelimiter, open.ToString(), start)
        };
        ParseSequence(text, dialect, ref cursor, close, kind, children, diagnostics, depth + 1);
        if (cursor < text.Length && text[cursor] == close) {
            children.Add(Token(SpreadsheetFormulaTokenKind.CloseDelimiter, close.ToString(), cursor));
            cursor++;
        } else {
            diagnostics.Add(Error(
                "FORMULA_UNTERMINATED_GROUP",
                $"The group beginning with '{open}' has no closing '{close}'.",
                start,
                Math.Max(1, cursor - start)));
        }
        destination.Add(new SpreadsheetFormulaSyntaxNode(
            kind,
            null,
            text.Substring(start, cursor - start),
            start,
            children,
            name: name));
    }

    private static void ReadString(
        string text,
        ref int cursor,
        ICollection<SpreadsheetFormulaSyntaxNode> children,
        ICollection<SpreadsheetFormulaDiagnostic> diagnostics) {
        int start = cursor++;
        bool closed = false;
        while (cursor < text.Length) {
            if (text[cursor] != '"') {
                cursor++;
                continue;
            }
            if (cursor + 1 < text.Length && text[cursor + 1] == '"') {
                cursor += 2;
                continue;
            }
            cursor++;
            closed = true;
            break;
        }
        children.Add(Token(SpreadsheetFormulaTokenKind.StringLiteral, text.Substring(start, cursor - start), start));
        if (!closed) {
            diagnostics.Add(Error(
                "FORMULA_UNTERMINATED_STRING",
                "The formula contains an unterminated string literal.",
                start,
                cursor - start));
        }
    }

    private static void ReadOpenFormulaReference(
        string text,
        ref int cursor,
        ICollection<SpreadsheetFormulaSyntaxNode> children,
        ICollection<SpreadsheetFormulaDiagnostic> diagnostics) {
        int start = cursor++;
        bool quoted = false;
        while (cursor < text.Length) {
            char character = text[cursor];
            if (character == '\'') {
                if (quoted && cursor + 1 < text.Length && text[cursor + 1] == '\'') {
                    cursor += 2;
                    continue;
                }
                quoted = !quoted;
                cursor++;
                continue;
            }
            if (character == ']' && !quoted) break;
            cursor++;
        }
        if (cursor >= text.Length) {
            children.Add(Token(SpreadsheetFormulaTokenKind.Unsupported, text.Substring(start), start));
            diagnostics.Add(Error(
                "FORMULA_UNTERMINATED_REFERENCE",
                "The OpenFormula reference has no closing bracket.",
                start,
                text.Length - start));
            return;
        }

        cursor++;
        string authored = text.Substring(start, cursor - start);
        string address = authored.Substring(1, authored.Length - 2);
        if (string.Equals(address, "#REF!", StringComparison.OrdinalIgnoreCase)) {
            children.Add(Token(SpreadsheetFormulaTokenKind.ErrorLiteral, authored, start));
            return;
        }
        if (!SpreadsheetRangeReference.TryParse(address, SpreadsheetAddressDialect.OpenDocument,
                out SpreadsheetRangeReference? reference)) {
            children.Add(Token(SpreadsheetFormulaTokenKind.Unsupported, authored, start));
            diagnostics.Add(Error(
                "FORMULA_INVALID_REFERENCE",
                $"'{authored}' is not a supported OpenFormula reference.",
                start,
                authored.Length));
            return;
        }
        children.Add(new SpreadsheetFormulaSyntaxNode(
            SpreadsheetFormulaSyntaxKind.Token,
            SpreadsheetFormulaTokenKind.Reference,
            authored,
            start,
            reference: reference));
    }

    private static void ReadWhitespace(
        string text,
        SpreadsheetFormulaDialect dialect,
        ref int cursor,
        ICollection<SpreadsheetFormulaSyntaxNode> children) {
        int start = cursor;
        while (cursor < text.Length && char.IsWhiteSpace(text[cursor])) cursor++;
        SpreadsheetFormulaTokenKind kind = SpreadsheetFormulaTokenKind.Whitespace;
        if (dialect == SpreadsheetFormulaDialect.ExcelA1 && IsReferenceLikeLast(children)) {
            int next = cursor;
            if (next < text.Length && SpreadsheetRangeReference.TryReadExcelAt(
                    text, next, out SpreadsheetRangeReference? _, out int _)) {
                kind = SpreadsheetFormulaTokenKind.IntersectionOperator;
            }
        }
        children.Add(Token(kind, text.Substring(start, cursor - start), start));
    }

    private static bool TryReadSeparator(
        string text,
        SpreadsheetFormulaDialect dialect,
        ref int cursor,
        SpreadsheetFormulaSyntaxKind containerKind,
        ICollection<SpreadsheetFormulaSyntaxNode> children,
        ICollection<SpreadsheetFormulaDiagnostic> diagnostics) {
        char character = text[cursor];
        SpreadsheetFormulaTokenKind? kind = null;
        if (dialect == SpreadsheetFormulaDialect.ExcelA1) {
            if (character == ',') {
                kind = containerKind == SpreadsheetFormulaSyntaxKind.InlineArray
                    ? SpreadsheetFormulaTokenKind.ArrayColumnSeparator
                    : containerKind == SpreadsheetFormulaSyntaxKind.FunctionCall
                        ? SpreadsheetFormulaTokenKind.ArgumentSeparator
                        : SpreadsheetFormulaTokenKind.UnionOperator;
            } else if (character == ';' && containerKind == SpreadsheetFormulaSyntaxKind.InlineArray) {
                kind = SpreadsheetFormulaTokenKind.ArrayRowSeparator;
            }
        } else {
            if (character == ';') {
                kind = containerKind == SpreadsheetFormulaSyntaxKind.InlineArray
                    ? SpreadsheetFormulaTokenKind.ArrayColumnSeparator
                    : containerKind == SpreadsheetFormulaSyntaxKind.FunctionCall
                        ? SpreadsheetFormulaTokenKind.ArgumentSeparator
                        : (SpreadsheetFormulaTokenKind?)null;
            } else if (character == '|' && containerKind == SpreadsheetFormulaSyntaxKind.InlineArray) {
                kind = SpreadsheetFormulaTokenKind.ArrayRowSeparator;
            } else if (character == '~') {
                kind = SpreadsheetFormulaTokenKind.UnionOperator;
            }
        }
        if (!kind.HasValue) {
            if (character != ',' && character != ';' && character != '|') return false;
            diagnostics.Add(Error(
                "FORMULA_SEPARATOR_CONTEXT",
                $"Separator '{character}' is not valid in this {dialect} formula context.",
                cursor,
                1));
            kind = SpreadsheetFormulaTokenKind.Unsupported;
        }
        children.Add(Token(kind.Value, character.ToString(), cursor));
        cursor++;
        return true;
    }

    private static bool TryReadNumber(
        string text,
        ref int cursor,
        ICollection<SpreadsheetFormulaSyntaxNode> children) {
        int start = cursor;
        if (!(text[cursor] >= '0' && text[cursor] <= '9') &&
            !(text[cursor] == '.' && cursor + 1 < text.Length && text[cursor + 1] >= '0' && text[cursor + 1] <= '9')) {
            return false;
        }
        while (cursor < text.Length && text[cursor] >= '0' && text[cursor] <= '9') cursor++;
        if (cursor < text.Length && text[cursor] == '.') {
            cursor++;
            while (cursor < text.Length && text[cursor] >= '0' && text[cursor] <= '9') cursor++;
        }
        if (cursor < text.Length && (text[cursor] == 'e' || text[cursor] == 'E')) {
            int exponent = cursor++;
            if (cursor < text.Length && (text[cursor] == '+' || text[cursor] == '-')) cursor++;
            int digits = cursor;
            while (cursor < text.Length && text[cursor] >= '0' && text[cursor] <= '9') cursor++;
            if (digits == cursor) cursor = exponent;
        }
        children.Add(Token(SpreadsheetFormulaTokenKind.NumberLiteral, text.Substring(start, cursor - start), start));
        return true;
    }

    private static bool TryReadError(
        string text,
        SpreadsheetFormulaDialect dialect,
        ref int cursor,
        ICollection<SpreadsheetFormulaSyntaxNode> children,
        ICollection<SpreadsheetFormulaDiagnostic> diagnostics) {
        if (text[cursor] != '#') return false;
        if (cursor + 1 >= text.Length || !IsIdentifierStart(text[cursor + 1])) return false;
        int start = cursor++;
        while (cursor < text.Length) {
            char character = text[cursor];
            if (char.IsWhiteSpace(character) || character == ',' || character == ';' || character == ')' || character == '}') break;
            cursor++;
            if (character == '!' || character == '?') break;
        }
        if (dialect == SpreadsheetFormulaDialect.ExcelA1
            && string.Equals(text.Substring(start, cursor - start), "#REF!", StringComparison.OrdinalIgnoreCase)
            && cursor < text.Length
            && SpreadsheetRangeReference.TryReadExcelAt(text, cursor, out _, out int consumed)) {
            cursor += consumed;
            children.Add(Token(SpreadsheetFormulaTokenKind.Unsupported, text.Substring(start, cursor - start), start));
            diagnostics.Add(Error(
                "FORMULA_DELETED_REFERENCE",
                "Excel deleted references with an attached address cannot be represented safely in OpenFormula.",
                start,
                cursor - start));
            return true;
        }
        children.Add(Token(SpreadsheetFormulaTokenKind.ErrorLiteral, text.Substring(start, cursor - start), start));
        return true;
    }

    private static bool TryReadIdentifier(
        string text,
        ref int cursor,
        ICollection<SpreadsheetFormulaSyntaxNode> children) {
        if (!IsIdentifierStart(text[cursor])) return false;
        int start = cursor;
        cursor = ScanIdentifier(text, cursor);
        children.Add(Token(SpreadsheetFormulaTokenKind.Identifier, text.Substring(start, cursor - start), start));
        return true;
    }

    private static int ScanIdentifier(string text, int cursor) {
        cursor++;
        while (cursor < text.Length) {
            char character = text[cursor];
            if (!IsIdentifierStart(character) && !(character >= '0' && character <= '9') && character != '.') break;
            cursor++;
        }
        return cursor;
    }

    private static bool TryReadOperator(
        string text,
        SpreadsheetFormulaDialect dialect,
        ref int cursor,
        ICollection<SpreadsheetFormulaSyntaxNode> children,
        ICollection<SpreadsheetFormulaDiagnostic> diagnostics) {
        int start = cursor;
        char character = text[cursor];
        if (dialect == SpreadsheetFormulaDialect.OpenFormula && character == '!') {
            children.Add(Token(SpreadsheetFormulaTokenKind.IntersectionOperator, "!", cursor++));
            return true;
        }
        if (character == '@' || (character == '#' && start > 0)) {
            children.Add(Token(SpreadsheetFormulaTokenKind.Unsupported, character.ToString(), cursor++));
            return true;
        }
        if ("+-*/^&%=<>:".IndexOf(character) < 0) return false;
        cursor++;
        if (cursor < text.Length &&
            ((character == '<' && (text[cursor] == '=' || text[cursor] == '>')) ||
             (character == '>' && text[cursor] == '='))) {
            cursor++;
        }
        SpreadsheetFormulaTokenKind kind = character == ':'
            ? SpreadsheetFormulaTokenKind.Unsupported
            : SpreadsheetFormulaTokenKind.Operator;
        children.Add(Token(kind, text.Substring(start, cursor - start), start));
        if (character == ':') {
            diagnostics.Add(Error(
                "FORMULA_UNSUPPORTED_RANGE_OPERATOR",
                "A range or 3-D reference was not represented by typed address syntax.",
                start,
                cursor - start));
        }
        return true;
    }

    private static bool IsReferenceLikeLast(ICollection<SpreadsheetFormulaSyntaxNode> children) {
        SpreadsheetFormulaSyntaxNode? last = null;
        foreach (SpreadsheetFormulaSyntaxNode child in children) last = child;
        if (last == null) return false;
        return last.TokenKind == SpreadsheetFormulaTokenKind.Reference ||
               last.Kind == SpreadsheetFormulaSyntaxKind.ParenthesizedExpression;
    }

    private static bool IsIdentifierStart(char character) =>
        (character >= 'A' && character <= 'Z') ||
        (character >= 'a' && character <= 'z') ||
        character == '_';

    private static SpreadsheetFormulaSyntaxNode Token(
        SpreadsheetFormulaTokenKind kind,
        string text,
        int position) => new SpreadsheetFormulaSyntaxNode(
            SpreadsheetFormulaSyntaxKind.Token,
            kind,
            text,
            position);

    private static SpreadsheetFormulaDiagnostic Error(string code, string message, int position, int length) =>
        new SpreadsheetFormulaDiagnostic(code, SpreadsheetFormulaDiagnosticSeverity.Error, message, position, length);
}
