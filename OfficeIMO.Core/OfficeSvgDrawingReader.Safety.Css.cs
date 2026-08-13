using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static bool ContainsLocalCssUrlReference(string? value, ISet<string> relevantIds) {
        if (string.IsNullOrWhiteSpace(value) || relevantIds.Count == 0) return false;
        string normalized = StripCssComments(value!);
        char outerQuote = '\0';
        for (int start = 0; start < normalized.Length;) {
            char current = normalized[start];
            if (outerQuote != '\0') {
                if (current == '\\') {
                    int escapedIndex = start;
                    if (TryReadCssCharacter(normalized, ref escapedIndex, out _)) {
                        start = escapedIndex;
                        continue;
                    }
                }
                if (current == outerQuote) outerQuote = '\0';
                start++;
                continue;
            }
            if (current is '\'' or '"') {
                outerQuote = current;
                start++;
                continue;
            }

            int index = start;
            if (!TryReadCssIdentifier(normalized, ref index, "url") || index >= normalized.Length || normalized[index] != '(') {
                start++;
                continue;
            }
            index++;
            while (index < normalized.Length && char.IsWhiteSpace(normalized[index])) index++;
            var target = new StringBuilder();
            char quote = '\0';
            if (index < normalized.Length && normalized[index] is '\'' or '"') quote = normalized[index++];
            bool closed = false;
            while (index < normalized.Length) {
                current = normalized[index];
                if (quote == '\0' && current == ')') {
                    index++;
                    closed = true;
                    break;
                }
                if (quote != '\0' && current == quote) {
                    index++;
                    quote = '\0';
                    while (index < normalized.Length && char.IsWhiteSpace(normalized[index])) index++;
                    if (index < normalized.Length && normalized[index] == ')') {
                        index++;
                        closed = true;
                    }
                    break;
                }
                if (!TryReadCssCharacter(normalized, ref index, out char decoded)) break;
                target.Append(decoded);
            }
            start = Math.Max(index, start + 1);
            if (!closed || quote != '\0') continue;
            string reference = target.ToString().Trim();
            if (reference.Length > 1 && reference[0] == '#' && relevantIds.Contains(reference.Substring(1))) return true;
        }
        return false;
    }

    // Keep safety inspection aligned with ChartForgeX's stylesheet parser, which removes
    // comments before parsing selectors and declarations (including comments between tokens).
    private static string StripCssComments(string value) {
        var result = new StringBuilder(value.Length);
        for (int index = 0; index < value.Length; index++) {
            if (index + 1 < value.Length && value[index] == '/' && value[index + 1] == '*') {
                index += 2;
                while (index + 1 < value.Length && !(value[index] == '*' && value[index + 1] == '/')) index++;
                if (index + 1 < value.Length) index++;
                continue;
            }
            result.Append(value[index]);
        }
        return result.ToString();
    }

    private static bool ContainsLocalCssCustomPropertyUrlReference(string? value, ISet<string> relevantIds) {
        if (string.IsNullOrWhiteSpace(value) || relevantIds.Count == 0) return false;
        string normalized = StripCssComments(value!);
        foreach (string declaration in SplitRasterStyleDeclarations(normalized)) {
            int colon = declaration.IndexOf(':');
            if (colon <= 0) continue;
            string name = declaration.Substring(0, colon).Trim();
            if (name.StartsWith("--", StringComparison.Ordinal)
                && ContainsLocalCssUrlReference(declaration.Substring(colon + 1), relevantIds)) return true;
        }
        return false;
    }

    private static bool ContainsPotentialCssIdentifier(string? value, string identifier) {
        if (string.IsNullOrWhiteSpace(value)) return false;
        for (int start = 0; start < value!.Length; start++) {
            int index = start;
            if (TryReadCssIdentifier(value, ref index, identifier)
                && index < value.Length
                && value[index] == '(') return true;
        }
        return false;
    }

    private static bool TryReadCssIdentifier(string value, ref int index, string identifier) {
        for (int expectedIndex = 0; expectedIndex < identifier.Length; expectedIndex++) {
            if (!TryReadCssCharacter(value, ref index, out char actual)
                || char.ToLowerInvariant(actual) != char.ToLowerInvariant(identifier[expectedIndex])) return false;
        }
        return true;
    }

    private static bool TryDecodeCssIdentifier(string value, out string decoded) {
        var result = new StringBuilder(value.Length);
        int index = 0;
        while (index < value.Length) {
            if (!TryReadCssCharacter(value, ref index, out char character)) {
                decoded = string.Empty;
                return false;
            }
            result.Append(character);
        }
        decoded = result.ToString();
        return decoded.Length > 0;
    }

    private static bool TryReadCssCharacter(string value, ref int index, out char character) {
        character = default;
        if (index >= value.Length) return false;
        char current = value[index++];
        if (current != '\\') {
            character = current;
            return true;
        }
        if (index >= value.Length || value[index] is '\r' or '\n' or '\f') return false;
        if (!TryGetCssHexValue(value[index], out int digit)) {
            character = value[index++];
            return true;
        }

        int scalar = 0;
        int digits = 0;
        do {
            scalar = scalar * 16 + digit;
            index++;
            digits++;
        } while (digits < 6 && index < value.Length && TryGetCssHexValue(value[index], out digit));
        if (index < value.Length && char.IsWhiteSpace(value[index])) {
            if (value[index++] == '\r' && index < value.Length && value[index] == '\n') index++;
        }
        character = scalar is > 0 and <= char.MaxValue ? (char) scalar : '\uFFFD';
        return true;
    }

    private static bool TryGetCssHexValue(char value, out int digit) {
        if (value is >= '0' and <= '9') {
            digit = value - '0';
            return true;
        }
        char lower = char.ToLowerInvariant(value);
        if (lower is >= 'a' and <= 'f') {
            digit = lower - 'a' + 10;
            return true;
        }
        digit = 0;
        return false;
    }
}
