using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html.Css;

// An+B is a token grammar. Comments may separate tokens without becoming whitespace;
// concatenating source strings would turn separate numbers or identifiers into a formula.
internal static class HtmlCssAnPlusBParser {
    internal static bool TryParse(string source, IReadOnlyList<HtmlCssToken> input, int start, int end, out HtmlCssAnPlusB formula) {
        formula = default;
        var tokens = new List<HtmlCssToken>();
        for (int index = start; index < end; index++)
            if (input[index].Kind != HtmlCssTokenKind.Comment) tokens.Add(input[index]);
        int position = 0;
        SkipWhitespace(tokens, ref position);
        if (position == tokens.Count) return false;
        bool leadingPlus = tokens[position].Kind == HtmlCssTokenKind.Delimiter && tokens[position].Value == "+";
        if (leadingPlus && (++position == tokens.Count || tokens[position].Kind != HtmlCssTokenKind.Identifier)) return false;
        HtmlCssToken first = tokens[position++];
        if (!leadingPlus && first.Kind == HtmlCssTokenKind.Number) {
            if (!Integer(first.GetText(source), out int constant) || !AtEnd(tokens, position)) return false;
            formula = new HtmlCssAnPlusB(0, constant); return true;
        }
        string unit;
        int coefficient;
        if (first.Kind == HtmlCssTokenKind.Identifier) {
            unit = first.Value ?? string.Empty;
            if (!leadingPlus && (unit.Equals("odd", StringComparison.OrdinalIgnoreCase) || unit.Equals("even", StringComparison.OrdinalIgnoreCase))) {
                if (!AtEnd(tokens, position)) return false;
                formula = new HtmlCssAnPlusB(2, unit.Equals("odd", StringComparison.OrdinalIgnoreCase) ? 1 : 0); return true;
            }
            coefficient = unit.StartsWith("-", StringComparison.Ordinal) ? -1 : 1;
            if (coefficient < 0) {
                if (leadingPlus) return false;
                unit = unit.Substring(1);
            }
        } else if (!leadingPlus && first.Kind == HtmlCssTokenKind.Dimension) {
            string number = Regex.Match(first.GetText(source), @"^[+-]?(?:\d*\.\d+|\d+\.?\d*)(?:[eE][+-]?\d+)?").Value;
            if (!Integer(number, out coefficient)) return false;
            unit = first.Value ?? string.Empty;
        } else return false;
        Match n = Regex.Match(unit, @"\An(?:-(\d*))?\z", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        if (!n.Success) return false;
        int offset = 0;
        if (n.Groups[1].Success) {
            string digits = n.Groups[1].Value;
            if (digits.Length > 0) {
                if (!Integer("-" + digits, out offset) || !AtEnd(tokens, position)) return false;
            } else {
                SkipWhitespace(tokens, ref position);
                if (position == tokens.Count || tokens[position].Kind != HtmlCssTokenKind.Number) return false;
                digits = tokens[position++].GetText(source);
                if (!UnsignedInteger(digits) || !Integer("-" + digits, out offset) || !AtEnd(tokens, position)) return false;
            }
        } else {
            SkipWhitespace(tokens, ref position);
            if (position < tokens.Count) {
                HtmlCssToken next = tokens[position++];
                if (next.Kind == HtmlCssTokenKind.Number) {
                    string number = next.GetText(source);
                    if (number.Length == 0 || (number[0] != '+' && number[0] != '-') || !Integer(number, out offset)) return false;
                } else if (next.Kind == HtmlCssTokenKind.Delimiter && (next.Value == "+" || next.Value == "-")) {
                    SkipWhitespace(tokens, ref position);
                    if (position == tokens.Count || tokens[position].Kind != HtmlCssTokenKind.Number) return false;
                    string number = tokens[position++].GetText(source);
                    if (!UnsignedInteger(number) || !Integer(next.Value + number, out offset)) return false;
                } else return false;
                if (!AtEnd(tokens, position)) return false;
            }
        }
        formula = new HtmlCssAnPlusB(coefficient, offset); return true;
    }

    private static bool Integer(string value, out int number) {
        number = 0;
        return Regex.IsMatch(value, @"\A[+-]?\d+\z", RegexOptions.CultureInvariant)
            && int.TryParse(value, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out number);
    }
    private static bool UnsignedInteger(string value) => Regex.IsMatch(value, @"\A\d+\z", RegexOptions.CultureInvariant);
    private static bool AtEnd(List<HtmlCssToken> tokens, int position) {
        SkipWhitespace(tokens, ref position); return position == tokens.Count;
    }
    private static void SkipWhitespace(List<HtmlCssToken> tokens, ref int position) {
        while (position < tokens.Count && tokens[position].Kind == HtmlCssTokenKind.Whitespace) position++;
    }
}
