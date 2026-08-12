using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace OfficeIMO.Spreadsheet;

/// <summary>Lexical role of one losslessly parsed spreadsheet number-format token.</summary>
public enum SpreadsheetNumberFormatTokenKind {
    /// <summary>A digit placeholder such as <c>0</c>, <c>#</c>, or <c>?</c>.</summary>
    Placeholder,
    /// <summary>A decimal point.</summary>
    DecimalSeparator,
    /// <summary>A thousands-grouping comma.</summary>
    GroupSeparator,
    /// <summary>A trailing comma that scales the displayed value by one thousand.</summary>
    ScalingSeparator,
    /// <summary>An unquoted percent operator that scales the displayed value.</summary>
    Percent,
    /// <summary>A currency symbol or bracketed currency directive.</summary>
    Currency,
    /// <summary>A date or time component run.</summary>
    DateTimeSymbol,
    /// <summary>Literal text, including quoted, escaped, spacing, and fill text.</summary>
    Literal,
    /// <summary>A bracketed condition, color, locale, elapsed-time, or other directive.</summary>
    BracketedDirective,
    /// <summary>A semicolon separating positive, negative, zero, or text sections.</summary>
    SectionSeparator,
    /// <summary>The text-value placeholder <c>@</c>.</summary>
    TextPlaceholder,
    /// <summary>Other authored syntax.</summary>
    Other
}

/// <summary>One lossless token in a spreadsheet number format.</summary>
public sealed class SpreadsheetNumberFormatToken {
    internal SpreadsheetNumberFormatToken(
        SpreadsheetNumberFormatTokenKind kind,
        string text,
        string value,
        int position,
        string? localeCode = null) {
        Kind = kind;
        Text = text;
        Value = value;
        Position = position;
        LocaleCode = localeCode;
    }

    /// <summary>Token role.</summary>
    public SpreadsheetNumberFormatTokenKind Kind { get; }
    /// <summary>Exact authored text.</summary>
    public string Text { get; }
    /// <summary>Decoded semantic value for literal and currency tokens.</summary>
    public string Value { get; }
    /// <summary>Zero-based source position.</summary>
    public int Position { get; }
    /// <summary>Authored locale code from a bracketed currency or locale directive, without the leading hyphen.</summary>
    public string? LocaleCode { get; }
}

/// <summary>
/// Lossless, non-regex syntax model for Excel-style number formats. It distinguishes display literals
/// from scaling operators and exposes the conservative common subset used by format adapters.
/// </summary>
public sealed class SpreadsheetNumberFormatSyntax {
    private static readonly char[] CurrencyCharacters = { '$', '€', '£', '¥', '₹', '₽', '₩', '₺' };

    private SpreadsheetNumberFormatSyntax(string text, IReadOnlyList<SpreadsheetNumberFormatToken> tokens, bool isValid) {
        Text = text;
        Tokens = tokens;
        IsValid = isValid;
    }

    /// <summary>Exact authored format.</summary>
    public string Text { get; }
    /// <summary>Ordered lossless tokens.</summary>
    public IReadOnlyList<SpreadsheetNumberFormatToken> Tokens { get; }
    /// <summary>Whether all quoted and bracketed regions were terminated.</summary>
    public bool IsValid { get; }
    /// <summary>Number of authored format sections.</summary>
    public int SectionCount => 1 + Tokens.Count(token => token.Kind == SpreadsheetNumberFormatTokenKind.SectionSeparator);
    /// <summary>Whether the first section contains an unquoted percentage-scaling operator.</summary>
    public bool IsPercentage => FirstSection().Any(token => token.Kind == SpreadsheetNumberFormatTokenKind.Percent);
    /// <summary>First currency symbol or code in the first section.</summary>
    public string? CurrencySymbol => FirstSection()
        .FirstOrDefault(token => token.Kind == SpreadsheetNumberFormatTokenKind.Currency)?.Value;
    /// <summary>Whether the first section requests thousands grouping.</summary>
    public bool UsesGrouping => FirstSection().Any(token => token.Kind == SpreadsheetNumberFormatTokenKind.GroupSeparator);
    /// <summary>Number of first-section commas that scale the displayed value by one thousand.</summary>
    public int ScaleThousands => FirstSection().Count(token => token.Kind == SpreadsheetNumberFormatTokenKind.ScalingSeparator);
    /// <summary>Digit placeholders following the first decimal separator in the first section.</summary>
    public int DecimalPlaces {
        get {
            bool afterDecimal = false;
            int count = 0;
            foreach (SpreadsheetNumberFormatToken token in FirstSection()) {
                if (!afterDecimal) {
                    if (token.Kind == SpreadsheetNumberFormatTokenKind.DecimalSeparator) afterDecimal = true;
                    continue;
                }
                if (token.Kind == SpreadsheetNumberFormatTokenKind.Placeholder) count += token.Text.Length;
                else if (token.Kind != SpreadsheetNumberFormatTokenKind.Literal) break;
            }
            return Math.Min(30, count);
        }
    }

    /// <summary>Parses a number format while retaining every authored character.</summary>
    public static SpreadsheetNumberFormatSyntax Parse(string format) {
        if (format == null) throw new ArgumentNullException(nameof(format));
        var tokens = new List<SpreadsheetNumberFormatToken>();
        bool valid = true;
        int index = 0;
        while (index < format.Length) {
            int start = index;
            char current = format[index];
            if (current == '"') {
                index++;
                var decoded = new System.Text.StringBuilder();
                bool terminated = false;
                while (index < format.Length) {
                    if (format[index] != '"') {
                        decoded.Append(format[index++]);
                        continue;
                    }
                    if (index + 1 < format.Length && format[index + 1] == '"') {
                        decoded.Append('"');
                        index += 2;
                        continue;
                    }
                    index++;
                    terminated = true;
                    break;
                }
                valid &= terminated;
                Add(tokens, SpreadsheetNumberFormatTokenKind.Literal, format, start, index, decoded.ToString());
            } else if (current == '[') {
                index++;
                while (index < format.Length && format[index] != ']') index++;
                bool terminated = index < format.Length;
                if (terminated) index++;
                valid &= terminated;
                string value = format.Substring(start + 1, Math.Max(0, index - start - (terminated ? 2 : 1)));
                if (value.StartsWith("$", StringComparison.Ordinal)) {
                    string currency = value.Substring(1);
                    string? localeCode = null;
                    int locale = currency.LastIndexOf('-');
                    if (locale >= 0) {
                        localeCode = currency.Substring(locale + 1);
                        currency = currency.Substring(0, locale);
                    }
                    Add(tokens, SpreadsheetNumberFormatTokenKind.Currency, format, start, index, currency, localeCode);
                } else {
                    Add(tokens, SpreadsheetNumberFormatTokenKind.BracketedDirective, format, start, index, value);
                }
            } else if (current == '\\' || current == '_' || current == '*') {
                index = Math.Min(format.Length, index + 2);
                string value = start + 1 < format.Length ? format[start + 1].ToString() : string.Empty;
                Add(tokens, SpreadsheetNumberFormatTokenKind.Literal, format, start, index, value);
            } else if (current == ';') {
                index++;
                Add(tokens, SpreadsheetNumberFormatTokenKind.SectionSeparator, format, start, index, ";");
            } else if (current == '%') {
                index++;
                Add(tokens, SpreadsheetNumberFormatTokenKind.Percent, format, start, index, "%");
            } else if (Array.IndexOf(CurrencyCharacters, current) >= 0) {
                index++;
                Add(tokens, SpreadsheetNumberFormatTokenKind.Currency, format, start, index, current.ToString());
            } else if (current == '0' || current == '#' || current == '?') {
                index++;
                while (index < format.Length && format[index] == current) index++;
                Add(tokens, SpreadsheetNumberFormatTokenKind.Placeholder, format, start, index, format.Substring(start, index - start));
            } else if (current == '.') {
                index++;
                Add(tokens, SpreadsheetNumberFormatTokenKind.DecimalSeparator, format, start, index, ".");
            } else if (current == ',') {
                index++;
                Add(tokens, SpreadsheetNumberFormatTokenKind.GroupSeparator, format, start, index, ",");
            } else if (current == '@') {
                index++;
                Add(tokens, SpreadsheetNumberFormatTokenKind.TextPlaceholder, format, start, index, "@");
            } else if (IsDateTimeLetter(current)) {
                index++;
                while (index < format.Length && char.ToLowerInvariant(format[index]) == char.ToLowerInvariant(current)) index++;
                Add(tokens, SpreadsheetNumberFormatTokenKind.DateTimeSymbol, format, start, index, format.Substring(start, index - start));
            } else {
                index++;
                Add(tokens, SpreadsheetNumberFormatTokenKind.Other, format, start, index, current.ToString());
            }
        }
        ClassifyCommas(tokens);
        return new SpreadsheetNumberFormatSyntax(format, new ReadOnlyCollection<SpreadsheetNumberFormatToken>(tokens), valid);
    }

    private static void ClassifyCommas(List<SpreadsheetNumberFormatToken> tokens) {
        bool followingIntegerPlaceholder = false;
        for (int index = tokens.Count - 1; index >= 0; index--) {
            SpreadsheetNumberFormatToken token = tokens[index];
            if (token.Kind == SpreadsheetNumberFormatTokenKind.SectionSeparator
                || token.Kind == SpreadsheetNumberFormatTokenKind.DecimalSeparator) {
                followingIntegerPlaceholder = false;
                continue;
            }
            if (token.Kind == SpreadsheetNumberFormatTokenKind.Placeholder) {
                followingIntegerPlaceholder = true;
                continue;
            }
            if (token.Kind == SpreadsheetNumberFormatTokenKind.GroupSeparator
                && !followingIntegerPlaceholder) {
                tokens[index] = new SpreadsheetNumberFormatToken(
                    SpreadsheetNumberFormatTokenKind.ScalingSeparator,
                    token.Text,
                    token.Value,
                    token.Position);
            }
        }
    }

    private IEnumerable<SpreadsheetNumberFormatToken> FirstSection() =>
        Tokens.TakeWhile(token => token.Kind != SpreadsheetNumberFormatTokenKind.SectionSeparator);

    private static bool IsDateTimeLetter(char value) {
        char lower = char.ToLowerInvariant(value);
        return lower == 'y' || lower == 'm' || lower == 'd' || lower == 'h' || lower == 's';
    }

    private static void Add(List<SpreadsheetNumberFormatToken> tokens, SpreadsheetNumberFormatTokenKind kind,
        string source, int start, int end, string value, string? localeCode = null) =>
        tokens.Add(new SpreadsheetNumberFormatToken(
            kind,
            source.Substring(start, end - start),
            value,
            start,
            localeCode));
}
