using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Spreadsheet;

internal static partial class SpreadsheetNumberFormatDisplay {
    /// <summary>Formats an already resolved calendar value without imposing a workbook date system.</summary>
    internal static string? FormatDateTimeValue(DateTime date, string formatCode, double sectionValue = 0) {
        string section = SelectNumberFormatSection(formatCode, sectionValue < 0 ? 1 : sectionValue == 0 ? 2 : 0,
            sectionValue, out _);
        var syntax = SpreadsheetNumberFormatSyntax.Parse(section);
        if (!syntax.IsValid) return null;
        if (!TryRoundTemporalTicks(date.Ticks, syntax.Tokens, out long roundedDateTicks)) return null;
        try {
            date = new DateTime(roundedDateTicks, date.Kind);
        } catch (ArgumentOutOfRangeException) {
            return null;
        }
        bool twelveHour = syntax.Tokens.Select((_, index) => index).Any(index =>
            TryReadMeridiem(syntax, index, out _, out _));
        var builder = new StringBuilder(section.Length);
        for (int index = 0; index < syntax.Tokens.Count; index++) {
            var token = syntax.Tokens[index];
            if (TryReadMeridiem(syntax, index, out string meridiem, out int lastIndex)) {
                string[] halves = meridiem.Split('/');
                builder.Append(halves[date.Hour < 12 ? 0 : 1]);
                index = lastIndex;
            } else if (token.Kind == SpreadsheetNumberFormatTokenKind.DateTimeSymbol) {
                char unit = char.ToLowerInvariant(token.Text[0]);
                int length = token.Text.Length;
                string pattern;
                switch (unit) {
                    case 'y': pattern = length <= 2 ? "yy" : "yyyy"; break;
                    case 'd': pattern = length >= 4 ? "dddd" : length == 3 ? "ddd" : length == 2 ? "dd" : "%d"; break;
                    case 'm':
                        if (length == 5) {
                            builder.Append(date.ToString("MMMM", CultureInfo.InvariantCulture)[0]);
                            continue;
                        }
                        pattern = length >= 4 ? "MMMM" : length == 3 ? "MMM"
                            : IsTemporalMinute(syntax.Tokens, index) ? length == 1 ? "%m" : "mm"
                            : length == 1 ? "%M" : "MM";
                        break;
                    case 'h': pattern = twelveHour ? length == 1 ? "%h" : "hh" : length == 1 ? "%H" : "HH"; break;
                    case 's': pattern = length == 1 ? "%s" : "ss"; break;
                    default: return null;
                }
                builder.Append(date.ToString(pattern, CultureInfo.InvariantCulture));
            } else if (IsFractionalSecond(syntax.Tokens, index)) {
                if (token.Text.Length > 7) return null;
                builder.Append(date.ToString(token.Text.Length == 1 ? "%f" : new string('f', token.Text.Length), CultureInfo.InvariantCulture));
            } else if (!AppendTemporalLiteral(builder, token)) {
                return null;
            }
        }
        return builder.ToString();
    }

    /// <summary>Recognizes actual elapsed directives, excluding quoted or escaped bracket text.</summary>
    internal static bool HasElapsedTimeToken(string formatCode) {
        var syntax = SpreadsheetNumberFormatSyntax.Parse(formatCode);
        return syntax.IsValid && syntax.Tokens.Any(token => ElapsedUnit(token) != '\0');
    }

    /// <summary>Recognizes fractional-second placeholders without treating literals as precision.</summary>
    internal static bool HasFractionalSecondToken(string formatCode) {
        var syntax = SpreadsheetNumberFormatSyntax.Parse(formatCode);
        return syntax.IsValid && syntax.Tokens.Select((_, index) => index)
            .Any(index => IsFractionalSecond(syntax.Tokens, index));
    }

    /// <summary>Formats Excel-style elapsed time from a serial measured in days.</summary>
    internal static bool TryFormatElapsedValue(double value, string formatCode, out string text) {
        text = string.Empty;
        string section = SelectNumberFormatSection(formatCode, value < 0 ? 1 : value == 0 ? 2 : 0, value, out int selectedSection);
        var syntax = SpreadsheetNumberFormatSyntax.Parse(section);
        if (!syntax.IsValid || !syntax.Tokens.Any(token => ElapsedUnit(token) != '\0')) return false;
        TimeSpan duration;
        try {
            duration = TimeSpan.FromDays(value).Duration();
        } catch (ArgumentException) {
            return false;
        } catch (OverflowException) {
            return false;
        }
        if (!TryRoundTemporalTicks(duration.Ticks, syntax.Tokens, out long roundedDurationTicks)) return false;
        duration = TimeSpan.FromTicks(roundedDurationTicks);
        var builder = new StringBuilder(section.Length);
        if (value < 0 && selectedSection != 1) builder.Append('-');
        for (int index = 0; index < syntax.Tokens.Count; index++) {
            var token = syntax.Tokens[index];
            char elapsedUnit = ElapsedUnit(token);
            if (elapsedUnit != '\0') {
                long ticksPerUnit = elapsedUnit == 'h' ? TimeSpan.TicksPerHour
                    : elapsedUnit == 'm' ? TimeSpan.TicksPerMinute : TimeSpan.TicksPerSecond;
                builder.Append((duration.Ticks / ticksPerUnit).ToString(
                    "D" + token.Value.Length.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture));
            } else if (token.Kind == SpreadsheetNumberFormatTokenKind.DateTimeSymbol) {
                char unit = char.ToLowerInvariant(token.Text[0]);
                int component;
                if (unit == 'h' && token.Text.Length <= 2) component = duration.Hours;
                else if (unit == 'm' && token.Text.Length <= 2 && IsTemporalMinute(syntax.Tokens, index)) component = duration.Minutes;
                else if (unit == 's' && token.Text.Length <= 2) component = duration.Seconds;
                else return false;
                builder.Append(component.ToString("D" + token.Text.Length.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture));
            } else if (IsFractionalSecond(syntax.Tokens, index)) {
                if (token.Text.Length > 7) return false;
                builder.Append((duration.Ticks % TimeSpan.TicksPerSecond)
                    .ToString("D7", CultureInfo.InvariantCulture).Substring(0, token.Text.Length));
            } else if (!AppendTemporalLiteral(builder, token)) {
                return false;
            }
        }
        text = builder.ToString();
        return true;
    }

    private static char ElapsedUnit(SpreadsheetNumberFormatToken token) {
        if (token.Kind != SpreadsheetNumberFormatTokenKind.BracketedDirective
            || token.Value.Length is < 1 or > 2) return '\0';
        char unit = char.ToLowerInvariant(token.Value[0]);
        return unit is 'h' or 'm' or 's' && token.Value.All(ch => char.ToLowerInvariant(ch) == unit) ? unit : '\0';
    }

    private static bool IsTemporalMinute(IReadOnlyList<SpreadsheetNumberFormatToken> tokens, int index) =>
        AdjacentTemporalUnit(tokens, index, -1) == 'h' || AdjacentTemporalUnit(tokens, index, 1) == 's';

    private static char AdjacentTemporalUnit(IReadOnlyList<SpreadsheetNumberFormatToken> tokens, int index, int direction) {
        for (int adjacent = index + direction; adjacent >= 0 && adjacent < tokens.Count; adjacent += direction) {
            var token = tokens[adjacent];
            if (token.Kind == SpreadsheetNumberFormatTokenKind.DateTimeSymbol) return char.ToLowerInvariant(token.Text[0]);
            char elapsed = ElapsedUnit(token);
            if (elapsed != '\0') return elapsed;
            if (token.Kind is SpreadsheetNumberFormatTokenKind.Placeholder or SpreadsheetNumberFormatTokenKind.TextPlaceholder) return '\0';
        }
        return '\0';
    }

    private static bool IsFractionalSecond(IReadOnlyList<SpreadsheetNumberFormatToken> tokens, int index) =>
        index >= 2 && tokens[index].Kind == SpreadsheetNumberFormatTokenKind.Placeholder
        && tokens[index].Text.All(ch => ch == '0')
        && tokens[index - 1].Kind == SpreadsheetNumberFormatTokenKind.DecimalSeparator
        && (ElapsedUnit(tokens[index - 2]) == 's' || tokens[index - 2].Kind == SpreadsheetNumberFormatTokenKind.DateTimeSymbol
            && char.ToLowerInvariant(tokens[index - 2].Text[0]) == 's');

    // Rounding before splitting components carries a displayed fractional second into
    // the seconds, minutes, hours and calendar date instead of printing an inconsistent tail.
    private static bool TryRoundTemporalTicks(long ticks, IReadOnlyList<SpreadsheetNumberFormatToken> tokens, out long rounded) {
        rounded = ticks;
        int precision = 7;
        for (int index = 0; index < tokens.Count; index++) {
            if (!IsFractionalSecond(tokens, index)) continue;
            if (tokens[index].Text.Length > 7) return false;
            precision = Math.Min(precision, tokens[index].Text.Length);
        }
        long quantum = 1;
        for (int index = precision; index < 7; index++) quantum *= 10;
        long remainder = ticks % quantum;
        long adjustment = remainder * 2 >= quantum ? quantum - remainder : -remainder;
        if (adjustment > 0 && ticks > long.MaxValue - adjustment) return false;
        rounded = ticks + adjustment;
        return true;
    }

    private static bool AppendTemporalLiteral(StringBuilder builder, SpreadsheetNumberFormatToken token) {
        switch (token.Kind) {
            case SpreadsheetNumberFormatTokenKind.BracketedDirective: return ElapsedUnit(token) == '\0';
            case SpreadsheetNumberFormatTokenKind.Currency: builder.Append(token.Value); return true;
            case SpreadsheetNumberFormatTokenKind.Literal:
                if (token.Text[0] == '_') builder.Append(' ');
                else if (token.Text[0] != '*') builder.Append(token.Value);
                return true;
            case SpreadsheetNumberFormatTokenKind.DecimalSeparator:
            case SpreadsheetNumberFormatTokenKind.GroupSeparator:
            case SpreadsheetNumberFormatTokenKind.ScalingSeparator:
            case SpreadsheetNumberFormatTokenKind.Other:
                builder.Append(token.Text);
                return true;
            default: return false;
        }
    }

    private static bool TryReadMeridiem(SpreadsheetNumberFormatSyntax syntax, int index, out string marker, out int lastIndex) {
        marker = string.Empty;
        lastIndex = index;
        var start = syntax.Tokens[index];
        if (start.Kind != SpreadsheetNumberFormatTokenKind.Other) return false;
        foreach (string candidate in new[] { "AM/PM", "A/P" }) {
            if (start.Position + candidate.Length > syntax.Text.Length
                || string.Compare(syntax.Text, start.Position, candidate, 0, candidate.Length, StringComparison.OrdinalIgnoreCase) != 0) continue;
            int end = start.Position + candidate.Length;
            for (int next = index; next < syntax.Tokens.Count; next++) {
                var token = syntax.Tokens[next];
                if (token.Kind is not SpreadsheetNumberFormatTokenKind.Other and not SpreadsheetNumberFormatTokenKind.DateTimeSymbol
                    || token.Position + token.Text.Length > end) break;
                if (token.Position + token.Text.Length == end) {
                    marker = syntax.Text.Substring(start.Position, candidate.Length);
                    lastIndex = next;
                    return true;
                }
            }
        }
        return false;
    }
}
