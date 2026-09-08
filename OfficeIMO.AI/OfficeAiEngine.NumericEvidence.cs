using System.Globalization;

namespace OfficeIMO.AI;

public sealed partial class OfficeAiEngine {
    // Substring provenance alone can drop a sign, leading digits, or a fractional part.
    // Check the original observation, including context outside a quote or request slice.
    private static bool SupportsCompleteNumber(string raw, OfficeAiCitation citation, OfficeAiDocument document, string cultureName) {
        if (citation.Quote is not { } quote || citation.QuoteStart is not { } quoteStart) return false;
        OfficeAiEvidence? evidence = document.Evidence.FirstOrDefault(item => item.Id == citation.EvidenceId);
        if (evidence is null) return false;
        string value = raw.Trim();
        if (value.Length == 0) return false;
        NumberFormatInfo format = CultureInfo.GetCultureInfo(cultureName).NumberFormat;
        for (int offset = quote.IndexOf(value, StringComparison.Ordinal); offset >= 0;
            offset = quote.IndexOf(value, offset + 1, StringComparison.Ordinal)) {
            int start = quoteStart + offset;
            int end = start + value.Length;
            string source = evidence.Text;
            if (start > 0 && char.IsDigit(source[start - 1])) continue;
            if (start > 1 && source[start - 1] is 'e' or 'E' && char.IsDigit(source[start - 2])) continue;
            if (end < source.Length && (char.IsDigit(source[end]) || source[end] is '-' or '+' or '\u2212')) continue;
            if (end < source.Length && source[end] is 'e' or 'E' && end + 1 < source.Length
                && (char.IsDigit(source[end + 1]) || source[end + 1] is '-' or '+')) continue;
            int before = SkipCurrencyContext(source, start - 1, -1, format);
            if (before >= 0 && (source[before] is '-' or '+' or '\u2212' or '(')) continue;
            int after = SkipCurrencyContext(source, end, 1, format);
            if (after < source.Length && source[after] is '-' or '+' or '\u2212') continue;
            if (HasNumericContinuation(source, start, backwards: true, format)
                || HasNumericContinuation(source, end, backwards: false, format)) continue;
            return true;
        }
        return false;
    }

    private static int SkipCurrencyContext(string source, int index, int direction, NumberFormatInfo format) {
        while (index >= 0 && index < source.Length) {
            if (char.IsWhiteSpace(source[index]) || char.GetUnicodeCategory(source[index]) == UnicodeCategory.CurrencySymbol) {
                index += direction;
                continue;
            }
            string symbol = format.CurrencySymbol;
            int symbolStart = direction < 0 ? index - symbol.Length + 1 : index;
            if (symbol.Length > 0 && symbolStart >= 0 && symbolStart + symbol.Length <= source.Length
                && source.AsSpan(symbolStart, symbol.Length).SequenceEqual(symbol.AsSpan())) {
                index += direction * symbol.Length;
                continue;
            }
            int codeStart = direction < 0 ? index - 2 : index;
            if (codeStart >= 0 && codeStart + 3 <= source.Length
                && source[codeStart] is >= 'A' and <= 'Z' && source[codeStart + 1] is >= 'A' and <= 'Z'
                && source[codeStart + 2] is >= 'A' and <= 'Z'
                && (codeStart == 0 || !char.IsLetter(source[codeStart - 1]))
                && (codeStart + 3 == source.Length || !char.IsLetter(source[codeStart + 3]))) {
                index += direction * 3;
                continue;
            }
            break;
        }
        return index;
    }

    private static bool HasNumericContinuation(string source, int boundary, bool backwards, NumberFormatInfo format) {
        // Dot/comma also catch an incompatible-culture token being shortened to a valid integer.
        string[] separators = { format.NumberGroupSeparator, format.NumberDecimalSeparator, ".", ",", "\u00a0", "\u202f", " " };
        foreach (string separator in separators) {
            if (separator.Length == 0) continue;
            int start = backwards ? boundary - separator.Length : boundary;
            int digit = backwards ? start - 1 : start + separator.Length;
            if (backwards && separator == format.NumberDecimalSeparator && start >= 0
                && source.AsSpan(start, separator.Length).SequenceEqual(separator.AsSpan())) return true;
            if (start >= 0 && start + separator.Length <= source.Length && digit >= 0 && digit < source.Length
                && source.AsSpan(start, separator.Length).SequenceEqual(separator.AsSpan()) && char.IsDigit(source[digit])) return true;
        }
        return false;
    }
}
