using System.Globalization;

namespace OfficeIMO.AI;

public sealed partial class OfficeAiEngine {
    private static bool TryValidateNumericEvidence(string raw, IReadOnlyList<OfficeAiCitation> citations, Batch batch,
        OfficeAiDocument document, string cultureName, out IReadOnlyList<OfficeAiCitation> locatedCitations) {
        bool supported = false;
        var located = new List<OfficeAiCitation>(citations.Count);
        foreach (OfficeAiCitation citation in citations) {
            if (batch.Images.ContainsKey(citation.EvidenceId)) { supported = true; located.Add(citation); continue; }
            OfficeAiCitation? complete = FindCompleteNumberCitation(raw, citation, batch, document, cultureName);
            supported |= complete is not null;
            located.Add(complete ?? citation);
        }
        locatedCitations = located.AsReadOnly();
        return supported;
    }

    private static OfficeAiCitation? FindCompleteNumberCitation(string raw, OfficeAiCitation citation, Batch batch,
        OfficeAiDocument document, string cultureName) {
        if (citation.Quote is not { } quote) return null;
        string value = raw.Trim();
        OfficeAiEvidence? original = document.Evidence.FirstOrDefault(item => item.Id == citation.EvidenceId);
        if (original is null || value.Length == 0) return null;
        NumberFormatInfo format = CultureInfo.GetCultureInfo(cultureName).NumberFormat;
        foreach (var item in batch.Evidence) {
            EvidenceSlice slice = batch.Slices.TryGetValue(item.Key, out var fragment) ? fragment : new(item.Key, 0, item.Value.Text.Length);
            if (slice.OriginalId != citation.EvidenceId) continue;
            // Search only the evidence actually sent in this batch; the complete-number check may
            // inspect surrounding original context, but cannot borrow an occurrence from an omitted slice.
            string observed = item.Value.Text;
            for (int offset = observed.IndexOf(value, StringComparison.Ordinal); offset >= 0;
                offset = observed.IndexOf(value, offset + 1, StringComparison.Ordinal)) {
                if (!IsCompleteNumberAt(original.Text, slice.Start + offset, value.Length, format)) continue;
                for (int withinQuote = quote.IndexOf(value, StringComparison.Ordinal); withinQuote >= 0;
                    withinQuote = quote.IndexOf(value, withinQuote + 1, StringComparison.Ordinal)) {
                    int quoteStart = offset - withinQuote;
                    if (quoteStart >= 0 && quoteStart + quote.Length <= observed.Length
                        && observed.AsSpan(quoteStart, quote.Length).SequenceEqual(quote.AsSpan()))
                        return citation with { QuoteStart = slice.Start + quoteStart };
                }
            }
        }
        return null;
    }

    // Substring provenance alone can drop a sign, leading digits, or a fractional part.
    // Check the original observation, including context outside a quote or request slice.
    private static bool IsCompleteNumberAt(string source, int start, int length, NumberFormatInfo format) {
        int end = start + length;
        if (start > 0 && char.IsDigit(source[start - 1])) return false;
        if (start > 1 && source[start - 1] is 'e' or 'E' && char.IsDigit(source[start - 2])) return false;
        if (end < source.Length && (char.IsDigit(source[end]) || source[end] is '-' or '+' or '\u2212')) return false;
        if (end < source.Length && source[end] is 'e' or 'E' && end + 1 < source.Length
            && (char.IsDigit(source[end + 1]) || source[end + 1] is '-' or '+')) return false;
        int before = SkipCurrencyContext(source, start - 1, -1, format);
        if (before >= 0 && (source[before] is '-' or '+' or '\u2212' or '(')) return false;
        int after = SkipCurrencyContext(source, end, 1, format);
        if (after < source.Length && source[after] is '-' or '+' or '\u2212') return false;
        return !HasNumericContinuation(source, start, backwards: true, format)
            && !HasNumericContinuation(source, end, backwards: false, format);
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
