using System.Globalization;

namespace OfficeIMO.AI;

public sealed partial class OfficeAiEngine {
    private static readonly string[] CurrencyTokens = CreateCurrencyTokens();
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
        if (end < source.Length && (char.IsDigit(source[end]) || HasNumberSign(source, end, backwards: false, format))) return false;
        if (end < source.Length && source[end] is 'e' or 'E' && end + 1 < source.Length
            && (char.IsDigit(source[end + 1]) || HasNumberSign(source, end + 1, backwards: false, format))) return false;
        int before = SkipCurrencyContext(source, start - 1, -1);
        if (before >= 0 && (source[before] == '(' || HasNumberSign(source, before, backwards: true, format))) return false;
        int after = SkipCurrencyContext(source, end, 1);
        if (after < source.Length && HasNumberSign(source, after, backwards: false, format)) return false;
        return !HasNumericContinuation(source, start, backwards: true, format)
            && !HasNumericContinuation(source, end, backwards: false, format);
    }

    private static bool HasNumberSign(string source, int index, bool backwards, NumberFormatInfo format) {
        if (source[index] is '-' or '+' or '\u2212') return true;
        foreach (string sign in new[] { format.NegativeSign, format.PositiveSign }) {
            int start = backwards ? index - sign.Length + 1 : index;
            if (sign.Length > 0 && start >= 0 && start + sign.Length <= source.Length
                && source.AsSpan(start, sign.Length).SequenceEqual(sign.AsSpan())) return true;
        }
        return false;
    }

    private static int SkipCurrencyContext(string source, int index, int direction) {
        while (index >= 0 && index < source.Length) {
            if (char.IsWhiteSpace(source[index])) {
                index += direction;
                continue;
            }
            // Longest complete tokens must win before individual symbols: R$ cannot be reduced to $.
            int tokenLength = CurrencyTokenLength(source, index, direction);
            if (tokenLength > 0) {
                index += direction * tokenLength;
                continue;
            }
            if (char.GetUnicodeCategory(source[index]) == UnicodeCategory.CurrencySymbol) {
                index += direction;
                continue;
            }
            break;
        }
        return index;
    }

    private static int CurrencyTokenLength(string source, int index, int direction) {
        foreach (string token in CurrencyTokens) {
            int start = direction < 0 ? index - token.Length + 1 : index;
            int end = start + token.Length;
            if (start < 0 || end > source.Length
                || !source.AsSpan(start, token.Length).Equals(token.AsSpan(), StringComparison.OrdinalIgnoreCase)) continue;
            if (char.IsLetter(token[0]) && start > 0 && char.IsLetter(source[start - 1])) continue;
            if (char.IsLetter(token[^1]) && end < source.Length && (char.IsLetter(source[end])
                || (source[end] == '-' && end + 1 < source.Length && char.IsLetter(source[end + 1])))) continue;
            return token.Length;
        }
        return 0;
    }

    private static string[] CreateCurrencyTokens() {
        var tokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (CultureInfo culture in CultureInfo.GetCultures(CultureTypes.SpecificCultures)) {
            string symbol = culture.NumberFormat.CurrencySymbol;
            if (!string.IsNullOrWhiteSpace(symbol)) tokens.Add(symbol);
            try {
                var region = new RegionInfo(culture.Name);
                tokens.Add(region.ISOCurrencySymbol);
                if (symbol == "$") tokens.Add(region.TwoLetterISORegionName + "$");
            } catch (ArgumentException) { /* Some runtime cultures do not identify a geographic region. */ }
        }
        return tokens.OrderByDescending(token => token.Length).ThenBy(token => token, StringComparer.Ordinal).ToArray();
    }

    private static bool HasNumericContinuation(string source, int boundary, bool backwards, NumberFormatInfo format) {
        // Dot/comma also catch an incompatible-culture token being shortened to a valid integer.
        var separators = new List<string> { format.NumberGroupSeparator, format.NumberDecimalSeparator, ".", "," };
        // Normalize the common space variants only for cultures that actually group with whitespace.
        if (format.NumberGroupSeparator.Length > 0 && string.IsNullOrWhiteSpace(format.NumberGroupSeparator))
            separators.AddRange(new[] { "\u00a0", "\u202f", " " });
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
