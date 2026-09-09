using System.Globalization;

namespace OfficeIMO.AI;

public sealed partial class OfficeAiEngine {
    private static bool TryValidateValueEvidence(string raw, IReadOnlyList<OfficeAiCitation> citations, Batch batch,
        OfficeAiDocument document, Func<string, int, int, bool> isComplete, out IReadOnlyList<OfficeAiCitation> locatedCitations) {
        bool supported = false;
        var located = new List<OfficeAiCitation>(citations.Count);
        foreach (OfficeAiCitation citation in citations) {
            if (batch.Images.ContainsKey(citation.EvidenceId)) { supported = true; located.Add(citation); continue; }
            OfficeAiCitation? complete = FindCompleteValueCitation(raw, citation, batch, document, isComplete);
            supported |= complete is not null;
            located.Add(complete ?? citation);
        }
        locatedCitations = located.AsReadOnly();
        return supported;
    }

    private static OfficeAiCitation? FindCompleteValueCitation(string raw, OfficeAiCitation citation, Batch batch,
        OfficeAiDocument document, Func<string, int, int, bool> isComplete) {
        if (citation.Quote is not { } quote) return null;
        string value = raw.Trim();
        OfficeAiEvidence? original = document.Evidence.FirstOrDefault(item => item.Id == citation.EvidenceId);
        if (original is null || value.Length == 0) return null;
        foreach (var item in batch.Evidence) {
            EvidenceSlice slice = batch.Slices.TryGetValue(item.Key, out var fragment) ? fragment : new(item.Key, 0, item.Value.Text.Length);
            if (slice.OriginalId != citation.EvidenceId) continue;
            // Search only the evidence actually sent in this batch; the complete-value check may
            // inspect surrounding original context, but cannot borrow an occurrence from an omitted slice.
            string observed = item.Value.Text;
            for (int offset = observed.IndexOf(value, StringComparison.Ordinal); offset >= 0;
                offset = observed.IndexOf(value, offset + 1, StringComparison.Ordinal)) {
                if (!isComplete(original.Text, slice.Start + offset, value.Length)) continue;
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

    private static bool IsCompleteWordValueAt(string source, int start, int length) =>
        (start == 0 || !IsWordAt(source, start - 1))
        && (start + length == source.Length || !IsWordAt(source, start + length));

    private static bool IsCompleteDateAt(string source, int start, int length, string raw, DateTimeFormatInfo format) {
        if (!IsCompleteWordValueAt(source, start, length)) return false;
        // Separators joining another date component are part of the source value; ordinary
        // sentence punctuation and spaced delimiters remain valid citation boundaries.
        foreach (string separator in DateValueSeparators(raw, format)) {
            if (separator.Length == 0 || string.IsNullOrWhiteSpace(separator)) continue;
            int before = start - separator.Length;
            int after = start + length;
            if (before > 0 && source.AsSpan(before, separator.Length).SequenceEqual(separator.AsSpan())
                && IsWordAt(source, before - 1)) return false;
            if (after + separator.Length < source.Length && source.AsSpan(after, separator.Length).SequenceEqual(separator.AsSpan())
                && IsWordAt(source, after + separator.Length)) return false;
        }
        return true;
    }

    private static IEnumerable<string> DateValueSeparators(string raw, DateTimeFormatInfo format) {
        yield return "-";
        yield return "/";
        yield return ".";
        yield return format.DateSeparator;
        // TryParseExact has already accepted this value against the requested format. Its
        // observed punctuation therefore includes custom, escaped and culture-specific
        // literals without duplicating the runtime's date-format grammar.
        for (int index = 0; index < raw.Length; index++) {
            if (char.IsWhiteSpace(raw[index]) || IsWordAt(raw, index)) continue;
            int start = index;
            while (index + 1 < raw.Length && !char.IsWhiteSpace(raw[index + 1]) && !IsWordAt(raw, index + 1)) index++;
            yield return raw.Substring(start, index - start + 1);
        }
    }

    private static bool IsWordAt(string source, int index) {
        if (char.IsLowSurrogate(source[index]) && index > 0 && char.IsHighSurrogate(source[index - 1])) index--;
        return CharUnicodeInfo.GetUnicodeCategory(source, index) is UnicodeCategory.UppercaseLetter
            or UnicodeCategory.LowercaseLetter or UnicodeCategory.TitlecaseLetter or UnicodeCategory.ModifierLetter
            or UnicodeCategory.OtherLetter or UnicodeCategory.DecimalDigitNumber or UnicodeCategory.LetterNumber
            or UnicodeCategory.OtherNumber or UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark
            or UnicodeCategory.EnclosingMark or UnicodeCategory.ConnectorPunctuation;
    }
}
