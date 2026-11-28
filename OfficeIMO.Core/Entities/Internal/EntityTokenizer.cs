using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Entities.Internal;

internal static class EntityTokenizer {
    internal static List<EntityToken> Tokenize(string text, EntityTokenShape shapes, CancellationToken cancellationToken) {
        var tokens = new List<EntityToken>();
        if (shapes == EntityTokenShape.None) return tokens;

        int digitEnd = 0;
        int nextCheck = 0;
        for (int start = 0; start < text.Length;) {
            CheckCancellation(start, ref nextCheck, cancellationToken);
            char first = text[start];
            if ((shapes & EntityTokenShape.DigitRun) != 0 && start >= digitEnd && EntityCharacters.IsDigit(first)) {
                digitEnd = ReadDigits(text, start, ref nextCheck, cancellationToken, out int digitCount);
                tokens.Add(new EntityToken(EntityTokenShape.DigitRun, start, digitEnd - start, digitCount, TouchesWord(text, start, digitEnd)));
            }

            if ((shapes & EntityTokenShape.At) != 0 && first == '@')
                tokens.Add(new EntityToken(EntityTokenShape.At, start, 1, 1, TouchesWord(text, start, start + 1)));
            if ((shapes & EntityTokenShape.Scheme) != 0 &&
                (start == 0 || !EntityContextMatcher.IsWord(text[start - 1]) && text[start - 1] != '.') && UrlPrefixLength(text, start) > 0) {
                tokens.Add(new EntityToken(EntityTokenShape.Scheme, start, 1, 1, false));
            }

            if (!EntityCharacters.IsAlphaNumeric(first)) {
                start++;
                continue;
            }

            int end = start + 1;
            while (end < text.Length && IsAlphaAt(text, end)) {
                CheckCancellation(end, ref nextCheck, cancellationToken);
                end++;
            }
            if ((shapes & EntityTokenShape.AlphaNumericRun) != 0)
                tokens.Add(new EntityToken(EntityTokenShape.AlphaNumericRun, start, end - start, end - start, TouchesWord(text, start, end)));
            start = end;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return tokens;
    }

    private static int ReadDigits(string text, int start, ref int nextCheck, CancellationToken cancellationToken, out int count) {
        int end = start;
        count = 0;
        do {
            CheckCancellation(end, ref nextCheck, cancellationToken);
            count++;
            end++;
            if (end + 1 < text.Length && EntityCharacters.IsSeparator(text[end]) && EntityCharacters.IsDigit(text[end + 1])) end++;
        } while (end < text.Length && EntityCharacters.IsDigit(text[end]));
        return end;
    }

    private static void CheckCancellation(int position, ref int nextCheck, CancellationToken cancellationToken) {
        if (position < nextCheck) return;
        cancellationToken.ThrowIfCancellationRequested();
        nextCheck = position + 32_768; // 64 KiB of UTF-16 text.
    }

    internal static bool IsAlphaAt(string text, int index) =>
        EntityCharacters.IsAlphaNumeric(text[index]) ||
        (text[index] is '*' or '+') && index > 0 && index + 1 < text.Length &&
        EntityCharacters.IsAlphaNumeric(text[index - 1]) && EntityCharacters.IsAlphaNumeric(text[index + 1]);

    internal static bool TouchesWord(string text, int start, int end) =>
        start > 0 && EntityContextMatcher.IsWord(text[start - 1]) ||
        end < text.Length && EntityContextMatcher.IsWord(text[end]);

    internal static bool StartsWith(string text, int start, string value) =>
        text.Length - start >= value.Length && string.Compare(text, start, value, 0, value.Length, StringComparison.OrdinalIgnoreCase) == 0;

    internal static int UrlPrefixLength(string text, int start) => EntityCharacters.Upper(text[start]) switch {
        'H' when StartsWith(text, start, "https://") => 8,
        'H' when StartsWith(text, start, "http://") => 7,
        'F' when StartsWith(text, start, "ftp://") => 6,
        'M' when StartsWith(text, start, "mailto:") => 7,
        'W' when StartsWith(text, start, "www.") => 4,
        _ => 0,
    };
}
