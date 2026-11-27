using System;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

internal static class InternetRecognizers {
    internal static void Recognize(EntityRecognitionContext context, EntityToken token) {
        if (token.Shape == EntityTokenShape.At && context.IsEnabled(InternetSchemes.Email))
            FindEmail(context, token.Start);
        if (token.Shape == EntityTokenShape.Scheme && context.IsEnabled(InternetSchemes.Url))
            FindUrl(context, token.Start);
    }

    private static void FindEmail(EntityRecognitionContext context, int at) {
        string text = context.Text;
        if (at == 0 || at + 1 == text.Length || text[at + 1] == '[') return;
        int minimum = Math.Max(0, at - 64);
        int start = at;
        if (text[start - 1] == '"') {
            start--;
            do {
                start--;
            } while (start >= minimum && (text[start] != '"' || IsEscaped(text, start, minimum)));
            if (start < minimum || start > 0 && EmailScheme.IsAtom(text[start - 1])) return;
        } else {
            while (start > minimum && (EmailScheme.IsAtom(text[start - 1]) || text[start - 1] == '.')) start--;
            if (start > 0 && (EmailScheme.IsAtom(text[start - 1]) || text[start - 1] is '.' or '@' or '"')) return;
        }

        int end = at + 1;
        int limit = Math.Min(text.Length, start + InternetSchemes.Email.MaxLength + 1);
        while (end < limit && EmailScheme.IsDomainCharacter(text[end])) end++;
        if (end == limit && end < text.Length && EmailScheme.IsDomainCharacter(text[end]) ||
            end < text.Length && (text[end] == '@' || EntityContextMatcher.IsWord(text[end]))) return;
        while (end > at + 1 && text[end - 1] == '.') end--;
        if (end <= at + 1) return;

        EntityValidationResult validation = InternetSchemes.EmailInstance.ValidateMailbox(text.Substring(start, end - start),
            allowLiteral: false, out _);
        if (validation.IsValid)
            context.Add(new EntityCandidate(validation, start, end - start, EntityCandidateStrength.DistinctiveShape));
    }

    private static bool IsEscaped(string text, int position, int minimum) {
        int escapes = 0;
        for (int index = position - 1; index >= minimum && text[index] == '\\'; index--) escapes++;
        return (escapes & 1) != 0;
    }

    private static void FindUrl(EntityRecognitionContext context, int start) {
        string text = context.Text;
        int prefixLength = EntityTokenizer.UrlPrefixLength(text, start);
        int limit = Math.Min(text.Length, start + InternetSchemes.Url.MaxLength);
        int end = start + prefixLength;
        while (end < limit && !UrlScheme.IsTerminator(text[end])) end++;
        if (end == limit && end < text.Length && !UrlScheme.IsTerminator(text[end])) return;
        end = TrimUrl(text, start, end);
        if (end <= start + prefixLength) return;
        IdentifierReadings.Submit(context, InternetSchemes.Url, start, end, EntityCandidateStrength.DistinctiveShape);
    }

    private static int TrimUrl(string text, int start, int end) {
        int parentheses = 0, squares = 0, braces = 0;
        for (int index = start; index < end; index++) {
            switch (text[index]) {
                case '(': parentheses++; break;
                case ')': parentheses--; break;
                case '[': squares++; break;
                case ']': squares--; break;
                case '{': braces++; break;
                case '}': braces--; break;
            }
        }

        while (end > start) {
            switch (text[end - 1]) {
                case '.' or ',' or ';' or ':' or '!' or '?' or '\'':
                    break;
                case ')' when parentheses++ < 0:
                case ']' when squares++ < 0:
                case '}' when braces++ < 0:
                    break;
                default:
                    return end;
            }
            end--;
        }
        return end;
    }
}
