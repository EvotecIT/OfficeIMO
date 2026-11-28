using System;
using System.Collections.Generic;

namespace OfficeIMO.Entities.Internal;

internal static class IdentifierReadings {
    internal static bool Submit(EntityRecognitionContext context, EntityScheme scheme, int start, int end,
        EntityCandidateStrength strength, EntityEvidence evidence = EntityEvidence.None, bool needsContext = false,
        bool optionalPrefix = false, string? validationInput = null, Func<string, bool>? plausibleDate = null) {
        if (!context.IsEnabled(scheme) || end <= start || needsContext && !context.HasContext(scheme, start)) return false;

        string written = context.Text.Substring(start, end - start);
        EntityValidationResult validation = scheme.Validate(validationInput ?? written);
        if (!validation.IsValid) return false;
        string normalized = validation.Normalized!;
        if (!string.Equals(written, normalized, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(written, scheme.Format(normalized), StringComparison.OrdinalIgnoreCase)) {
            evidence |= EntityEvidence.IssuerFormatting;
        }

        if (plausibleDate?.Invoke(normalized) == true) evidence |= EntityEvidence.PlausibleDate;

        context.Add(new EntityCandidate(validation, start, end - start, strength, evidence) {
            RequiresContext = needsContext,
            OptionalPrefix = optionalPrefix,
            RequiresPlausibleDate = plausibleDate != null,
        });
        return true;
    }

    // A recognizer may extend an alphanumeric atom across single separators, within its scheme's bound.
    // Endpoints always occur at atom boundaries so a check cannot match the middle of a longer token.
    internal static IEnumerable<(int End, int Count)> Endpoints(string text, EntityToken token, int maxCharacters,
        bool spacesOnly = false, bool vatPunctuation = false) {
        int end = token.Start + token.Length;
        int count = token.CharacterCount;
        while (count <= maxCharacters) {
            if (end < text.Length && EntityContextMatcher.IsWord(text[end])) yield break;
            yield return (end, count);
            if (end >= text.Length) yield break;
            int next = end;
            if (spacesOnly ? EntityCharacters.IsSpace(text[end]) : EntityCharacters.IsSeparator(text[end])) {
                next++;
            } else if (!vatPunctuation || text[end] is not ('+' or '*')) {
                yield break;
            }
            if (next >= text.Length || !IsCandidateCharacter(text, next, vatPunctuation)) yield break;

            end = next;
            do {
                end++;
                if (++count > maxCharacters) yield break;
            } while (end < text.Length && IsCandidateCharacter(text, end, vatPunctuation));
        }
    }

    private static bool IsCandidateCharacter(string text, int position, bool vatPunctuation) =>
        EntityTokenizer.IsAlphaAt(text, position) || vatPunctuation && text[position] is '+' or '*';

    internal static bool IsRepeated(string value, int start, int? end = null) {
        int limit = end ?? value.Length;
        if (start >= limit) return false;
        char first = value[start];
        if (!EntityCharacters.IsDigit(first)) return false;
        for (int index = start + 1; index < limit; index++) {
            if (value[index] != first) return false;
        }
        return true;
    }
}

internal sealed class TokenEntityRecognizer : IEntityRecognizer {
    private readonly Action<EntityRecognitionContext, EntityToken> _recognize;
    private readonly int _maximumDigits;

    internal TokenEntityRecognizer(string id, EntityTokenShape shapes, Action<EntityRecognitionContext, EntityToken> recognize,
        params EntityScheme[] schemes) {
        Id = id;
        Shapes = shapes;
        Schemes = Array.AsReadOnly(schemes);
        foreach (EntityScheme scheme in schemes) _maximumDigits = Math.Max(_maximumDigits, scheme.MaxLength);
        _recognize = recognize;
    }

    public string Id { get; }
    public IReadOnlyList<EntityScheme> Schemes { get; }
    public EntityTokenShape Shapes { get; }

    public void Recognize(EntityRecognitionContext context) {
        foreach (EntityToken token in context.GetTokens(Shapes)) {
            _recognize(context, token);
            if (token.Shape != EntityTokenShape.DigitRun || token.Length == token.CharacterCount || context.RejectDigitRun(token)) continue;
            context.DigitGroupRun = token;
            try {
                foreach (EntityToken candidate in EntityDigitGroups.Candidates(context.Text, token, _maximumDigits, context.CancellationToken)) {
                    context.CancellationToken.ThrowIfCancellationRequested();
                    _recognize(context, candidate);
                }
            } finally {
                context.DigitGroupRun = null;
            }
        }
    }
}
