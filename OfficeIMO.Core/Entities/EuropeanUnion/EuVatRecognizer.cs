using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

internal static class EuVatRecognizer {
    internal static void Recognize(EntityRecognitionContext context, EntityToken token) {
        if (token.TouchesWord) return;
        if (token.Shape == EntityTokenShape.DigitRun) {
            RecognizeUnprefixed(context, token);
            return;
        }

        if (token.CharacterCount is < 2 or > 14 || !EntityCharacters.IsLetter(EntityCharacters.Upper(context.Text[token.Start])) ||
            !EntityCharacters.IsLetter(EntityCharacters.Upper(context.Text[token.Start + 1]))) return;
        string prefix = context.Text.Substring(token.Start, 2);
        if (EuVatSchemes.TryGet(prefix, out PrefixedVatScheme? scheme) && context.IsEnabled(scheme!))
            RecognizePrefixed(context, token, scheme!);
    }

    private static void RecognizeUnprefixed(EntityRecognitionContext context, EntityToken token) {
        if (token.CharacterCount < 2 || token.CharacterCount > 12 || context.RejectDigitRun(token)) return;
        foreach (PrefixedVatScheme scheme in EuVatSchemes.All) {
            if (token.CharacterCount + 2 < scheme.MinLength || token.CharacterCount + 2 > scheme.MaxLength || !context.IsEnabled(scheme)) continue;
            if (context.HasContext(scheme, token.Start)) Submit(context, scheme, token.Start, token.Start + token.Length, prefixed: false);
        }
    }

    internal static void RecognizePrefixed(EntityRecognitionContext context, EntityToken token, PrefixedVatScheme scheme) {
        int bestEnd = 0;
        foreach ((int end, int count) in IdentifierReadings.Endpoints(context.Text, token, scheme.MaxLength,
                     vatPunctuation: scheme.AllowsVatPunctuation)) {
            if (count < scheme.MinLength) continue;
            // Extend an already valid spelling only for the UK branch-number variant.
            if (bestEnd > 0 && (scheme.Country != "GB" || count != 14)) break;
            if (scheme.IsValid(context.Text.Substring(token.Start, end - token.Start))) bestEnd = end;
        }
        if (bestEnd > 0) Submit(context, scheme, token.Start, bestEnd, prefixed: true);
    }

    internal static void Submit(EntityRecognitionContext context, PrefixedVatScheme scheme, int start, int end, bool prefixed) {
        EntityValidationResult validation = scheme.Validate(context.Text.Substring(start, end - start));
        if (!validation.IsValid || IdentifierReadings.IsRepeated(validation.Normalized!, 2)) return;
        if (!prefixed && validation.CheckDigit != EntityCheckOutcome.Passed) return;

        EntityCandidateStrength strength = !prefixed ? EntityCandidateStrength.DigitCheck :
            validation.CheckDigit == EntityCheckOutcome.Passed ? EntityCandidateStrength.StrongCheck : EntityCandidateStrength.PrefixedShape;
        EntityEvidence evidence = prefixed ? EntityEvidence.CountryPrefix : EntityEvidence.None;
        context.Add(new EntityCandidate(validation, start, end - start, strength, evidence));
    }
}
