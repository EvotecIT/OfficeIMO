using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

internal static class UnitedKingdomRecognizers {
    internal static void Recognize(EntityRecognitionContext context, EntityToken token) {
        if (token.TouchesWord) return;
        string text = context.Text;
        if (token.Shape == EntityTokenShape.DigitRun) {
            if (token.CharacterCount == 8 && !context.RejectDigitRun(token)) {
                IdentifierReadings.Submit(context, UnitedKingdomSchemes.Company, token.Start, token.Start + token.Length,
                    EntityCandidateStrength.ShapeOnly, needsContext: true);
            }
            if (token.CharacterCount is 9 or 12 && !context.RejectDigitRun(token) &&
                context.IsEnabled(UnitedKingdomSchemes.Vat) && context.HasContext(UnitedKingdomSchemes.Vat, token.Start)) {
                EuVatRecognizer.Submit(context, (PrefixedVatScheme)UnitedKingdomSchemes.Vat,
                    token.Start, token.Start + token.Length, prefixed: false);
            }
            return;
        }

        if (token.CharacterCount < 2 || !EntityCharacters.IsLetter(EntityCharacters.Upper(text[token.Start]))) return;
        if (EntityTokenizer.StartsWith(text, token.Start, "GB") && context.IsEnabled(UnitedKingdomSchemes.Vat))
            EuVatRecognizer.RecognizePrefixed(context, token, (PrefixedVatScheme)UnitedKingdomSchemes.Vat);

        if (token.CharacterCount == 8) {
            IdentifierReadings.Submit(context, UnitedKingdomSchemes.Company, token.Start, token.Start + token.Length,
                EntityCandidateStrength.ShapeOnly, needsContext: true);
        }

        if (!context.IsEnabled(UnitedKingdomSchemes.Nino) || token.CharacterCount is not (2 or 8 or 9)) return;
        int longest = 0;
        bool hasSuffix = false;
        foreach ((int end, int count) in IdentifierReadings.Endpoints(text, token, 9)) {
            bool suffix = count == 9 && EntityCharacters.Upper(text[end - 1]) is >= 'A' and <= 'D';
            if (count != 8 && !suffix) continue;
            longest = end;
            hasSuffix = suffix;
        }
        if (longest > 0) IdentifierReadings.Submit(context, UnitedKingdomSchemes.Nino, token.Start, longest,
            hasSuffix ? EntityCandidateStrength.DistinctiveShape : EntityCandidateStrength.ShapeOnly, needsContext: !hasSuffix);
    }
}
