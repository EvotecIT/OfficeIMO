using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

internal static class FinancialRecognizers {
    internal static void Recognize(EntityRecognitionContext context, EntityToken token) {
        if (token.TouchesWord) return;
        string text = context.Text;
        int end = token.Start + token.Length;
        if (token.Shape == EntityTokenShape.DigitRun) {
            if (token.CharacterCount < 12 || token.CharacterCount > 19 || context.RejectDigitRun(token)) return;
            bool knownRange = text[token.Start] is >= '3' and <= '6';
            IdentifierReadings.Submit(context, FinancialSchemes.Card, token.Start, end,
                knownRange ? EntityCandidateStrength.DigitCheck : EntityCandidateStrength.ShapeOnly,
                knownRange ? EntityEvidence.KnownIssuerRange : EntityEvidence.None);
            return;
        }

        if (token.CharacterCount == 20) {
            IdentifierReadings.Submit(context, FinancialSchemes.Lei, token.Start, end, EntityCandidateStrength.StrongCheck);
        }

        if (token.CharacterCount < 2 || !EntityCharacters.IsLetter(EntityCharacters.Upper(text[token.Start])) ||
            !EntityCharacters.IsLetter(EntityCharacters.Upper(text[token.Start + 1]))) return;
        if (token.CharacterCount == 12)
            IdentifierReadings.Submit(context, FinancialSchemes.Isin, token.Start, end, EntityCandidateStrength.StrongCheck);
        if (token.CharacterCount is 8 or 11)
            IdentifierReadings.Submit(context, FinancialSchemes.Bic, token.Start, end, EntityCandidateStrength.ShapeOnly, needsContext: true);

        if (!context.IsEnabled(FinancialSchemes.Iban) || token.CharacterCount < 4 ||
            !EntityCharacters.IsDigit(text[token.Start + 2]) || !EntityCharacters.IsDigit(text[token.Start + 3])) return;
        foreach ((int candidateEnd, int count) in IdentifierReadings.Endpoints(text, token, FinancialSchemes.Iban.MaxLength, spacesOnly: true)) {
            if (count >= FinancialSchemes.Iban.MinLength &&
                IdentifierReadings.Submit(context, FinancialSchemes.Iban, token.Start, candidateEnd, EntityCandidateStrength.StrongCheck,
                    EntityEvidence.CountryPrefix)) break;
        }
    }
}
