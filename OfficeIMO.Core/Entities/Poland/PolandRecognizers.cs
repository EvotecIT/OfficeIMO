using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

internal static class PolandRecognizers {
    internal static void Recognize(EntityRecognitionContext context, EntityToken token) {
        if (token.TouchesWord) return;
        if (token.Shape == EntityTokenShape.DigitRun) {
            RecognizeDigits(context, token);
            return;
        }

        string text = context.Text;
        if (token.CharacterCount is < 2 or > 12 || !EntityCharacters.IsLetter(EntityCharacters.Upper(text[token.Start]))) return;
        if (EntityTokenizer.StartsWith(text, token.Start, "PL") && !context.HasRecognizerFor("EU.VAT.PL")) {
            foreach ((int end, int count) in IdentifierReadings.Endpoints(text, token, 12)) {
                if (count == 12) IdentifierReadings.Submit(context, PolandSchemes.Nip, token.Start, end,
                    EntityCandidateStrength.DigitCheck, EntityEvidence.CountryPrefix, optionalPrefix: true);
            }
        }

        // The identity document is recognized only in a field that identifies the document type.
        if (token.CharacterCount is not (2 or 3 or 9)) return;
        EntityScheme scheme = token.CharacterCount == 3 ||
            token.CharacterCount == 9 && EntityCharacters.IsLetter(EntityCharacters.Upper(text[token.Start + 2]))
            ? PolandSchemes.IdentityCard : PolandSchemes.Passport;
        if (!context.IsEnabled(scheme) || !context.HasContext(scheme, token.Start)) return;
        foreach ((int end, int count) in IdentifierReadings.Endpoints(text, token, 9)) {
            if (count == 9) IdentifierReadings.Submit(context, scheme, token.Start, end, EntityCandidateStrength.ShortLetterDigitCheck, needsContext: true);
        }
    }

    private static void RecognizeDigits(EntityRecognitionContext context, EntityToken token) {
        if (token.CharacterCount > 14 || context.RejectDigitRun(token)) return;
        int end = token.Start + token.Length;
        switch (token.CharacterCount) {
            case 9:
            case 14:
                IdentifierReadings.Submit(context, PolandSchemes.Regon, token.Start, end, EntityCandidateStrength.DigitCheck);
                break;
            case 10:
                EntityEvidence evidence = MatchesNipGrouping(context.Text, token) ? EntityEvidence.IssuerFormatting : EntityEvidence.None;
                IdentifierReadings.Submit(context, PolandSchemes.Nip, token.Start, end, EntityCandidateStrength.DigitCheck, evidence);
                break;
            case 11:
                IdentifierReadings.Submit(context, PolandSchemes.Pesel, token.Start, end, EntityCandidateStrength.DigitCheck,
                    plausibleDate: PeselScheme.HasPlausibleDate);
                break;
        }

        if (token.CharacterCount > 10 || !context.IsEnabled(PolandSchemes.Krs) || !context.HasContext(PolandSchemes.Krs, token.Start)) return;
        string written = context.Text.Substring(token.Start, token.Length);
        string padded = new string('0', 10 - token.CharacterCount) + written;
        IdentifierReadings.Submit(context, PolandSchemes.Krs, token.Start, end, EntityCandidateStrength.ShapeOnly, needsContext: true, validationInput: padded);
    }

    internal static bool MatchesNipGrouping(string text, EntityToken token) {
        int start = token.Start;
        if (token.Length != 13) return false;
        return text[start + 3] == '-' && text[start + 6] == '-' && text[start + 9] == '-' ||
               text[start + 3] == '-' && text[start + 7] == '-' && text[start + 10] == '-';
    }
}
