using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

internal static class UnitedStatesRecognizers {
    internal static void Recognize(EntityRecognitionContext context, EntityToken token) {
        if (token.CharacterCount != 9 || context.RejectDigitRun(token)) return;
        int end = token.Start + token.Length;
        IdentifierReadings.Submit(context, UnitedStatesSchemes.Ssn, token.Start, end, EntityCandidateStrength.ShapeOnly);
        IdentifierReadings.Submit(context, UnitedStatesSchemes.Ein, token.Start, end, EntityCandidateStrength.ShapeOnly,
            EntityEvidence.KnownIssuerRange);
    }
}
