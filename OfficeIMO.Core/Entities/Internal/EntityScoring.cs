using System;

namespace OfficeIMO.Entities.Internal;

internal sealed class ScoredEntity {
    internal ScoredEntity(EntityCandidate candidate, EntityEvidence evidence, EntityKeyword? keyword, int registrationOrder) {
        Candidate = candidate;
        Evidence = evidence;
        Keyword = keyword;
        RegistrationOrder = registrationOrder;
        RawConfidence = EntityScoring.Score(candidate, evidence);
    }

    internal EntityCandidate Candidate { get; }
    internal EntityEvidence Evidence { get; set; }
    internal EntityKeyword? Keyword { get; }
    internal int RegistrationOrder { get; }
    internal double RawConfidence { get; set; }
    internal double Confidence => Math.Min(1, Math.Max(0, Math.Round(RawConfidence, 2)));
    internal EntityScheme Scheme => Candidate.Validation.Scheme;
    internal int Start => Candidate.Start;
    internal int End => Start + Candidate.Length;
    internal bool HasContext => (Evidence & (EntityEvidence.Label | EntityEvidence.ContextKeyword | EntityEvidence.StructuredSource)) != 0;
}

internal static class EntityScoring {
    internal static double Score(EntityCandidate candidate, EntityEvidence evidence) {
        if ((evidence & EntityEvidence.StructuredSource) != 0) return 1;
        double confidence = candidate.Strength switch {
            EntityCandidateStrength.StrongCheck => 0.85,
            EntityCandidateStrength.DistinctiveShape => 0.80,
            EntityCandidateStrength.PrefixedShape => 0.60,
            EntityCandidateStrength.DigitCheck => 0.40,
            _ => 0.20,
        };
        if ((evidence & EntityEvidence.Label) != 0) confidence += 0.50;
        if ((evidence & EntityEvidence.ContextKeyword) != 0) confidence += 0.40;
        if ((evidence & EntityEvidence.IssuerFormatting) != 0) confidence += 0.10;
        if (candidate.OptionalPrefix && (evidence & EntityEvidence.CountryPrefix) != 0) confidence += 0.10;
        if ((evidence & EntityEvidence.KnownIssuerRange) != 0) confidence += 0.10;
        if (candidate.RequiresPlausibleDate && (evidence & EntityEvidence.PlausibleDate) == 0) confidence -= 0.30;
        return confidence;
    }
}
