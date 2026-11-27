using System;
using System.Collections.Generic;
using System.Globalization;
using OfficeIMO.Entities.Internal;

namespace OfficeIMO.Entities;

/// <summary>A scored reading and its coordinates in the source text.</summary>
public sealed class EntityMatch {
    internal EntityMatch(ScoredEntity reading, EntitySegmentChain chain,
        EntityMatch[]? alternatives = null, EntityMatch[]? contained = null, EntityMatch[]? maskingFallbacks = null) {
        EntityMatchPart[] parts = chain.GetParts(reading.Start, reading.End);
        Parts = Array.AsReadOnly(parts);
        Start = parts[0].Start;
        Length = parts[0].Length;
        Validation = reading.Candidate.Validation;
        Confidence = reading.Confidence;
        Evidence = reading.Evidence;
        ContextKeyword = reading.Keyword?.Value;
        ContextLanguage = reading.Keyword?.Language;
        SegmentIndex = parts[0].SegmentIndex;
        Location = chain.SegmentAt(reading.Start).Location;
        Alternatives = alternatives == null ? Array.Empty<EntityMatch>() : Array.AsReadOnly(alternatives);
        Contained = contained == null ? Array.Empty<EntityMatch>() : Array.AsReadOnly(contained);
        MaskingFallbacks = maskingFallbacks ?? Array.Empty<EntityMatch>();
    }

    /// <summary>UTF-16 offset relative to the first source segment.</summary>
    public int Start { get; }
    /// <summary>UTF-16 length in the first source segment, including written separators.</summary>
    public int Length { get; }
    /// <summary>Category of the recognized entity.</summary>
    public EntityKind Kind => Scheme.Kind;
    /// <summary>Scheme that validated the reading.</summary>
    public EntityScheme Scheme => Validation.Scheme;
    /// <summary>Canonical value. This property may contain sensitive data.</summary>
    public string Normalized => Validation.Normalized!;
    /// <summary>Validation details including the check outcome and provenance.</summary>
    public EntityValidationResult Validation { get; }
    /// <summary>Outcome of the applied check digit.</summary>
    public EntityCheckOutcome CheckDigit => Validation.CheckDigit;
    /// <summary>Checksum applied to this variant, if any.</summary>
    public EntityCheckInfo? Check => Validation.Check;
    /// <summary>Publication status of the applied check.</summary>
    public EntityCheckProvenance CheckProvenance => Validation.CheckProvenance;
    /// <summary>Confidence assigned from the evidence, from zero to one.</summary>
    public double Confidence { get; }
    /// <summary>Evidence used to score this reading.</summary>
    public EntityEvidence Evidence { get; }
    /// <summary>The registered keyword that supported this reading, if any.</summary>
    public string? ContextKeyword { get; }
    /// <summary>BCP 47 language of the context keyword, if any.</summary>
    public string? ContextLanguage { get; }
    /// <summary>Zero-based index of the first source segment in this find result.</summary>
    public int SegmentIndex { get; }
    /// <summary>Ordered portions of the match, each using coordinates from its own input segment.</summary>
    public IReadOnlyList<EntityMatchPart> Parts { get; }
    /// <summary>Coordinates supplied by the host.</summary>
    public EntitySegmentLocation? Location { get; }
    /// <summary>Competing readings of the same span.</summary>
    public IReadOnlyList<EntityMatch> Alternatives { get; }
    /// <summary>Readings of identifiers declared to be contained inside this span.</summary>
    public IReadOnlyList<EntityMatch> Contained { get; }

    // Preserve losing readings with distinct sensitivity or coverage for masking.
    internal IReadOnlyList<EntityMatch> MaskingFallbacks { get; }

    /// <inheritdoc />
    public override string ToString() => string.Format(CultureInfo.InvariantCulture, "{0}/{1} [{2}..{3}] {4:0.00}",
        Kind, Scheme.Id, Start, Start + Length, Confidence);
}
