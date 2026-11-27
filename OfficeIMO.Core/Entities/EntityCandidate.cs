using System;

namespace OfficeIMO.Entities;

/// <summary>A recognizer's validated reading before confidence and overlap resolution.</summary>
public sealed class EntityCandidate {
    /// <summary>Creates a reading for a written span within the current segment.</summary>
    public EntityCandidate(EntityValidationResult validation, int start, int length, EntityCandidateStrength strength,
        EntityEvidence evidence = EntityEvidence.None) {
        Validation = validation ?? throw new ArgumentNullException(nameof(validation));
        if (start < 0) throw new ArgumentOutOfRangeException(nameof(start));
        if (length <= 0) throw new ArgumentOutOfRangeException(nameof(length));
        Start = start;
        Length = length;
        Strength = strength;
        Evidence = evidence;
    }

    /// <summary>Validation and normalization for the written value.</summary>
    public EntityValidationResult Validation { get; }
    /// <summary>UTF-16 offset in the segment.</summary>
    public int Start { get; }
    /// <summary>Written length in UTF-16 code units.</summary>
    public int Length { get; }
    /// <summary>Strength before contextual evidence.</summary>
    public EntityCandidateStrength Strength { get; }
    /// <summary>Recognizer-supplied evidence about the written form.</summary>
    public EntityEvidence Evidence { get; }

    internal bool RequiresContext { get; set; }
    internal bool OptionalPrefix { get; set; }
    internal bool RequiresPlausibleDate { get; set; }
}
