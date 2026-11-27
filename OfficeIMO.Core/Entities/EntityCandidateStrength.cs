namespace OfficeIMO.Entities;

/// <summary>Evidence inherent in a candidate before considering formatting and context.</summary>
public enum EntityCandidateStrength {
    /// <summary>Format alone provides weak evidence.</summary>
    ShapeOnly,
    /// <summary>A checksum over otherwise ordinary digits passed.</summary>
    DigitCheck,
    /// <summary>A country prefix distinguishes a scheme without a checksum.</summary>
    PrefixedShape,
    /// <summary>A distinctive format provides evidence without a checksum.</summary>
    DistinctiveShape,
    /// <summary>A checksum passed on a distinctive format or prefixed identifier.</summary>
    StrongCheck,
    /// <summary>A checksum over a short letter and digit shape passed.</summary>
    ShortLetterDigitCheck,
}
