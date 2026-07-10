namespace OfficeIMO.Rtf;

/// <summary>Action taken for a source feature during conversion.</summary>
public enum RtfConversionAction {
    /// <summary>The feature was preserved.</summary>
    Preserved,

    /// <summary>The feature was represented with reduced structure or fidelity.</summary>
    Flattened,

    /// <summary>The feature was not written to the destination.</summary>
    Omitted,

    /// <summary>The feature or resource was intentionally rejected by policy.</summary>
    Blocked,

    /// <summary>The destination used a replacement representation.</summary>
    Substituted
}
