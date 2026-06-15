namespace OfficeIMO.Rtf;

/// <summary>
/// RTF font family classification.
/// </summary>
public enum RtfFontFamily {
    /// <summary>Unknown or default family, represented by <c>\fnil</c>.</summary>
    Nil,

    /// <summary>Roman serif font family.</summary>
    Roman,

    /// <summary>Swiss sans-serif font family.</summary>
    Swiss,

    /// <summary>Fixed-pitch modern font family.</summary>
    Modern,

    /// <summary>Script font family.</summary>
    Script,

    /// <summary>Decorative font family.</summary>
    Decorative,

    /// <summary>Technical or symbol font family.</summary>
    Technical,

    /// <summary>Bidirectional font family.</summary>
    Bidirectional
}
