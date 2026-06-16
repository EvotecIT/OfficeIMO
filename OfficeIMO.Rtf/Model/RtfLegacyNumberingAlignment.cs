namespace OfficeIMO.Rtf;

/// <summary>
/// Alignment for Word 6/95 legacy RTF paragraph numbering.
/// </summary>
public enum RtfLegacyNumberingAlignment {
    /// <summary>Left-justified numbering represented by <c>\pnql</c>.</summary>
    Left,

    /// <summary>Centered numbering represented by <c>\pnqc</c>.</summary>
    Center,

    /// <summary>Right-justified numbering represented by <c>\pnqr</c>.</summary>
    Right
}
