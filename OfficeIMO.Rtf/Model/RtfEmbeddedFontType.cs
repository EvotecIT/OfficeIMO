namespace OfficeIMO.Rtf;

/// <summary>
/// Embedded font type from an RTF <c>{\*\fontemb ...}</c> destination.
/// </summary>
public enum RtfEmbeddedFontType {
    /// <summary>Unknown or default embedded font type.</summary>
    Unknown,

    /// <summary>TrueType embedded font.</summary>
    TrueType
}
