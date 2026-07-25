namespace OfficeIMO.Drawing;

/// <summary>Known container formats for reusable OpenType font data.</summary>
public enum OfficeFontContainerFormat {
    /// <summary>The byte sequence is not a recognized font container.</summary>
    Unknown,
    /// <summary>A direct TrueType, OpenType, or TrueType Collection sfnt container.</summary>
    OpenType,
    /// <summary>A Web Open Font Format 1 container.</summary>
    Woff,
    /// <summary>A Web Open Font Format 2 container.</summary>
    Woff2
}
