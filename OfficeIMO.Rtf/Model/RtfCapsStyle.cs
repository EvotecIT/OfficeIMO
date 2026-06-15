namespace OfficeIMO.Rtf;

/// <summary>
/// Character capitalization effect for an RTF run.
/// </summary>
public enum RtfCapsStyle {
    /// <summary>No capitalization effect.</summary>
    None,

    /// <summary>All caps effect using the RTF <c>\caps</c> control.</summary>
    Caps,

    /// <summary>Small caps effect using the RTF <c>\scaps</c> control.</summary>
    SmallCaps
}
