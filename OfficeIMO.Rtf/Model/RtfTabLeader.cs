namespace OfficeIMO.Rtf;

/// <summary>
/// Leader style used by an RTF paragraph tab stop.
/// </summary>
public enum RtfTabLeader {
    /// <summary>No leader characters.</summary>
    None,

    /// <summary>Dotted leader.</summary>
    Dots,

    /// <summary>Centered dot leader.</summary>
    MiddleDots,

    /// <summary>Hyphen leader.</summary>
    Hyphen,

    /// <summary>Underline leader.</summary>
    Underline,

    /// <summary>Thick underline leader.</summary>
    ThickLine,

    /// <summary>Equal-sign leader.</summary>
    EqualSign
}
