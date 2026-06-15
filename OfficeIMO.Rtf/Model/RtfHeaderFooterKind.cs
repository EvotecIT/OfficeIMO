namespace OfficeIMO.Rtf;

/// <summary>
/// Header and footer destinations supported by RTF.
/// </summary>
public enum RtfHeaderFooterKind {
    /// <summary>Default header for pages that do not use first, left, or right variants.</summary>
    Header,

    /// <summary>Left-page header.</summary>
    LeftHeader,

    /// <summary>Right-page header.</summary>
    RightHeader,

    /// <summary>First-page header.</summary>
    FirstHeader,

    /// <summary>Default footer for pages that do not use first, left, or right variants.</summary>
    Footer,

    /// <summary>Left-page footer.</summary>
    LeftFooter,

    /// <summary>Right-page footer.</summary>
    RightFooter,

    /// <summary>First-page footer.</summary>
    FirstFooter
}
