namespace OfficeIMO.Rtf;

/// <summary>
/// Placement for RTF footnotes.
/// </summary>
public enum RtfFootnotePlacement {
    /// <summary>Footnotes are placed at the bottom of the page.</summary>
    PageBottom,

    /// <summary>Footnotes are placed beneath the body text.</summary>
    BeneathText,

    /// <summary>Footnotes are placed at the end of the section.</summary>
    SectionEnd,

    /// <summary>Footnotes are placed at the end of the document.</summary>
    DocumentEnd
}
