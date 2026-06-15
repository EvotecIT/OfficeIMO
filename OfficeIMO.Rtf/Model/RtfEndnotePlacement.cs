namespace OfficeIMO.Rtf;

/// <summary>
/// Placement for RTF endnotes.
/// </summary>
public enum RtfEndnotePlacement {
    /// <summary>Endnotes are placed at the end of the section.</summary>
    SectionEnd,

    /// <summary>Endnotes are placed at the end of the document.</summary>
    DocumentEnd,

    /// <summary>Endnotes are placed at the bottom of the page.</summary>
    PageBottom,

    /// <summary>Endnotes are placed beneath the body text.</summary>
    BeneathText
}
