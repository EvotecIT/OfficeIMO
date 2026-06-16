namespace OfficeIMO.Rtf;

/// <summary>
/// Footnote and endnote numbering restart behavior.
/// </summary>
public enum RtfNoteNumberRestart {
    /// <summary>Number notes continuously through the document or section.</summary>
    Continuous,

    /// <summary>Restart note numbering for each section.</summary>
    EachSection,

    /// <summary>Restart note numbering on each page. RTF defines this for footnotes.</summary>
    EachPage
}
