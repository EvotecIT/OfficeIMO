namespace OfficeIMO.Rtf;

/// <summary>
/// Vertical reference frame for positioned RTF paragraphs.
/// </summary>
public enum RtfParagraphFrameVerticalAnchor {
    /// <summary>Use the page margin as the vertical frame, represented by <c>\pvmrg</c>.</summary>
    Margin,

    /// <summary>Use the following unframed paragraph as the vertical frame, represented by <c>\pvpara</c>.</summary>
    Paragraph,

    /// <summary>Use the page as the vertical frame, represented by <c>\pvpg</c>.</summary>
    Page
}
