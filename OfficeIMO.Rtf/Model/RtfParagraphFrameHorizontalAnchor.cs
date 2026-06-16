namespace OfficeIMO.Rtf;

/// <summary>
/// Horizontal reference frame for positioned RTF paragraphs.
/// </summary>
public enum RtfParagraphFrameHorizontalAnchor {
    /// <summary>Use the containing column as the horizontal frame, represented by <c>\phcol</c>.</summary>
    Column,

    /// <summary>Use the page margin as the horizontal frame, represented by <c>\phmrg</c>.</summary>
    Margin,

    /// <summary>Use the page as the horizontal frame, represented by <c>\phpg</c>.</summary>
    Page
}
