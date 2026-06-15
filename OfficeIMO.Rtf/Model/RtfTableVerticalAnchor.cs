namespace OfficeIMO.Rtf;

/// <summary>
/// Vertical reference frame for a positioned RTF table row.
/// </summary>
public enum RtfTableVerticalAnchor {
    /// <summary>Use the page margin as the vertical frame, represented by <c>\tpvmrg</c>.</summary>
    Margin,

    /// <summary>Use the paragraph as the vertical frame, represented by <c>\tpvpara</c>.</summary>
    Paragraph,

    /// <summary>Use the page as the vertical frame, represented by <c>\tpvpg</c>.</summary>
    Page
}
