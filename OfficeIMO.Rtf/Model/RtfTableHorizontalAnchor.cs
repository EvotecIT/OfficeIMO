namespace OfficeIMO.Rtf;

/// <summary>
/// Horizontal reference frame for a positioned RTF table row.
/// </summary>
public enum RtfTableHorizontalAnchor {
    /// <summary>Use the containing column as the horizontal frame, represented by <c>\tphcol</c>.</summary>
    Column,

    /// <summary>Use the page margin as the horizontal frame, represented by <c>\tphmrg</c>.</summary>
    Margin,

    /// <summary>Use the page as the horizontal frame, represented by <c>\tphpg</c>.</summary>
    Page
}
