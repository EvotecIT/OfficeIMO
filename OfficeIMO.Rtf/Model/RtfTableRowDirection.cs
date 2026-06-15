namespace OfficeIMO.Rtf;

/// <summary>
/// Reading order for cells in an RTF table row.
/// </summary>
public enum RtfTableRowDirection {
    /// <summary>Cells have left-to-right precedence, represented by <c>\ltrrow</c>.</summary>
    LeftToRight,

    /// <summary>Cells have right-to-left precedence, represented by <c>\rtlrow</c>.</summary>
    RightToLeft
}
