namespace OfficeIMO.Rtf;

/// <summary>
/// Text flow controls used by RTF table-cell definitions.
/// </summary>
public enum RtfTableCellTextFlow {
    /// <summary>Left-to-right lines laid out from top to bottom.</summary>
    LeftToRightTopToBottom,

    /// <summary>Top-to-bottom text laid out from right to left.</summary>
    TopToBottomRightToLeft,

    /// <summary>Bottom-to-top text laid out from left to right.</summary>
    BottomToTopLeftToRight,

    /// <summary>Vertical left-to-right, top-to-bottom text flow.</summary>
    LeftToRightTopToBottomVertical,

    /// <summary>Vertical top-to-bottom, right-to-left text flow.</summary>
    TopToBottomRightToLeftVertical
}
