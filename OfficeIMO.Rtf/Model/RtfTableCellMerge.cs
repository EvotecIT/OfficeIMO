namespace OfficeIMO.Rtf;

/// <summary>
/// Merge marker for RTF table cells.
/// </summary>
public enum RtfTableCellMerge {
    /// <summary>No merge marker is present.</summary>
    None,

    /// <summary>The cell starts a merged range.</summary>
    First,

    /// <summary>The cell continues a merged range.</summary>
    Continue
}
