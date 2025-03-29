namespace OfficeIMO.Word;

/// <summary>
/// Defines the layout type options for the table
/// </summary>
public enum WordTableLayoutType {
    /// <summary>
    /// AutoFit to Contents: Table width will be determined by content
    /// </summary>
    AutoFitToContents,

    /// <summary>
    /// AutoFit to Window: Table width will be 100% of the window
    /// </summary>
    AutoFitToWindow,

    /// <summary>
    /// Fixed Width: Table width will be set to a specific percentage
    /// </summary>
    FixedWidth
}