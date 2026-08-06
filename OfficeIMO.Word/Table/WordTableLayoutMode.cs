namespace OfficeIMO.Word;

/// <summary>
/// Specifies the layout algorithm Word uses for a table.
/// </summary>
public enum WordTableLayoutMode {
    /// <summary>
    /// Allows Word to adjust columns based on their contents.
    /// </summary>
    AutoFit,

    /// <summary>
    /// Uses the table and cell preferred widths without content-driven column resizing.
    /// </summary>
    Fixed
}
