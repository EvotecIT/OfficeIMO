namespace OfficeIMO.Pdf;

/// <summary>
/// Leader fill rendered across a paragraph tab advance.
/// </summary>
public enum PdfTabLeaderStyle {
    /// <summary>No leader fill; the tab advances by whitespace only.</summary>
    None = 0,
    /// <summary>Fill the tab advance with dot leaders.</summary>
    Dots = 1
}
