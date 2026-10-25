namespace OfficeIMO.Pdf;

/// <summary>
/// Text alignment or a stationary rule declared at a paragraph tab stop.
/// </summary>
public enum PdfTabAlignment {
    /// <summary>The following text starts at the tab stop.</summary>
    Left = 0,
    /// <summary>The following text ends at the tab stop.</summary>
    Right = 1,
    /// <summary>The following text is centered on the tab stop.</summary>
    Center = 2,
    /// <summary>The following text's decimal separator is aligned to the tab stop.</summary>
    DecimalSeparator = 3,
    /// <summary>Draws a stationary vertical rule on each paragraph line without positioning text.</summary>
    /// <remarks>Bar declarations belong to paragraph styles, not explicit tab runs.</remarks>
    Bar = 4
}
