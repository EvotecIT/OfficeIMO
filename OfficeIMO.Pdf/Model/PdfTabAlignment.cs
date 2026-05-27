namespace OfficeIMO.Pdf;

/// <summary>
/// Alignment used when text is positioned at a paragraph tab stop.
/// </summary>
public enum PdfTabAlignment {
    /// <summary>The following text starts at the tab stop.</summary>
    Left = 0,
    /// <summary>The following text is centered on the tab stop.</summary>
    Center = 1,
    /// <summary>The following text ends at the tab stop.</summary>
    Right = 2,
    /// <summary>The following text's decimal separator is aligned to the tab stop.</summary>
    DecimalSeparator = 3
}
