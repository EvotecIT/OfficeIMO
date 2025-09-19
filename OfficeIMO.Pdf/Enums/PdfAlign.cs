namespace OfficeIMO.Pdf;

/// <summary>Text alignment within the content area.</summary>
public enum PdfAlign {
    /// <summary>Align text to the left.</summary>
    Left,
    /// <summary>Center text within the content area.</summary>
    Center,
    /// <summary>Align text to the right.</summary>
    Right,
    /// <summary>Distribute extra space between words to fill the line width (last line not justified).</summary>
    Justify
}
