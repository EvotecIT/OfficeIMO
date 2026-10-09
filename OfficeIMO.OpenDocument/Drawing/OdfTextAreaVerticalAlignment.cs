namespace OfficeIMO.OpenDocument;

/// <summary>Native vertical placement of text inside an OpenDocument shape.</summary>
public enum OdfTextAreaVerticalAlignment {
    /// <summary>Aligns the text area with the top of the shape's text frame.</summary>
    Top,
    /// <summary>Centers the text area vertically in the shape's text frame.</summary>
    Middle,
    /// <summary>Aligns the text area with the bottom of the shape's text frame.</summary>
    Bottom,
    /// <summary>Distributes the text vertically across the frame; shared drawing projection does not reproduce this native mode.</summary>
    Justify
}
