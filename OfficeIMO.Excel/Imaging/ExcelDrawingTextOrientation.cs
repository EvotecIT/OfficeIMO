namespace OfficeIMO.Excel;

/// <summary>
/// Text orientation requested by an Excel worksheet drawing object's DrawingML body properties.
/// </summary>
public enum ExcelDrawingTextOrientation {
    /// <summary>Horizontal text layout.</summary>
    Horizontal,

    /// <summary>Vertical text layout.</summary>
    Vertical,

    /// <summary>Vertical text layout rotated 270 degrees.</summary>
    Vertical270,

    /// <summary>East Asian vertical text layout.</summary>
    EastAsianVertical,

    /// <summary>Mongolian vertical text layout.</summary>
    MongolianVertical,

    /// <summary>WordArt vertical text layout.</summary>
    WordArtVertical,

    /// <summary>WordArt left-to-right vertical text layout.</summary>
    WordArtLeftToRight,

    /// <summary>Text orientation could not be mapped to a known DrawingML value.</summary>
    Unknown
}
