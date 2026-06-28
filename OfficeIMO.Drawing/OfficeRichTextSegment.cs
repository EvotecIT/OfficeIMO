namespace OfficeIMO.Drawing;

/// <summary>
/// Describes one measured rich text segment on a laid-out line.
/// </summary>
public sealed class OfficeRichTextSegment {
    /// <summary>
    /// Creates a measured rich text segment.
    /// </summary>
    public OfficeRichTextSegment(string text, double width, double fontSize, OfficeColor color, bool bold, bool italic, bool underline, string fontFamily, bool strikethrough = false, OfficeColor? backgroundColor = null) {
        Text = text;
        Width = width;
        FontSize = fontSize;
        Color = color;
        Bold = bold;
        Italic = italic;
        Underline = underline;
        Strikethrough = strikethrough;
        FontFamily = fontFamily;
        BackgroundColor = backgroundColor;
    }

    /// <summary>
    /// Gets the segment text.
    /// </summary>
    public string Text { get; }

    /// <summary>
    /// Gets the measured segment width.
    /// </summary>
    public double Width { get; }

    /// <summary>
    /// Gets the segment font size.
    /// </summary>
    public double FontSize { get; }

    /// <summary>
    /// Gets the segment color.
    /// </summary>
    public OfficeColor Color { get; }

    /// <summary>
    /// Gets whether the segment should render as bold.
    /// </summary>
    public bool Bold { get; }

    /// <summary>
    /// Gets whether the segment should render as italic.
    /// </summary>
    public bool Italic { get; }

    /// <summary>
    /// Gets whether the segment should render with underline.
    /// </summary>
    public bool Underline { get; }

    /// <summary>
    /// Gets whether the segment should render with strikethrough.
    /// </summary>
    public bool Strikethrough { get; }

    /// <summary>
    /// Gets the segment font family.
    /// </summary>
    public string FontFamily { get; }

    /// <summary>
    /// Gets the optional segment background/highlight color.
    /// </summary>
    public OfficeColor? BackgroundColor { get; }
}
