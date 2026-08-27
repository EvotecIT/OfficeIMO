namespace OfficeIMO.Drawing;

/// <summary>
/// Describes one measured rich text segment on a laid-out line.
/// </summary>
public sealed class OfficeRichTextSegment {
    /// <summary>Creates a measured segment using the pre-typography constructor signature.</summary>
    public OfficeRichTextSegment(string text, double width, double fontSize, OfficeColor color, bool bold, bool italic, bool underline, string fontFamily, bool strikethrough, OfficeColor? backgroundColor)
        : this(text, width, fontSize, color, bold, italic, underline, fontFamily, strikethrough, backgroundColor,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, OfficeTextBaseline.Normal) {
    }

    /// <summary>
    /// Creates a measured rich text segment.
    /// </summary>
    public OfficeRichTextSegment(string text, double width, double fontSize, OfficeColor color, bool bold, bool italic, bool underline, string fontFamily, bool strikethrough = false, OfficeColor? backgroundColor = null, OfficeTextDecorationStyle underlineStyle = OfficeTextDecorationStyle.None, OfficeTextDecorationStyle strikethroughStyle = OfficeTextDecorationStyle.None, OfficeTextBaseline baseline = OfficeTextBaseline.Normal) {
        if (underlineStyle < OfficeTextDecorationStyle.None || underlineStyle > OfficeTextDecorationStyle.Wavy) {
            throw new System.ArgumentOutOfRangeException(nameof(underlineStyle));
        }
        if (strikethroughStyle < OfficeTextDecorationStyle.None || strikethroughStyle > OfficeTextDecorationStyle.Wavy) {
            throw new System.ArgumentOutOfRangeException(nameof(strikethroughStyle));
        }
        if (baseline < OfficeTextBaseline.Normal || baseline > OfficeTextBaseline.Subscript) {
            throw new System.ArgumentOutOfRangeException(nameof(baseline));
        }
        Text = text;
        Width = width;
        FontSize = fontSize;
        Color = color;
        Bold = bold;
        Italic = italic;
        UnderlineStyle = underlineStyle != OfficeTextDecorationStyle.None ? underlineStyle : underline ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None;
        StrikethroughStyle = strikethroughStyle != OfficeTextDecorationStyle.None ? strikethroughStyle : strikethrough ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None;
        Baseline = baseline;
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
    public bool Underline => UnderlineStyle != OfficeTextDecorationStyle.None;

    /// <summary>Gets the underline line pattern.</summary>
    public OfficeTextDecorationStyle UnderlineStyle { get; }

    /// <summary>
    /// Gets whether the segment should render with strikethrough.
    /// </summary>
    public bool Strikethrough => StrikethroughStyle != OfficeTextDecorationStyle.None;

    /// <summary>Gets the strikethrough line pattern.</summary>
    public OfficeTextDecorationStyle StrikethroughStyle { get; }

    /// <summary>Gets the vertical baseline placement.</summary>
    public OfficeTextBaseline Baseline { get; }

    /// <summary>
    /// Gets the segment font family.
    /// </summary>
    public string FontFamily { get; }

    /// <summary>
    /// Gets the optional segment background/highlight color.
    /// </summary>
    public OfficeColor? BackgroundColor { get; }
}
