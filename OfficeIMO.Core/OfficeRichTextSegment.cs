namespace OfficeIMO.Drawing;

/// <summary>
/// Describes one measured rich text segment on a laid-out line.
/// </summary>
public sealed class OfficeRichTextSegment {
    /// <summary>Automatic pair-kerning threshold copied from the source run. Null retains normal font-feature defaults; zero disables pair kerning.</summary>
    public double? AutomaticPairKerningThreshold { get; internal set; }
    internal OfficeTextFeatureSettings? KerningFeatures => OfficeTextKerning.Features(AutomaticPairKerningThreshold, FontSize);
    internal OfficePreparedTextShaping? PreparedShaping { get; set; }
    /// <summary>Optional language identifier retained from the source run, including a generated hyphen painted with that run's style.</summary>
    public uint? LanguageId { get; internal set; }
    // Native document adapters keep annotations attached through wrapping and shrinking.
    internal string? LinkUri { get; set; }
    internal string? LinkDestinationName { get; set; }
    internal OfficeTextTabLineLeaderPaint? TabLinePaint { get; set; }
    internal OfficeFontStyle FontStyle => (Bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
        | (Italic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);

    internal OfficeRichTextSegment WithLayout(string text, double width, OfficePreparedTextShaping? prepared) =>
        new OfficeRichTextSegment(text, width, FontSize, Color, Bold, Italic, Underline, FontFamily,
            Strikethrough, BackgroundColor, UnderlineStyle, StrikethroughStyle, Baseline) {
            AutomaticPairKerningThreshold = AutomaticPairKerningThreshold, LanguageId = LanguageId,
            LinkUri = LinkUri, LinkDestinationName = LinkDestinationName, PreparedShaping = prepared,
            TabLinePaint = text == Text ? TabLinePaint : null
        };

    /// <summary>Creates a measured segment using the pre-typography constructor signature.</summary>
    public OfficeRichTextSegment(string text, double width, double fontSize, OfficeColor color, bool bold, bool italic, bool underline, string fontFamily, bool strikethrough, OfficeColor? backgroundColor)
        : this(text, width, fontSize, color, bold, italic, underline, fontFamily, strikethrough, backgroundColor,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, OfficeTextBaseline.Normal) {
    }

    /// <summary>
    /// Creates a measured rich text segment.
    /// </summary>
    public OfficeRichTextSegment(string text, double width, double fontSize, OfficeColor color, bool bold, bool italic, bool underline, string fontFamily, bool strikethrough = false, OfficeColor? backgroundColor = null, OfficeTextDecorationStyle underlineStyle = OfficeTextDecorationStyle.None, OfficeTextDecorationStyle strikethroughStyle = OfficeTextDecorationStyle.None, OfficeTextBaseline baseline = OfficeTextBaseline.Normal) {
        if (underlineStyle < OfficeTextDecorationStyle.None || underlineStyle > OfficeTextDecorationStyle.Words) {
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
