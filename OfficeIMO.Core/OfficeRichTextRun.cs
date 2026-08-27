namespace OfficeIMO.Drawing;

/// <summary>
/// Describes one styled text run for shared rich text layout and rendering.
/// </summary>
public sealed class OfficeRichTextRun {
    /// <summary>
    /// Creates a styled text run.
    /// </summary>
    /// <param name="text">Run text.</param>
    /// <param name="fontSize">Font size used for measurement and rendering.</param>
    /// <param name="color">Text color.</param>
    /// <param name="bold">Whether the run should render as bold.</param>
    /// <param name="italic">Whether the run should render as italic.</param>
    /// <param name="underline">Whether the run should render with underline.</param>
    /// <param name="fontFamily">Preferred font family for SVG or future font-aware renderers.</param>
    /// <param name="strikethrough">Whether the run should render with strikethrough.</param>
    /// <param name="backgroundColor">Optional run background/highlight color.</param>
    /// <param name="underlineStyle">Underline pattern. A non-none value takes precedence over <paramref name="underline"/>.</param>
    /// <param name="strikethroughStyle">Strikethrough pattern. A non-none value takes precedence over <paramref name="strikethrough"/>.</param>
    /// <param name="baseline">Vertical baseline placement.</param>
    public OfficeRichTextRun(string? text, double fontSize, OfficeColor color, bool bold = false, bool italic = false, bool underline = false, string? fontFamily = null, bool strikethrough = false, OfficeColor? backgroundColor = null, OfficeTextDecorationStyle underlineStyle = OfficeTextDecorationStyle.None, OfficeTextDecorationStyle strikethroughStyle = OfficeTextDecorationStyle.None, OfficeTextBaseline baseline = OfficeTextBaseline.Normal) {
        if (underlineStyle < OfficeTextDecorationStyle.None || underlineStyle > OfficeTextDecorationStyle.Wavy) {
            throw new System.ArgumentOutOfRangeException(nameof(underlineStyle));
        }
        if (strikethroughStyle < OfficeTextDecorationStyle.None || strikethroughStyle > OfficeTextDecorationStyle.Wavy) {
            throw new System.ArgumentOutOfRangeException(nameof(strikethroughStyle));
        }
        if (baseline < OfficeTextBaseline.Normal || baseline > OfficeTextBaseline.Subscript) {
            throw new System.ArgumentOutOfRangeException(nameof(baseline));
        }
        Text = text ?? string.Empty;
        FontSize = fontSize;
        Color = color;
        Bold = bold;
        Italic = italic;
        UnderlineStyle = underlineStyle != OfficeTextDecorationStyle.None
            ? underlineStyle
            : underline ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None;
        StrikethroughStyle = strikethroughStyle != OfficeTextDecorationStyle.None
            ? strikethroughStyle
            : strikethrough ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None;
        Baseline = baseline;
        FontFamily = string.IsNullOrWhiteSpace(fontFamily) ? "Arial, sans-serif" : fontFamily!;
        BackgroundColor = backgroundColor;
    }

    /// <summary>
    /// Gets the run text.
    /// </summary>
    public string Text { get; }

    /// <summary>
    /// Gets the run font size.
    /// </summary>
    public double FontSize { get; }

    /// <summary>
    /// Gets the run text color.
    /// </summary>
    public OfficeColor Color { get; }

    /// <summary>
    /// Gets whether the run should render as bold.
    /// </summary>
    public bool Bold { get; }

    /// <summary>
    /// Gets whether the run should render as italic.
    /// </summary>
    public bool Italic { get; }

    /// <summary>
    /// Gets whether the run should render with underline.
    /// </summary>
    public bool Underline => UnderlineStyle != OfficeTextDecorationStyle.None;

    /// <summary>Gets the underline line pattern.</summary>
    public OfficeTextDecorationStyle UnderlineStyle { get; }

    /// <summary>
    /// Gets whether the run should render with strikethrough.
    /// </summary>
    public bool Strikethrough => StrikethroughStyle != OfficeTextDecorationStyle.None;

    /// <summary>Gets the strikethrough line pattern.</summary>
    public OfficeTextDecorationStyle StrikethroughStyle { get; }

    /// <summary>Gets the vertical baseline placement.</summary>
    public OfficeTextBaseline Baseline { get; }

    /// <summary>Gets the effective font size used for measuring and rendering this run.</summary>
    public double EffectiveFontSize => Baseline == OfficeTextBaseline.Normal ? FontSize : FontSize * 0.65D;

    /// <summary>
    /// Gets the preferred font family for SVG or future font-aware renderers.
    /// </summary>
    public string FontFamily { get; }

    /// <summary>
    /// Gets the optional run background/highlight color.
    /// </summary>
    public OfficeColor? BackgroundColor { get; }

    /// <summary>Creates a copy with transformed text casing while preserving all drawing styles.</summary>
    public OfficeRichTextRun WithTextCase(OfficeTextCase textCase, System.Globalization.CultureInfo? culture = null) =>
        new OfficeRichTextRun(
            OfficeTextCaseTransformer.Apply(Text, textCase, culture),
            FontSize,
            Color,
            Bold,
            Italic,
            Underline,
            FontFamily,
            Strikethrough,
            BackgroundColor,
            UnderlineStyle,
            StrikethroughStyle,
            Baseline);
}
