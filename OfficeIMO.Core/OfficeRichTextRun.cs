namespace OfficeIMO.Drawing;

/// <summary>
/// Describes one styled text run for shared rich text layout and rendering.
/// </summary>
public sealed class OfficeRichTextRun {
    private double? _automaticPairKerningThreshold;
    /// <summary>Native automatic pair-kerning threshold, in points. Zero disables pair kerning; null leaves the native setting unspecified.</summary>
    /// <remarks>Font-aware drawing layout and raster, SVG and PDF rendering enable pair kerning when the base run size meets a positive threshold. Zero disables it; null retains the renderer's normal font-feature defaults. Custom width callbacks retain control of measurement.</remarks>
    public double? AutomaticPairKerningThreshold {
        get => _automaticPairKerningThreshold;
        internal set {
            if (value.HasValue && (double.IsNaN(value.Value) || double.IsInfinity(value.Value) || value.Value < 0))
                throw new System.ArgumentOutOfRangeException(nameof(value), "A pair-kerning threshold must be finite and nonnegative.");
            _automaticPairKerningThreshold = value;
        }
    }
    internal OfficeTextFeatureSettings? KerningFeatures => OfficeTextKerning.Features(AutomaticPairKerningThreshold, FontSize);

    /// <summary>Returns an independent run with a native pair-kerning threshold. Null leaves the setting unspecified; zero explicitly disables it.</summary>
    /// <param name="thresholdPoints">A finite, nonnegative point size, or null. A native writer may impose a narrower representable range.</param>
    public OfficeRichTextRun WithAutomaticPairKerningThreshold(double? thresholdPoints) {
        OfficeRichTextRun copy = WithLanguageId(LanguageId);
        copy.AutomaticPairKerningThreshold = thresholdPoints;
        return copy;
    }
    // Native document adapters keep annotations attached through wrapping and shrinking.
    internal string? LinkUri { get; set; }
    internal string? LinkDestinationName { get; set; }
    internal OfficeFontStyle FontStyle => (Bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
        | (Italic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);

    /// <summary>Creates a styled text run using the pre-typography constructor signature.</summary>
    public OfficeRichTextRun(string? text, double fontSize, OfficeColor color, bool bold, bool italic, bool underline, string? fontFamily, bool strikethrough, OfficeColor? backgroundColor)
        : this(text, fontSize, color, bold, italic, underline, fontFamily, strikethrough, backgroundColor,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, OfficeTextBaseline.Normal) {
    }

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
    public OfficeRichTextRun(string? text, double fontSize, OfficeColor color, bool bold = false, bool italic = false, bool underline = false, string? fontFamily = null, bool strikethrough = false, OfficeColor? backgroundColor = null, OfficeTextDecorationStyle underlineStyle = OfficeTextDecorationStyle.None, OfficeTextDecorationStyle strikethroughStyle = OfficeTextDecorationStyle.None, OfficeTextBaseline baseline = OfficeTextBaseline.Normal)
        : this(text, fontSize, color, bold, italic, underline, fontFamily, strikethrough,
            backgroundColor, underlineStyle, strikethroughStyle, baseline, paragraphIndent: null) {
    }

    /// <summary>Creates a styled text run with an optional per-paragraph indentation override.</summary>
    /// <param name="text">Run text.</param>
    /// <param name="fontSize">Font size used for measurement and rendering.</param>
    /// <param name="color">Text color.</param>
    /// <param name="bold">Whether the run should render as bold.</param>
    /// <param name="italic">Whether the run should render as italic.</param>
    /// <param name="underline">Whether the run should render with underline.</param>
    /// <param name="fontFamily">Preferred font family.</param>
    /// <param name="strikethrough">Whether the run should render with strikethrough.</param>
    /// <param name="backgroundColor">Optional run background color.</param>
    /// <param name="underlineStyle">Underline pattern.</param>
    /// <param name="strikethroughStyle">Strikethrough pattern.</param>
    /// <param name="baseline">Vertical baseline placement.</param>
    /// <param name="paragraphIndent">Indentation for a paragraph beginning at this run or after its line break.</param>
    public OfficeRichTextRun(string? text, double fontSize, OfficeColor color, bool bold, bool italic, bool underline, string? fontFamily, bool strikethrough, OfficeColor? backgroundColor, OfficeTextDecorationStyle underlineStyle, OfficeTextDecorationStyle strikethroughStyle, OfficeTextBaseline baseline, OfficeTextParagraphIndent? paragraphIndent) {
        if (underlineStyle < OfficeTextDecorationStyle.None || underlineStyle > OfficeTextDecorationStyle.Words) {
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
        ParagraphIndent = paragraphIndent;
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

    /// <summary>Optional first-line and continuation offsets for the paragraph beginning at this run.</summary>
    public OfficeTextParagraphIndent? ParagraphIndent { get; }

    /// <summary>Optional Windows language identifier for native document adapters. Null leaves language unspecified; it does not request removal of an imported override. Shared layout does not perform language-specific hyphenation.</summary>
    public uint? LanguageId { get; internal set; }

    /// <summary>Creates a copy with an explicit 16-bit language identifier, or an unspecified language when null. All text, styling and annotations are retained.</summary>
    public OfficeRichTextRun WithLanguageId(uint? languageId) {
        if (languageId > ushort.MaxValue) throw new System.ArgumentOutOfRangeException(nameof(languageId));
        return new OfficeRichTextRun(Text, FontSize, Color, Bold, Italic, Underline, FontFamily, Strikethrough,
            BackgroundColor, UnderlineStyle, StrikethroughStyle, Baseline, ParagraphIndent) {
            AutomaticPairKerningThreshold = AutomaticPairKerningThreshold, LanguageId = languageId, LinkUri = LinkUri, LinkDestinationName = LinkDestinationName
        };
    }

    /// <summary>Creates a styled copy that starts a paragraph with its own indentation.</summary>
    public OfficeRichTextRun WithParagraphIndent(OfficeTextParagraphIndent paragraphIndent) =>
        new OfficeRichTextRun(Text, FontSize, Color, Bold, Italic, Underline, FontFamily, Strikethrough,
            BackgroundColor, UnderlineStyle, StrikethroughStyle, Baseline, paragraphIndent) { AutomaticPairKerningThreshold = AutomaticPairKerningThreshold, LanguageId = LanguageId, LinkUri = LinkUri, LinkDestinationName = LinkDestinationName };

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
            Baseline,
            ParagraphIndent) { AutomaticPairKerningThreshold = AutomaticPairKerningThreshold, LanguageId = LanguageId, LinkUri = LinkUri, LinkDestinationName = LinkDestinationName };
}
