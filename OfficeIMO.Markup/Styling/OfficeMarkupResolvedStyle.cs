using OfficeIMO.Drawing;

namespace OfficeIMO.Markup;

/// <summary>Optional visual style values resolved from a theme, named style, and block attributes.</summary>
public sealed class OfficeMarkupResolvedStyle {
    /// <summary>Gets or sets the style name, when one was selected.</summary>
    public string? Name { get; set; }
    /// <summary>Gets or sets the font family.</summary>
    public string? FontName { get; set; }
    /// <summary>Gets or sets the numeric font size.</summary>
    public int? FontSize { get; set; }
    /// <summary>Gets or sets the optional bold setting.</summary>
    public bool? Bold { get; set; }
    /// <summary>Gets or sets the optional italic setting.</summary>
    public bool? Italic { get; set; }
    /// <summary>Gets or sets the underline style, including an explicit <c>None</c>.</summary>
    public OfficeTextDecorationStyle? UnderlineStyle { get; set; }
    /// <summary>Gets or sets the strikethrough style, including an explicit <c>None</c>.</summary>
    public OfficeTextDecorationStyle? StrikethroughStyle { get; set; }
    /// <summary>Gets or sets the text baseline, such as superscript or subscript.</summary>
    public OfficeTextBaseline? Baseline { get; set; }
    /// <summary>Gets or sets the text-case transformation.</summary>
    public OfficeTextCase? TextCase { get; set; }
    /// <summary>Gets or sets the optional small-caps setting.</summary>
    public bool? SmallCaps { get; set; }
    /// <summary>Gets or sets the text color expression.</summary>
    public string? TextColor { get; set; }
    /// <summary>Gets or sets the highlight color expression.</summary>
    public string? HighlightColor { get; set; }
    /// <summary>Gets or sets the fill color expression.</summary>
    public string? FillColor { get; set; }
    /// <summary>Gets or sets the border color expression.</summary>
    public string? BorderColor { get; set; }
    /// <summary>Gets or sets the text-alignment expression.</summary>
    public string? TextAlign { get; set; }

    /// <summary>Gets whether any visual value is set; a name alone does not count.</summary>
    public bool HasVisualValues =>
        !string.IsNullOrWhiteSpace(FontName)
        || FontSize != null
        || Bold != null
        || Italic != null
        || UnderlineStyle != null
        || StrikethroughStyle != null
        || Baseline != null
        || TextCase != null
        || SmallCaps != null
        || !string.IsNullOrWhiteSpace(TextColor)
        || !string.IsNullOrWhiteSpace(HighlightColor)
        || !string.IsNullOrWhiteSpace(FillColor)
        || !string.IsNullOrWhiteSpace(BorderColor)
        || !string.IsNullOrWhiteSpace(TextAlign);

    internal OfficeMarkupResolvedStyle Clone() =>
        new OfficeMarkupResolvedStyle {
            Name = Name,
            FontName = FontName,
            FontSize = FontSize,
            Bold = Bold,
            Italic = Italic,
            UnderlineStyle = UnderlineStyle,
            StrikethroughStyle = StrikethroughStyle,
            Baseline = Baseline,
            TextCase = TextCase,
            SmallCaps = SmallCaps,
            TextColor = TextColor,
            HighlightColor = HighlightColor,
            FillColor = FillColor,
            BorderColor = BorderColor,
            TextAlign = TextAlign
        };
}
