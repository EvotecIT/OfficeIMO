namespace OfficeIMO.Markup;

public sealed class OfficeMarkupResolvedStyle {
    public string? Name { get; set; }
    public string? FontName { get; set; }
    public int? FontSize { get; set; }
    public bool? Bold { get; set; }
    public bool? Italic { get; set; }
    public string? TextColor { get; set; }
    public string? FillColor { get; set; }
    public string? BorderColor { get; set; }
    public string? TextAlign { get; set; }

    public bool HasVisualValues =>
        !string.IsNullOrWhiteSpace(FontName)
        || FontSize != null
        || Bold != null
        || Italic != null
        || !string.IsNullOrWhiteSpace(TextColor)
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
            TextColor = TextColor,
            FillColor = FillColor,
            BorderColor = BorderColor,
            TextAlign = TextAlign
        };
}
