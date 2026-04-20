namespace OfficeIMO.Markup;

public sealed class OfficeMarkupStyleResolver {
    private readonly Dictionary<string, OfficeMarkupResolvedStyle> _styles;

    private OfficeMarkupStyleResolver(Dictionary<string, OfficeMarkupResolvedStyle> styles) {
        _styles = styles;
    }

    public static OfficeMarkupStyleResolver Create(OfficeMarkupDocument document) {
        if (document == null) {
            throw new ArgumentNullException(nameof(document));
        }

        var styles = CreateBaseStyles(document.Metadata.TryGetValue("theme", out var theme) ? theme : null);
        return new OfficeMarkupStyleResolver(styles);
    }

    public OfficeMarkupResolvedStyle? Resolve(OfficeMarkupBlock block) {
        if (block == null) {
            return null;
        }

        switch (block) {
            case OfficeMarkupHeadingBlock heading:
                return Resolve("heading-" + Math.Min(Math.Max(heading.Level, 1), 3), block.Attributes);
            case OfficeMarkupParagraphBlock:
                return Resolve("body", block.Attributes);
            case OfficeMarkupTextBoxBlock textBox:
                return Resolve(textBox.Style, block.Attributes);
            case OfficeMarkupCardBlock card:
                return Resolve(string.IsNullOrWhiteSpace(card.Style) ? "card" : card.Style, block.Attributes);
            case OfficeMarkupExtensionBlock extension:
                return Resolve(GetAttribute(extension.Attributes, "style"), block.Attributes);
            default:
                return Resolve(GetAttribute(block.Attributes, "style"), block.Attributes);
        }
    }

    public OfficeMarkupResolvedStyle? Resolve(string? styleName, IDictionary<string, string>? attributes = null) {
        OfficeMarkupResolvedStyle? style = null;
        if (!string.IsNullOrWhiteSpace(styleName)) {
            if (_styles.TryGetValue(Normalize(styleName!), out var knownStyle)) {
                style = knownStyle.Clone();
            } else {
                style = new OfficeMarkupResolvedStyle {
                    Name = styleName!.Trim()
                };
            }
        }

        if (attributes != null) {
            style = ApplyAttributeOverrides(style, attributes);
        }

        return style != null && (style.HasVisualValues || !string.IsNullOrWhiteSpace(style.Name)) ? style : null;
    }

    private static Dictionary<string, OfficeMarkupResolvedStyle> CreateBaseStyles(string? themeName) {
        var palette = CreatePalette(themeName);
        var font = palette.FontName;
        var styles = new Dictionary<string, OfficeMarkupResolvedStyle>(StringComparer.OrdinalIgnoreCase);

        Add(styles, "heading-1", font, 28, true, palette.Text, null, null, null);
        Add(styles, "heading-2", font, 22, true, palette.Text, null, null, null);
        Add(styles, "heading-3", font, 17, true, palette.Text, null, null, null);
        Add(styles, "slide-title", font, 30, true, palette.Text, null, null, null);
        Add(styles, "hero-title", font, 32, true, palette.Text, null, null, null);
        Add(styles, "lead", font, 18, false, palette.MutedText, null, null, null);
        Add(styles, "body", font, 14, null, palette.Text, null, null, null);
        Add(styles, "caption", font, 11, false, palette.MutedText, null, null, null);
        Add(styles, "muted", font, 13, false, palette.MutedText, null, null, null);
        Add(styles, "accent", font, 16, true, palette.Primary, null, null, null);
        Add(styles, "card", font, 13, null, palette.Text, palette.Surface, palette.Border, null);
        Add(styles, "callout", font, 15, true, palette.Text, palette.Callout, palette.Primary, null);

        return styles;
    }

    private static void Add(
        IDictionary<string, OfficeMarkupResolvedStyle> styles,
        string name,
        string fontName,
        int fontSize,
        bool? bold,
        string textColor,
        string? fillColor,
        string? borderColor,
        string? textAlign) {
        styles[Normalize(name)] = new OfficeMarkupResolvedStyle {
            Name = name,
            FontName = fontName,
            FontSize = fontSize,
            Bold = bold,
            TextColor = textColor,
            FillColor = fillColor,
            BorderColor = borderColor,
            TextAlign = textAlign
        };
    }

    private static OfficeMarkupResolvedStyle ApplyAttributeOverrides(OfficeMarkupResolvedStyle? style, IDictionary<string, string> attributes) {
        style ??= new OfficeMarkupResolvedStyle();

        if (TryGetAttribute(attributes, "font", out var value) || TryGetAttribute(attributes, "font-name", out value) || TryGetAttribute(attributes, "font-family", out value)) {
            style.FontName = value;
        }

        if (TryGetAttribute(attributes, "font-size", out value) || TryGetAttribute(attributes, "fontsize", out value) || TryGetAttribute(attributes, "size", out value)) {
            if (int.TryParse(TrimUnit(value), out var fontSize)) {
                style.FontSize = fontSize;
            }
        }

        if (TryGetAttribute(attributes, "bold", out value) && TryParseBool(value, out var bold)) {
            style.Bold = bold;
        }

        if (TryGetAttribute(attributes, "italic", out value) && TryParseBool(value, out var italic)) {
            style.Italic = italic;
        }

        if (TryGetAttribute(attributes, "color", out value) || TryGetAttribute(attributes, "text-color", out value) || TryGetAttribute(attributes, "textcolor", out value)) {
            style.TextColor = NormalizeColor(value);
        }

        if (TryGetAttribute(attributes, "fill", out value) || TryGetAttribute(attributes, "fill-color", out value)) {
            style.FillColor = NormalizeColor(value);
        }

        if (TryGetAttribute(attributes, "border", out value) || TryGetAttribute(attributes, "border-color", out value)) {
            style.BorderColor = NormalizeColor(value);
        }

        if (TryGetAttribute(attributes, "align", out value) || TryGetAttribute(attributes, "text-align", out value)) {
            style.TextAlign = value;
        }

        return style;
    }

    private static Palette CreatePalette(string? themeName) {
        switch (Normalize(themeName)) {
            case "evotecmodern":
            case "modernblue":
                return new Palette("Aptos", "#172033", "#4B5563", "#2563EB", "#F4F7FB", "#CAD5E3", "#E8F2FF");
            default:
                return new Palette("Aptos", "#111827", "#4B5563", "#2563EB", "#F3F4F6", "#D1D5DB", "#E0F2FE");
        }
    }

    private static bool TryGetAttribute(IDictionary<string, string> attributes, string name, out string value) =>
        attributes.TryGetValue(name, out value!) && !string.IsNullOrWhiteSpace(value);

    private static string? GetAttribute(IDictionary<string, string> attributes, string name) =>
        attributes.TryGetValue(name, out var value) ? value : null;

    private static bool TryParseBool(string value, out bool result) {
        if (bool.TryParse(value, out result)) {
            return true;
        }

        switch (Normalize(value)) {
            case "yes":
            case "y":
            case "1":
            case "on":
                result = true;
                return true;
            case "no":
            case "n":
            case "0":
            case "off":
                result = false;
                return true;
            default:
                result = false;
                return false;
        }
    }

    private static string TrimUnit(string value) {
        value = value.Trim();
        return value.EndsWith("pt", StringComparison.OrdinalIgnoreCase)
            ? value.Substring(0, value.Length - 2).Trim()
            : value;
    }

    private static string NormalizeColor(string value) {
        value = value.Trim();
        if (value.Length == 6 && value.All(IsHexDigit)) {
            return "#" + value.ToUpperInvariant();
        }

        if (value.StartsWith("#", StringComparison.Ordinal) && value.Length == 7 && value.Substring(1).All(IsHexDigit)) {
            return "#" + value.Substring(1).ToUpperInvariant();
        }

        return value;
    }

    private static bool IsHexDigit(char value) =>
        (value >= '0' && value <= '9')
        || (value >= 'a' && value <= 'f')
        || (value >= 'A' && value <= 'F');

    private static string Normalize(string? value) =>
        (value ?? string.Empty).Replace("-", string.Empty).Replace("_", string.Empty).Replace(" ", string.Empty).ToLowerInvariant();

    private sealed class Palette {
        public Palette(string fontName, string text, string mutedText, string primary, string surface, string border, string callout) {
            FontName = fontName;
            Text = text;
            MutedText = mutedText;
            Primary = primary;
            Surface = surface;
            Border = border;
            Callout = callout;
        }

        public string FontName { get; }
        public string Text { get; }
        public string MutedText { get; }
        public string Primary { get; }
        public string Surface { get; }
        public string Border { get; }
        public string Callout { get; }
    }
}
