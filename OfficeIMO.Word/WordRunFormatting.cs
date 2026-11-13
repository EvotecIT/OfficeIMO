using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word;

/// <summary>Resolves inherited Word run backgrounds without performing page layout or rendering.</summary>
internal static partial class WordRunFormatting {
    /// <summary>Resolves run shading, including theme colors and a blended fill for shading patterns.</summary>
    internal static string? ResolveRunShadingFillColorHex(WordParagraph paragraph) {
        RunPaintConversion? conversion = RunPaintConversions.Value;
        if (conversion == null || paragraph._paragraph == null) {
            return ResolveShadingFillColorHex(ResolveRunShading(paragraph), paragraph._document);
        }

        var key = new RunPaintKey(GetDirectRunProperties(paragraph), paragraph.CharacterStyleId, paragraph._paragraph);
        lock (conversion.RunShadings) {
            if (conversion.RunShadings.TryGetValue(key, out string? cached)) return cached;
        }
        string? resolved = ResolveShadingFillColorHex(ResolveRunShading(paragraph), paragraph._document);
        lock (conversion.RunShadings) conversion.RunShadings[key] = resolved;
        return resolved;
    }

    /// <summary>Resolves an authored table, cell or run fill against its document theme without changing its XML.</summary>
    internal static string? ResolveShadingFillColorHex(Shading? shading, WordDocument document) {
        A.ColorScheme? colorScheme = GetLiveColorScheme(document);
        return TryResolveShadingFillColor(shading, colorScheme, out OfficeColor fill)
            ? fill.ToRgbHex()
            : null;
    }

    /// <summary>Resolves the source background independently of the foreground and pattern retained by format writers.</summary>
    internal static string? ResolveShadingBackgroundColorHex(Shading? shading, WordDocument document) =>
        shading == null || shading.Val?.Value == ShadingPatternValues.Nil ? null :
            TryResolveShadingBackground(shading, GetLiveColorScheme(document), out OfficeColor background)
                ? background.ToRgbHex() : null;

    /// <summary>Resolves the source pattern foreground without flattening it into its background.</summary>
    internal static string? ResolveShadingForegroundColorHex(Shading? shading, WordDocument document) =>
        shading != null && TryResolveShadingForeground(shading, GetLiveColorScheme(document), out OfficeColor foreground)
            ? foreground.ToRgbHex() : null;

    /// <summary>Resolves Word text color with theme references taking precedence over cached RGB.</summary>
    internal static string? ResolveColorHex(Color? color, WordDocument? document) => color == null ? null :
        ResolveThemeColor(color.ThemeColor?.InnerText, color.ThemeTint?.Value, color.ThemeShade?.Value,
            GetLiveColorScheme(document)) ?? color.Val?.Value;

    /// <summary>Resolves an authored border color while retaining automatic-color declarations.</summary>
    internal static string? ResolveBorderColorHex(BorderType? border, WordDocument document) => border == null ? null :
        ResolveThemeColor(border.ThemeColor?.InnerText, border.ThemeTint?.Value, border.ThemeShade?.Value,
            GetLiveColorScheme(document)) ?? border.Color?.Value;

    /// <summary>Resolves a theme colour selector with its tint and shade against the document's current theme.</summary>
    internal static string? ResolveThemeColorHex(string? themeColor, string? tint, string? shade, WordDocument? document) =>
        ResolveThemeColor(themeColor, tint, shade, GetLiveColorScheme(document));

    private static A.ColorScheme? GetLiveColorScheme(WordDocument? document) =>
        document?._wordprocessingDocument?.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.ColorScheme;

    /// <summary>Resolves direct and styled highlighting, then table style and document defaults; explicit none stops inheritance.</summary>
    internal static HighlightColorValues? ResolveRunHighlight(WordParagraph paragraph, HighlightColorValues? tableStyleHighlight = null) {
        RunPaintConversion? conversion = RunPaintConversions.Value;
        if (conversion == null || paragraph._paragraph == null) return ResolveRunHighlightCore(paragraph, tableStyleHighlight);

        var key = (new RunPaintKey(GetDirectRunProperties(paragraph), paragraph.CharacterStyleId, paragraph._paragraph), tableStyleHighlight);
        lock (conversion.RunHighlights) {
            if (conversion.RunHighlights.TryGetValue(key, out HighlightColorValues? cached)) return cached;
        }
        HighlightColorValues? resolved = ResolveRunHighlightCore(paragraph, tableStyleHighlight);
        lock (conversion.RunHighlights) conversion.RunHighlights[key] = resolved;
        return resolved;
    }

    private static HighlightColorValues? ResolveRunHighlightCore(WordParagraph paragraph, HighlightColorValues? tableStyleHighlight) {
        foreach (OpenXmlCompositeElement properties in EnumerateRunPropertySources(paragraph, includeDocumentDefaults: false)) {
            HighlightColorValues? highlight = ReadRunHighlight(properties);
            if (highlight.HasValue) return highlight;
        }
        return tableStyleHighlight ?? ReadRunHighlight(GetDocumentDefaultRunProperties(paragraph));
    }

    /// <summary>Reads typed or unknown SDK highlight elements after a style/default run-property round trip.</summary>
    internal static HighlightColorValues? ReadRunHighlight(OpenXmlCompositeElement? properties) {
        OpenXmlElement? highlight = properties?.ChildElements.FirstOrDefault(child =>
            child.LocalName == "highlight" && child.NamespaceUri == properties.NamespaceUri);
        if (highlight == null) return null;
        string? value = highlight.GetAttributes().FirstOrDefault(attribute =>
            attribute.LocalName == "val" && attribute.NamespaceUri == highlight.NamespaceUri).Value;
        return string.IsNullOrEmpty(value) ? null : new HighlightColorValues(value!);
    }

    private static Shading? ResolveRunShading(WordParagraph paragraph) {
        foreach (OpenXmlCompositeElement properties in EnumerateRunPropertySources(paragraph)) {
            Shading? shading = properties.GetFirstChild<Shading>();
            if (shading != null) {
                return shading;
            }
        }

        return null;
    }

    private static RunProperties? GetDirectRunProperties(WordParagraph paragraph) =>
        paragraph.IsHyperLink
            ? paragraph.Hyperlink?._runProperties
            : paragraph._runProperties;

    /// <summary>Enumerates direct, character-style, paragraph-style, table-style, then document-default run properties.</summary>
    private static IEnumerable<OpenXmlCompositeElement> EnumerateRunPropertySources(WordParagraph paragraph, bool includeDocumentDefaults = true) {
        RunProperties? directProperties = GetDirectRunProperties(paragraph);
        if (directProperties != null) {
            yield return directProperties;
        }

        foreach (StyleRunProperties properties in EnumerateRunStyleProperties(paragraph)) {
            yield return properties;
        }

        foreach (OpenXmlCompositeElement properties in EnumerateTableRunPropertySources(paragraph)) {
            yield return properties;
        }

        RunPropertiesBaseStyle? defaults = includeDocumentDefaults ? GetDocumentDefaultRunProperties(paragraph) : null;
        if (defaults != null) yield return defaults;
    }

    private static RunPropertiesBaseStyle? GetDocumentDefaultRunProperties(WordParagraph paragraph) =>
        paragraph._document?._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles?
            .DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;

    /// <summary>Enumerates character-style then paragraph-style run properties, each in override order.</summary>
    internal static IEnumerable<StyleRunProperties> EnumerateRunStyleProperties(WordParagraph paragraph) {
        Styles? styles = paragraph._document?._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        foreach (Style style in GetStyleChain(styles, paragraph.CharacterStyleId, StyleValues.Character)) {
            if (style.StyleRunProperties != null) {
                yield return style.StyleRunProperties;
            }
        }
        foreach (Style style in GetStyleChain(styles, paragraph.StyleId, StyleValues.Paragraph)) {
            if (style.StyleRunProperties != null) {
                yield return style.StyleRunProperties;
            }
        }
    }

    /// <summary>
    /// Enumerates a style and its ancestors in override order, using the default style only when no ID is supplied.
    /// Missing styles, mismatched style types, and inheritance cycles terminate the chain.
    /// </summary>
    internal static IEnumerable<Style> EnumerateStyleChain(Styles? styles, string? styleId, StyleValues expectedType) {
        if (styles == null) yield break;
        if (string.IsNullOrWhiteSpace(styleId))
            styleId = styles.Elements<Style>().FirstOrDefault(style =>
                style.Type?.Value == expectedType && style.Default?.Value == true)?.StyleId?.Value;
        if (string.IsNullOrWhiteSpace(styleId)) yield break;

        HashSet<string> visited = new HashSet<string>(StringComparer.Ordinal);
        while (!string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId!)) {
            Style? style = styles.Elements<Style>()
                .FirstOrDefault(candidate =>
                    string.Equals(candidate.StyleId?.Value, styleId, StringComparison.Ordinal) &&
                    (candidate.Type == null || !candidate.Type.HasValue || candidate.Type.Value == expectedType));
            if (style == null) {
                yield break;
            }

            yield return style;
            styleId = style.BasedOn?.Val?.Value;
        }
    }

    private static bool TryResolveShadingFillColor(
        Shading? shading,
        A.ColorScheme? colorScheme,
        out OfficeColor fillColor) {
        fillColor = OfficeColor.White;
        if (shading == null || shading.Val?.Value == ShadingPatternValues.Nil) {
            return false;
        }

        string pattern = GetWordAttribute(shading, "val") ?? "clear";
        bool hasBackground = TryResolveShadingBackground(shading, colorScheme, out OfficeColor background);
        if (pattern.Equals("clear", StringComparison.OrdinalIgnoreCase)) {
            if (hasBackground) {
                fillColor = background;
            }
            return hasBackground;
        }

        OfficeColor foreground = TryResolveShadingForeground(shading, colorScheme, out OfficeColor resolvedForeground)
            ? resolvedForeground : OfficeColor.Black;
        if (pattern.Equals("solid", StringComparison.OrdinalIgnoreCase)) {
            fillColor = foreground;
            return true;
        }

        background = hasBackground ? background : OfficeColor.White;
        fillColor = BlendShadingPattern(
            foreground,
            background,
            ResolveShadingPatternForegroundRatio(pattern));
        return true;
    }

    private static bool TryResolveShadingForeground(Shading shading, A.ColorScheme? colorScheme, out OfficeColor foreground) {
        string? resolvedThemeForeground = ResolveThemeColor(
            GetWordAttribute(shading, "themeColor"),
            GetWordAttribute(shading, "themeTint"),
            GetWordAttribute(shading, "themeShade"),
            colorScheme);
        return OfficeColor.TryParseHex(resolvedThemeForeground, out foreground)
            || OfficeColor.TryParseHex(shading.Color?.Value, out foreground);
    }

    private static bool TryResolveShadingBackground(Shading shading, A.ColorScheme? colorScheme, out OfficeColor background) {
        string? resolvedThemeBackground = ResolveThemeColor(
            GetWordAttribute(shading, "themeFill"),
            GetWordAttribute(shading, "themeFillTint"),
            GetWordAttribute(shading, "themeFillShade"),
            colorScheme);
        if (OfficeColor.TryParseHex(resolvedThemeBackground, out background)) {
            return true;
        }

        return OfficeColor.TryParseHex(shading.Fill?.Value, out background);
    }

    private static double ResolveShadingPatternForegroundRatio(string pattern) {
        if (pattern.StartsWith("pct", StringComparison.OrdinalIgnoreCase) &&
            int.TryParse(pattern.Substring(3), out int percentage)) {
            return percentage switch {
                12 => 0.125d,
                37 => 0.375d,
                62 => 0.625d,
                87 => 0.875d,
                _ => Math.Max(0d, Math.Min(1d, percentage / 100d))
            };
        }

        return pattern.StartsWith("thin", StringComparison.OrdinalIgnoreCase)
            ? 0.25d
            : 0.5d;
    }

    private static OfficeColor BlendShadingPattern(OfficeColor foreground, OfficeColor background, double foregroundRatio) {
        double backgroundRatio = 1d - foregroundRatio;
        return OfficeColor.FromRgb(
            (byte)Math.Round(foreground.R * foregroundRatio + background.R * backgroundRatio),
            (byte)Math.Round(foreground.G * foregroundRatio + background.G * backgroundRatio),
            (byte)Math.Round(foreground.B * foregroundRatio + background.B * backgroundRatio));
    }

    private static string? ResolveThemeColor(string? themeColor, string? tint, string? shade, A.ColorScheme? colorScheme) {
        OfficeColor? resolvedColor = OfficeOpenXmlThemeColorResolver.ResolveSchemeColor(colorScheme, themeColor);
        if (!resolvedColor.HasValue) {
            return null;
        }

        return ApplyWordThemeTransforms(resolvedColor.Value, tint, shade).ToRgbHex();
    }

    private static OfficeColor ApplyWordThemeTransforms(OfficeColor color, string? tint, string? shade) {
        OfficeColor transformed = color;
        if (TryParseHexByte(shade, out int shadeValue)) {
            double amount = Math.Max(0D, Math.Min(255D, shadeValue)) / 255D;
            transformed = OfficeColorTransforms.Shade(transformed, amount);
        }

        if (TryParseHexByte(tint, out int tintValue)) {
            double inputRatio = Math.Max(0D, Math.Min(255D, tintValue)) / 255D;
            transformed = OfficeColorTransforms.Tint(transformed, inputRatio);
        }

        return transformed;
    }

    private static string? GetWordAttribute(OpenXmlElement? element, string localName) {
        if (element == null) {
            return null;
        }

        foreach (OpenXmlAttribute attribute in element.GetAttributes()) {
            if (string.Equals(attribute.LocalName, localName, StringComparison.Ordinal)) {
                return attribute.Value;
            }
        }

        return null;
    }

    private static bool TryParseHexByte(string? value, out int result) {
        result = 0;
        return !string.IsNullOrWhiteSpace(value) && int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out result);
    }
}
