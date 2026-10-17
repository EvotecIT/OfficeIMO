using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
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
        A.ColorScheme? colorScheme = paragraph._document.MainDocumentPartRoot.ThemePart?.Theme?.ThemeElements?.ColorScheme;
        return TryResolveShadingFillColor(ResolveRunShading(paragraph), colorScheme, out OfficeColor shading)
            ? shading.ToRgbHex()
            : null;
    }

    /// <summary>Resolves direct and styled highlighting, then table style and document defaults; explicit none stops inheritance.</summary>
    internal static HighlightColorValues? ResolveRunHighlight(WordParagraph paragraph, HighlightColorValues? tableStyleHighlight = null) {
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
        foreach (Style style in EnumerateStyleChain(styles, paragraph.CharacterStyleId, StyleValues.Character)
                     .Concat(EnumerateStyleChain(styles, paragraph.StyleId, StyleValues.Paragraph))) {
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

        HashSet<string> visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        while (!string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId!)) {
            Style? style = styles.Elements<Style>()
                .FirstOrDefault(candidate =>
                    string.Equals(candidate.StyleId?.Value, styleId, StringComparison.OrdinalIgnoreCase) &&
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

        OfficeColor foreground = ResolveShadingForeground(shading, colorScheme);
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

    private static OfficeColor ResolveShadingForeground(Shading shading, A.ColorScheme? colorScheme) {
        string? resolvedThemeForeground = ResolveThemeColor(
            GetWordAttribute(shading, "themeColor"),
            GetWordAttribute(shading, "themeTint"),
            GetWordAttribute(shading, "themeShade"),
            colorScheme);
        if (OfficeColor.TryParseHex(resolvedThemeForeground, out OfficeColor themeForeground)) {
            return themeForeground;
        }
        if (OfficeColor.TryParseHex(shading.Color?.Value, out OfficeColor foreground)) {
            return foreground;
        }

        return OfficeColor.Black;
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
