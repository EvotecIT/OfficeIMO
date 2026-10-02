using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeIMO.Drawing;

public static partial class OfficeTextBlockRenderer {
    /// <summary>
    /// Appends one SVG <c>text</c> element with optional <c>tspan</c> children for callers that already resolved placement.
    /// </summary>
    /// <param name="builder">SVG markup builder.</param>
    /// <param name="text">Text content. Hard breaks become <c>tspan</c> children.</param>
    /// <param name="x">Resolved text anchor x-coordinate.</param>
    /// <param name="y">Resolved first-line baseline y-coordinate.</param>
    /// <param name="lineHeight">Distance between line baselines.</param>
    /// <param name="color">Text fill color.</param>
    /// <param name="fontFamily">SVG font-family value.</param>
    /// <param name="fontSize">SVG font size.</param>
    /// <param name="horizontalAlignment">Horizontal alignment used to derive <c>text-anchor</c>.</param>
    /// <param name="bold">Whether to render bold text.</param>
    /// <param name="italic">Whether to render italic text.</param>
    /// <param name="underline">Whether to render underlined text.</param>
    /// <param name="rotationDegrees">Clockwise rotation in degrees.</param>
    /// <param name="rotationCenterX">Rotation center X coordinate.</param>
    /// <param name="rotationCenterY">Rotation center Y coordinate.</param>
    /// <param name="strikethrough">Whether to render strikethrough text.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendSvgTextElement(
        this StringBuilder builder,
        string text,
        double x,
        double y,
        double lineHeight,
        OfficeColor color,
        string? fontFamily,
        double fontSize,
        OfficeTextAlignment horizontalAlignment,
        bool bold,
        bool italic,
        bool underline,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY,
        bool strikethrough) =>
        AppendSvgTextElement(
            builder, text, x, y, lineHeight, color, fontFamily, fontSize,
            horizontalAlignment, bold, italic, underline, rotationDegrees,
            rotationCenterX, rotationCenterY, strikethrough,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, OfficeTextBaseline.Normal);

    /// <summary>Appends one SVG text element with native decoration and baseline styling.</summary>
    /// <remarks><paramref name="underlineStyle"/> and <paramref name="strikethroughStyle"/> take precedence over their legacy Boolean switches; <paramref name="baseline"/> controls subscript and superscript placement.</remarks>
    public static StringBuilder AppendSvgTextElement(
        this StringBuilder builder,
        string text,
        double x,
        double y,
        double lineHeight,
        OfficeColor color,
        string? fontFamily,
        double fontSize,
        OfficeTextAlignment horizontalAlignment = OfficeTextAlignment.Left,
        bool bold = false,
        bool italic = false,
        bool underline = false,
        double rotationDegrees = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D,
        bool strikethrough = false,
        OfficeTextDecorationStyle underlineStyle = OfficeTextDecorationStyle.None,
        OfficeTextDecorationStyle strikethroughStyle = OfficeTextDecorationStyle.None,
        OfficeTextBaseline baseline = OfficeTextBaseline.Normal,
        OfficeColor? decorationColor = null) =>
        AppendSvgTextElementCore(builder, text, x, y, lineHeight, color, fontFamily, fontSize, horizontalAlignment, bold, italic, underline, rotationDegrees, rotationCenterX, rotationCenterY, strikethrough, null, underlineStyle, strikethroughStyle, baseline, decorationColor, featureSettings: null, fontPalette: null);

    internal static StringBuilder AppendSvgFeaturedTextElement(
        this StringBuilder builder,
        string text,
        double x,
        double y,
        double lineHeight,
        OfficeColor color,
        string? fontFamily,
        double fontSize,
        OfficeTextAlignment horizontalAlignment,
        bool bold,
        bool italic,
        bool underline,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY,
        bool strikethrough,
        OfficeTextDecorationStyle underlineStyle,
        OfficeTextDecorationStyle strikethroughStyle,
        OfficeTextBaseline baseline,
        OfficeColor? decorationColor,
        OfficeTextFeatureSettings featureSettings,
        string? fontPalette, OfficeFontFaceDescriptor? face = null) =>
        AppendSvgTextElementCore(builder, text, x, y, lineHeight, color, fontFamily, fontSize, horizontalAlignment, bold, italic, underline, rotationDegrees, rotationCenterX, rotationCenterY, strikethrough, null, underlineStyle, strikethroughStyle, baseline, decorationColor, featureSettings, fontPalette);

    internal static StringBuilder AppendSvgVerticalTextElement(
        this StringBuilder builder,
        string text,
        double x,
        double y,
        OfficeColor color,
        string? fontFamily,
        double fontSize,
        bool bold,
        bool italic,
        bool underline,
        bool strikethrough,
        OfficeTextDecorationStyle underlineStyle,
        OfficeTextDecorationStyle strikethroughStyle,
        OfficeColor? decorationColor,
        OfficeTextFeatureSettings featureSettings,
        string? fontPalette, OfficeFontFaceDescriptor? face = null) =>
        AppendSvgTextElementCore(
            builder, text, x, y, fontSize, color, fontFamily, fontSize, OfficeTextAlignment.Left,
            bold, italic, underline, rotationDegrees: 0D, rotationCenterX: 0D, rotationCenterY: 0D,
            strikethrough, textAdvanceWidth: null, underlineStyle,
            strikethroughStyle, OfficeTextBaseline.Normal, decorationColor,
            featureSettings, fontPalette, OfficeTextDirection.TopToBottom, OfficeTextShapingBackend.BrowserNative, face: face);

    internal static StringBuilder AppendSvgPositionedTextElement(
        this StringBuilder builder,
        string text,
        double x,
        double y,
        double lineHeight,
        OfficeColor color,
        string? fontFamily,
        double fontSize,
        OfficeTextAlignment horizontalAlignment,
        bool bold,
        bool italic,
        bool underline,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY,
        bool strikethrough,
        double? textAdvanceWidth,
        OfficeTextDecorationStyle underlineStyle,
        OfficeTextDecorationStyle strikethroughStyle,
        OfficeTextBaseline baseline,
        OfficeColor? decorationColor = null,
        OfficeTextFeatureSettings? featureSettings = null,
        string? fontPalette = null,
        OfficeTextShapingBackend? shapingBackend = null,
        OfficeTextDirection direction = OfficeTextDirection.Auto,
        bool preservePaintedGlyphOrder = false, OfficeFontFaceDescriptor? face = null) =>
        AppendSvgTextElementCore(builder, text, x, y, lineHeight, color, fontFamily, fontSize, horizontalAlignment, bold, italic, underline, rotationDegrees, rotationCenterX, rotationCenterY, strikethrough, textAdvanceWidth, underlineStyle, strikethroughStyle, baseline, decorationColor, featureSettings, fontPalette, direction, shapingBackend, preservePaintedGlyphOrder, face);

    private static StringBuilder AppendSvgTextElementCore(
        StringBuilder builder,
        string text,
        double x,
        double y,
        double lineHeight,
        OfficeColor color,
        string? fontFamily,
        double fontSize,
        OfficeTextAlignment horizontalAlignment,
        bool bold,
        bool italic,
        bool underline,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY,
        bool strikethrough,
        double? textAdvanceWidth,
        OfficeTextDecorationStyle underlineStyle,
        OfficeTextDecorationStyle strikethroughStyle,
        OfficeTextBaseline baseline,
        OfficeColor? decorationColor,
        OfficeTextFeatureSettings? featureSettings,
        string? fontPalette,
        OfficeTextDirection direction = OfficeTextDirection.Auto,
        OfficeTextShapingBackend? shapingBackend = null,
        bool preservePaintedGlyphOrder = false, OfficeFontFaceDescriptor? face = null) {
        if (builder == null) {
            throw new ArgumentNullException(nameof(builder));
        }

        if (text == null) {
            throw new ArgumentNullException(nameof(text));
        }

        if (color.A == 0) {
            return builder;
        }
        if (textAdvanceWidth.HasValue && (textAdvanceWidth.Value <= 0D || double.IsNaN(textAdvanceWidth.Value) || double.IsInfinity(textAdvanceWidth.Value))) {
            throw new ArgumentOutOfRangeException(nameof(textAdvanceWidth));
        }

        double renderedFontSize = baseline == OfficeTextBaseline.Normal ? fontSize : fontSize * 0.65D;
        double renderedY = baseline == OfficeTextBaseline.Superscript
            ? y - (fontSize * 0.30D)
            : baseline == OfficeTextBaseline.Subscript ? y + (fontSize * 0.15D) : y;
        string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        OfficeTextDecorationStyle resolvedUnderlineStyle = underlineStyle != OfficeTextDecorationStyle.None
            ? underlineStyle : underline ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None;
        OfficeTextDecorationStyle resolvedStrikethroughStyle = strikethroughStyle != OfficeTextDecorationStyle.None
            ? strikethroughStyle : strikethrough ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None;
        OfficeTextDirection resolvedDirection = direction == OfficeTextDirection.Auto
            ? OfficeTextElements.ResolveBaseDirection(text)
            : direction;
        bool splitDecorations = RequiresSeparateSvgDecorations(resolvedUnderlineStyle, resolvedStrikethroughStyle);
        builder.Append("<text")
            .AppendNumberAttribute("x", x)
            .AppendNumberAttribute("y", renderedY)
            .AppendAttribute("font-family", string.IsNullOrWhiteSpace(fontFamily) ? "Arial, sans-serif" : fontFamily)
            .AppendNumberAttribute("font-size", renderedFontSize)
            .AppendAttribute("text-anchor", GetSvgTextAnchor(horizontalAlignment, resolvedDirection))
            .AppendPaintAttribute("fill", color);

        if (featureSettings != null && !featureSettings.IsDefault) {
            var tags = new List<string>(featureSettings.Features.Keys);
            tags.Sort(StringComparer.Ordinal);
            var declaration = new StringBuilder();
            foreach (string tag in tags) {
                if (declaration.Length > 0) declaration.Append(", ");
                declaration.Append('"').Append(tag).Append("\" ")
                    .Append(featureSettings.Features[tag].ToString(System.Globalization.CultureInfo.InvariantCulture));
            }
            builder.AppendAttribute("font-feature-settings", declaration.ToString());
        }

        if (!string.IsNullOrWhiteSpace(fontPalette) && !string.Equals(fontPalette, "normal", StringComparison.OrdinalIgnoreCase)) {
            builder.AppendAttribute("font-palette", fontPalette!.Trim());
        }

        if (direction == OfficeTextDirection.TopToBottom) {
            builder.AppendAttribute("writing-mode", "vertical-rl")
                .AppendAttribute("text-orientation", "mixed");
        }
        if (preservePaintedGlyphOrder) {
            builder.AppendAttribute("direction", "ltr")
                .AppendAttribute("unicode-bidi", "bidi-override");
        } else if (resolvedDirection == OfficeTextDirection.RightToLeft) {
            builder.AppendAttribute("direction", "rtl");
            if (direction == OfficeTextDirection.Auto) {
                builder.AppendAttribute("unicode-bidi", "plaintext");
            }
        }
        if (shapingBackend.HasValue) {
            builder.AppendAttribute("data-officeimo-shaping-backend", shapingBackend.Value == OfficeTextShapingBackend.BrowserNative
                ? "browser-native"
                : shapingBackend.Value.ToString().ToLowerInvariant());
        }

        if (textAdvanceWidth.HasValue && lines.Length == 1) {
            builder.AppendNumberAttribute("textLength", textAdvanceWidth.Value)
                .AppendAttribute("lengthAdjust", "spacingAndGlyphs");
        }

        if (RequiresSvgWhitespacePreserve(text)) {
            builder.Append(" xml:space=\"preserve\"");
        }

        AppendSvgFontFaceAttributes(builder, face, bold, italic);

        AppendSvgTextDecorationAttribute(
            builder,
            splitDecorations ? OfficeTextDecorationStyle.None : resolvedUnderlineStyle,
            resolvedStrikethroughStyle,
            decorationColor);

        if (Math.Abs(rotationDegrees) > 0.000001D) {
            builder.AppendRotateTransformAttribute(rotationDegrees, rotationCenterX, rotationCenterY);
        }

        builder.Append('>');
        if (splitDecorations) {
            for (int i = 0; i < lines.Length; i++) {
                builder.Append("<tspan")
                    .AppendNumberAttribute("x", x)
                    .AppendNumberAttribute("dy", i == 0 ? 0D : lineHeight);
                AppendSvgTextDecorationAttribute(builder, resolvedUnderlineStyle, OfficeTextDecorationStyle.None, decorationColor);
                builder.Append('>')
                    .Append(OfficeSvgFormatting.Escape(lines[i]))
                    .Append("</tspan>");
            }
            builder.Append("</text>");
            return builder;
        }
        for (int i = 0; i < lines.Length; i++) {
            if (i == 0) {
                builder.Append(OfficeSvgFormatting.Escape(lines[i]));
            } else {
                builder.Append("<tspan")
                    .AppendNumberAttribute("x", x)
                    .AppendNumberAttribute("dy", lineHeight)
                    .Append('>')
                    .Append(OfficeSvgFormatting.Escape(lines[i]))
                    .Append("</tspan>");
            }
        }

        builder.Append("</text>");
        return builder;
    }

}
