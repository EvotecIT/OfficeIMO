using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeIMO.Drawing;

public static partial class OfficeTextBlockRenderer {
    /// <summary>
    /// Writes an SVG text block using one <c>text</c> element with measured-line <c>tspan</c> children.
    /// </summary>
    /// <param name="writer">SVG XML writer.</param>
    /// <param name="layout">Measured text block layout.</param>
    /// <param name="left">Left edge of the available text rectangle.</param>
    /// <param name="top">Top edge of the available text rectangle.</param>
    /// <param name="width">Available text rectangle width.</param>
    /// <param name="height">Available text rectangle height.</param>
    /// <param name="color">Text color.</param>
    /// <param name="fontFamily">SVG font-family value.</param>
    /// <param name="horizontalAlignment">Horizontal alignment inside the rectangle.</param>
    /// <param name="verticalAlignment">Vertical alignment inside the rectangle.</param>
    /// <param name="bold">Whether to render bold text.</param>
    /// <param name="italic">Whether to render italic text.</param>
    /// <param name="underline">Whether to render underlined text.</param>
    /// <param name="rotationDegrees">Clockwise rotation in degrees.</param>
    /// <param name="rotationCenterX">Rotation center X coordinate.</param>
    /// <param name="rotationCenterY">Rotation center Y coordinate.</param>
    /// <param name="svgNamespace">SVG namespace URI. Pass <c>null</c> to write elements without a namespace.</param>
    /// <param name="configureTextAttributes">Optional callback for adapter-specific attributes on the <c>text</c> element.</param>
    /// <param name="strikethrough">Whether to render strikethrough text.</param>
    public static void WriteSvgTextBlock(
        XmlWriter writer,
        OfficeTextBlockLayout layout,
        double left,
        double top,
        double width,
        double height,
        OfficeColor color,
        string? fontFamily,
        OfficeTextAlignment horizontalAlignment,
        OfficeTextVerticalAlignment verticalAlignment,
        bool bold,
        bool italic,
        bool underline,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY,
        string? svgNamespace,
        Action<XmlWriter>? configureTextAttributes,
        bool strikethrough) =>
        WriteSvgTextBlock(
            writer, layout, left, top, width, height, color, fontFamily,
            horizontalAlignment, verticalAlignment, bold, italic, underline,
            rotationDegrees, rotationCenterX, rotationCenterY, svgNamespace,
            configureTextAttributes, strikethrough,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, OfficeTextBaseline.Normal);

    /// <summary>Writes an SVG text block with native decoration and baseline styling.</summary>
    /// <remarks><paramref name="underlineStyle"/> and <paramref name="strikethroughStyle"/> take precedence over their legacy Boolean switches; <paramref name="baseline"/> controls subscript and superscript placement.</remarks>
    public static void WriteSvgTextBlock(
        XmlWriter writer,
        OfficeTextBlockLayout layout,
        double left,
        double top,
        double width,
        double height,
        OfficeColor color,
        string? fontFamily,
        OfficeTextAlignment horizontalAlignment = OfficeTextAlignment.Left,
        OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top,
        bool bold = false,
        bool italic = false,
        bool underline = false,
        double rotationDegrees = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D,
        string? svgNamespace = null,
        Action<XmlWriter>? configureTextAttributes = null,
        bool strikethrough = false,
        OfficeTextDecorationStyle underlineStyle = OfficeTextDecorationStyle.None,
        OfficeTextDecorationStyle strikethroughStyle = OfficeTextDecorationStyle.None,
        OfficeTextBaseline baseline = OfficeTextBaseline.Normal) {
        if (writer == null) {
            throw new ArgumentNullException(nameof(writer));
        }

        if (layout == null) {
            throw new ArgumentNullException(nameof(layout));
        }

        if (layout.Lines.Count == 0 || color.A == 0 || width <= 0D || height <= 0D) {
            return;
        }

        double textTop = OfficeTextPlacement.ResolveTop(top, height, layout.Height, verticalAlignment) + layout.ContentOffsetY;
        OfficeTextDecorationStyle resolvedUnderlineStyle = underlineStyle != OfficeTextDecorationStyle.None
            ? underlineStyle : underline ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None;
        OfficeTextDecorationStyle resolvedStrikethroughStyle = strikethroughStyle != OfficeTextDecorationStyle.None
            ? strikethroughStyle : strikethrough ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None;
        bool splitDecorations = RequiresSeparateSvgDecorations(resolvedUnderlineStyle, resolvedStrikethroughStyle);
        double firstAnchorX = OfficeTextPlacement.ResolveAnchorX(left + layout.Lines[0].OffsetX, Math.Max(0D, width - layout.Lines[0].OffsetX), horizontalAlignment);
        writer.WriteStartElement("text", svgNamespace);
        configureTextAttributes?.Invoke(writer);
        writer.WriteNumberAttribute("x", firstAnchorX);
        double renderedFontSize = baseline == OfficeTextBaseline.Normal ? layout.FontSize : layout.FontSize * 0.65D;
        double baselineOffset = baseline == OfficeTextBaseline.Superscript
            ? -(layout.FontSize * 0.30D)
            : baseline == OfficeTextBaseline.Subscript ? layout.FontSize * 0.15D : 0D;
        writer.WriteNumberAttribute("y", textTop + (layout.FontSize / 2D) + baselineOffset);
        writer.WriteAttributeString("font-family", string.IsNullOrWhiteSpace(fontFamily) ? "Arial, sans-serif" : fontFamily);
        writer.WriteNumberAttribute("font-size", renderedFontSize);
        writer.WriteAttributeString("text-anchor", GetSvgTextAnchor(horizontalAlignment));
        writer.WriteAttributeString("dominant-baseline", "middle");
        if (RequiresSvgWhitespacePreserve(layout)) {
            writer.WriteAttributeString("xml", "space", "http://www.w3.org/XML/1998/namespace", "preserve");
        }

        OfficeSvgFormatting.WriteColorAttribute(writer, "fill", color);
        if (bold) {
            writer.WriteAttributeString("font-weight", "700");
        }

        if (italic) {
            writer.WriteAttributeString("font-style", "italic");
        }

        WriteSvgTextDecorationAttribute(
            writer,
            splitDecorations ? OfficeTextDecorationStyle.None : resolvedUnderlineStyle,
            resolvedStrikethroughStyle);

        if (Math.Abs(rotationDegrees) > 0.000001D) {
            writer.WriteRotateTransformAttribute(rotationDegrees, rotationCenterX, rotationCenterY);
        }

        for (int i = 0; i < layout.Lines.Count; i++) {
            OfficeTextLine line = layout.Lines[i];
            double lineAnchorX = OfficeTextPlacement.ResolveAnchorX(left + line.OffsetX, Math.Max(0D, width - line.OffsetX), horizontalAlignment);
            writer.WriteStartElement("tspan", svgNamespace);
            if (splitDecorations) {
                WriteSvgTextDecorationAttribute(writer, resolvedUnderlineStyle, OfficeTextDecorationStyle.None);
            }
            writer.WriteNumberAttribute("x", lineAnchorX);
            writer.WriteNumberAttribute("dy", i == 0 ? 0D : layout.LineHeight);
            double lineWidth = Math.Max(0D, width - line.OffsetX);
            if (ShouldJustifyLine(line, i, layout.Lines.Count, lineWidth, horizontalAlignment)) {
                writer.WriteNumberAttribute("textLength", lineWidth);
                writer.WriteAttributeString("lengthAdjust", "spacing");
            }

            writer.WriteString(line.Text);
            writer.WriteEndElement();
        }

        writer.WriteEndElement();
    }

    /// <summary>
    /// Writes a measured SVG text-box plan, including an optional text background.
    /// </summary>
    /// <param name="writer">SVG XML writer.</param>
    /// <param name="plan">Resolved text-box layout and placement.</param>
    /// <param name="color">Text color.</param>
    /// <param name="fontFamily">SVG font-family value.</param>
    /// <param name="bold">Whether to render bold text.</param>
    /// <param name="italic">Whether to render italic text.</param>
    /// <param name="underline">Whether to render underlined text.</param>
    /// <param name="rotationDegrees">Clockwise rotation in degrees.</param>
    /// <param name="rotationCenterX">Rotation center X coordinate.</param>
    /// <param name="rotationCenterY">Rotation center Y coordinate.</param>
    /// <param name="svgNamespace">SVG namespace URI. Pass <c>null</c> to write elements without a namespace.</param>
    /// <param name="backgroundColor">Optional background color around the measured text block.</param>
    /// <param name="backgroundPaddingX">Horizontal background padding.</param>
    /// <param name="backgroundPaddingY">Vertical background padding.</param>
    /// <param name="configureTextAttributes">Optional callback for adapter-specific attributes on the <c>text</c> element.</param>
    /// <param name="configureBackgroundAttributes">Optional callback for adapter-specific attributes on the background <c>rect</c> element.</param>
    /// <param name="strikethrough">Whether to render strikethrough text.</param>
    public static void WriteSvgTextBox(
        XmlWriter writer,
        OfficeTextBlockRenderPlan plan,
        OfficeColor color,
        string? fontFamily,
        bool bold,
        bool italic,
        bool underline,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY,
        string? svgNamespace,
        OfficeColor? backgroundColor,
        double backgroundPaddingX,
        double backgroundPaddingY,
        Action<XmlWriter>? configureTextAttributes,
        Action<XmlWriter>? configureBackgroundAttributes,
        bool strikethrough) =>
        WriteSvgTextBox(
            writer, plan, color, fontFamily, bold, italic, underline,
            rotationDegrees, rotationCenterX, rotationCenterY, svgNamespace,
            backgroundColor, backgroundPaddingX, backgroundPaddingY,
            configureTextAttributes, configureBackgroundAttributes, strikethrough,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, OfficeTextBaseline.Normal);

}
