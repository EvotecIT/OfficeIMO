using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeIMO.Drawing;

public static partial class OfficeTextBlockRenderer {
    /// <summary>
    /// Appends SVG text elements for a measured rich text block.
    /// </summary>
    /// <param name="builder">SVG markup builder.</param>
    /// <param name="layout">Measured rich text block layout.</param>
    /// <param name="left">Left edge of the available text rectangle.</param>
    /// <param name="top">Top edge of the available text rectangle.</param>
    /// <param name="width">Available text rectangle width.</param>
    /// <param name="height">Available text rectangle height.</param>
    /// <param name="horizontalAlignment">Horizontal alignment inside the rectangle.</param>
    /// <param name="verticalAlignment">Vertical alignment inside the rectangle.</param>
    /// <param name="rotationDegrees">Clockwise rotation in degrees.</param>
    /// <param name="rotationCenterX">Rotation center X coordinate.</param>
    /// <param name="rotationCenterY">Rotation center Y coordinate.</param>
    /// <param name="centerLineInLineHeight">Whether each run glyph box should be vertically centered inside its measured line height.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendSvgRichTextBlock(
        this StringBuilder builder,
        OfficeRichTextBlockLayout layout,
        double left,
        double top,
        double width,
        double height,
        OfficeTextAlignment horizontalAlignment = OfficeTextAlignment.Left,
        OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top,
        double rotationDegrees = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D,
        bool centerLineInLineHeight = true) {
        if (builder == null) {
            throw new ArgumentNullException(nameof(builder));
        }

        if (layout == null) {
            throw new ArgumentNullException(nameof(layout));
        }

        if (layout.Lines.Count == 0 || width <= 0D || height <= 0D) {
            return builder;
        }

        double textTop = OfficeTextPlacement.ResolveTop(top, height, layout.Height, verticalAlignment) + layout.ContentOffsetY;
        double lineTop = textTop;
        for (int lineIndex = 0; lineIndex < layout.Lines.Count; lineIndex++) {
            OfficeRichTextLine line = layout.Lines[lineIndex];
            if (line.Segments.Count == 0) {
                lineTop += ResolveRichTextRenderLineHeight(line, layout.LineHeight);
                continue;
            }

            double lineHeight = ResolveRichTextRenderLineHeight(line, layout.LineHeight);
            double baseline = ResolveRichTextRenderBaseline(line, lineTop, lineHeight, centerLineInLineHeight);
            double lineLeft = left + line.OffsetX;
            double lineWidth = Math.Max(0D, width - line.OffsetX);
            if (ShouldJustifyRichTextLine(line, lineIndex, layout.Lines.Count, lineWidth, horizontalAlignment)) {
                builder.AppendSvgJustifiedRichTextLine(line, lineLeft, baseline, lineWidth, rotationDegrees, rotationCenterX, rotationCenterY);
                lineTop += lineHeight;
                continue;
            }

            double cursor = OfficeTextPlacement.ResolveLineLeft(lineLeft, lineWidth, line.Width, horizontalAlignment);
            for (int segmentIndex = 0; segmentIndex < line.Segments.Count; segmentIndex++) {
                OfficeRichTextSegment segment = line.Segments[segmentIndex];
                builder.AppendSvgRichTextSegmentBackground(segment, cursor, baseline, rotationDegrees, rotationCenterX, rotationCenterY);
                builder.AppendSvgRichTextSegment(segment, cursor, baseline, rotationDegrees, rotationCenterX, rotationCenterY);
                cursor += segment.Width;
            }

            lineTop += lineHeight;
        }

        return builder;
    }

    /// <summary>
    /// Appends SVG text elements for a measured text block.
    /// </summary>
    /// <param name="builder">SVG markup builder.</param>
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
    /// <param name="centerLineInLineHeight">Whether the text glyph box should be vertically centered inside each measured line height.</param>
    /// <param name="strikethrough">Whether to render strikethrough text.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendSvgTextBlock(
        this StringBuilder builder,
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
        bool centerLineInLineHeight = true,
        bool strikethrough = false) =>
        AppendSvgStyledTextBlock(
            builder,
            layout,
            left,
            top,
            width,
            height,
            color,
            fontFamily,
            horizontalAlignment,
            verticalAlignment,
            bold,
            italic,
            underline,
            rotationDegrees,
            rotationCenterX,
            rotationCenterY,
            centerLineInLineHeight,
            strikethrough,
            OfficeTextDecorationStyle.None,
            OfficeTextDecorationStyle.None,
            OfficeTextBaseline.Normal);

    /// <summary>Appends a measured SVG text block with typed decoration and baseline styling.</summary>
    public static StringBuilder AppendSvgStyledTextBlock(
        this StringBuilder builder,
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
        bool centerLineInLineHeight,
        bool strikethrough,
        OfficeTextDecorationStyle underlineStyle,
        OfficeTextDecorationStyle strikethroughStyle,
        OfficeTextBaseline baseline,
        OfficeColor? decorationColor = null) {
        return AppendSvgStyledTextBlockWithFace(builder, layout, null, left, top, width, height, color, fontFamily,
            horizontalAlignment, verticalAlignment, bold, italic, underline, rotationDegrees, rotationCenterX, rotationCenterY,
            centerLineInLineHeight, strikethrough, underlineStyle, strikethroughStyle, baseline, decorationColor);
    }
    internal static StringBuilder AppendSvgStyledTextBlockWithFace(
        this StringBuilder builder,
        OfficeTextBlockLayout layout,
        OfficeFontFaceDescriptor? face,
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
        bool centerLineInLineHeight,
        bool strikethrough,
        OfficeTextDecorationStyle underlineStyle,
        OfficeTextDecorationStyle strikethroughStyle,
        OfficeTextBaseline baseline,
        OfficeColor? decorationColor = null) {
        if (builder == null) {
            throw new ArgumentNullException(nameof(builder));
        }

        if (layout == null) {
            throw new ArgumentNullException(nameof(layout));
        }

        if (layout.Lines.Count == 0 || color.A == 0 || width <= 0D || height <= 0D) {
            return builder;
        }

        string textAnchor = GetSvgTextAnchor(horizontalAlignment);
        OfficeTextDecorationStyle resolvedUnderlineStyle = underlineStyle != OfficeTextDecorationStyle.None
            ? underlineStyle : underline ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None;
        OfficeTextDecorationStyle resolvedStrikethroughStyle = strikethroughStyle != OfficeTextDecorationStyle.None
            ? strikethroughStyle : strikethrough ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None;
        bool splitDecorations = RequiresSeparateSvgDecorations(resolvedUnderlineStyle, resolvedStrikethroughStyle);
        double textTop = OfficeTextPlacement.ResolveTop(top, height, layout.Height, verticalAlignment) + layout.ContentOffsetY;
        for (int i = 0; i < layout.Lines.Count; i++) {
            OfficeTextLine line = layout.Lines[i];
            double lineLeft = left + line.OffsetX;
            double lineWidth = Math.Max(0D, width - line.OffsetX);
            double anchorX = OfficeTextPlacement.ResolveAnchorX(lineLeft, lineWidth, horizontalAlignment);
            double lineTop = textTop + (i * layout.LineHeight);
            double runTop = centerLineInLineHeight
                ? lineTop + Math.Max(0D, (layout.LineHeight - layout.FontSize) / 2D)
                : lineTop;
            double renderedFontSize = baseline == OfficeTextBaseline.Normal ? layout.FontSize : layout.FontSize * 0.65D;
            double renderedBaseline = runTop + (layout.FontSize * 0.84D) + (baseline == OfficeTextBaseline.Superscript
                ? -(layout.FontSize * 0.30D)
                : baseline == OfficeTextBaseline.Subscript ? layout.FontSize * 0.15D : 0D);
            bool justifyLine = ShouldJustifyLine(line, i, layout.Lines.Count, lineWidth, horizontalAlignment);
            builder.Append("<text")
                .AppendNumberAttribute("x", anchorX)
                .AppendNumberAttribute("y", renderedBaseline)
                .AppendPaintAttribute("fill", color)
                .AppendAttribute("font-family", string.IsNullOrWhiteSpace(fontFamily) ? "Arial, sans-serif" : fontFamily)
                .AppendNumberAttribute("font-size", renderedFontSize)
                .AppendAttribute("text-anchor", textAnchor);
            if (justifyLine) {
                builder.AppendNumberAttribute("textLength", lineWidth)
                    .AppendAttribute("lengthAdjust", "spacing");
            }

            if (RequiresSvgWhitespacePreserve(line.Text)) {
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
                builder.Append("<tspan");
                AppendSvgTextDecorationAttribute(builder, resolvedUnderlineStyle, OfficeTextDecorationStyle.None, decorationColor);
                builder.Append('>')
                    .Append(OfficeSvgFormatting.Escape(line.Text))
                    .Append("</tspan>");
            } else {
                builder.Append(OfficeSvgFormatting.Escape(line.Text));
            }
            builder.Append("</text>");
        }

        return builder;
    }

}
