using System;
using System.Text;
using System.Xml;

namespace OfficeIMO.Drawing;

/// <summary>
/// Renders measured text blocks through the shared dependency-free Drawing primitives.
/// </summary>
public static class OfficeTextBlockRenderer {
    /// <summary>
    /// Draws a measured text block on a raster canvas.
    /// </summary>
    /// <param name="canvas">Raster canvas receiving the text.</param>
    /// <param name="layout">Measured text block layout.</param>
    /// <param name="left">Left edge of the available text rectangle.</param>
    /// <param name="top">Top edge of the available text rectangle.</param>
    /// <param name="width">Available text rectangle width.</param>
    /// <param name="height">Available text rectangle height.</param>
    /// <param name="color">Text color.</param>
    /// <param name="horizontalAlignment">Horizontal alignment inside the rectangle.</param>
    /// <param name="verticalAlignment">Vertical alignment inside the rectangle.</param>
    /// <param name="bold">Whether to render bold text.</param>
    /// <param name="italic">Whether to render italic text.</param>
    /// <param name="underline">Whether to render an underline for each visible line.</param>
    /// <param name="rotationDegrees">Clockwise rotation in degrees.</param>
    /// <param name="rotationCenterX">Rotation center X coordinate.</param>
    /// <param name="rotationCenterY">Rotation center Y coordinate.</param>
    /// <param name="centerLineInLineHeight">Whether the text glyph box should be vertically centered inside each measured line height.</param>
    /// <param name="underlineOffsetFactor">Underline baseline offset as a factor of the resolved font size.</param>
    /// <param name="strikethrough">Whether to render a strikethrough for each visible line.</param>
    public static void DrawRasterTextBlock(
        OfficeRasterCanvas canvas,
        OfficeTextBlockLayout layout,
        double left,
        double top,
        double width,
        double height,
        OfficeColor color,
        OfficeTextAlignment horizontalAlignment = OfficeTextAlignment.Left,
        OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top,
        bool bold = false,
        bool italic = false,
        bool underline = false,
        double rotationDegrees = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D,
        bool centerLineInLineHeight = true,
        double underlineOffsetFactor = 0.86D,
        bool strikethrough = false) {
        if (canvas == null) {
            throw new ArgumentNullException(nameof(canvas));
        }

        if (layout == null) {
            throw new ArgumentNullException(nameof(layout));
        }

        if (layout.Lines.Count == 0 || color.A == 0 || width <= 0D || height <= 0D) {
            return;
        }

        double textTop = OfficeTextPlacement.ResolveTop(top, height, layout.Height, verticalAlignment);
        double anchorX = OfficeTextPlacement.ResolveAnchorX(left, width, horizontalAlignment);
        for (int i = 0; i < layout.Lines.Count; i++) {
            OfficeTextLine line = layout.Lines[i];
            double lineTop = textTop + (i * layout.LineHeight);
            double runTop = centerLineInLineHeight
                ? lineTop + Math.Max(0D, (layout.LineHeight - layout.FontSize) / 2D)
                : lineTop;
            canvas.DrawTextLine(line.Text, anchorX, runTop, layout.FontSize, color, bold, italic, horizontalAlignment, rotationDegrees, rotationCenterX, rotationCenterY, underline, strikethrough);
        }
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
        bool strikethrough = false) {
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
        double textTop = OfficeTextPlacement.ResolveTop(top, height, layout.Height, verticalAlignment);
        double anchorX = OfficeTextPlacement.ResolveAnchorX(left, width, horizontalAlignment);
        for (int i = 0; i < layout.Lines.Count; i++) {
            OfficeTextLine line = layout.Lines[i];
            double lineTop = textTop + (i * layout.LineHeight);
            double runTop = centerLineInLineHeight
                ? lineTop + Math.Max(0D, (layout.LineHeight - layout.FontSize) / 2D)
                : lineTop;
            double baseline = runTop + (layout.FontSize * 0.84D);
            builder.Append("<text")
                .AppendNumberAttribute("x", anchorX)
                .AppendNumberAttribute("y", baseline)
                .AppendPaintAttribute("fill", color)
                .AppendAttribute("font-family", string.IsNullOrWhiteSpace(fontFamily) ? "Arial, sans-serif" : fontFamily)
                .AppendNumberAttribute("font-size", layout.FontSize)
                .AppendAttribute("text-anchor", textAnchor);
            if (bold) {
                builder.Append(" font-weight=\"700\"");
            }

            if (italic) {
                builder.Append(" font-style=\"italic\"");
            }

            AppendSvgTextDecorationAttribute(builder, underline, strikethrough);

            if (Math.Abs(rotationDegrees) > 0.000001D) {
                builder.AppendRotateTransformAttribute(rotationDegrees, rotationCenterX, rotationCenterY);
            }

            builder.Append(">")
                .Append(OfficeSvgFormatting.Escape(line.Text))
                .Append("</text>");
        }

        return builder;
    }

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
        OfficeTextAlignment horizontalAlignment = OfficeTextAlignment.Left,
        bool bold = false,
        bool italic = false,
        bool underline = false,
        double rotationDegrees = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D,
        bool strikethrough = false) {
        if (builder == null) {
            throw new ArgumentNullException(nameof(builder));
        }

        if (text == null) {
            throw new ArgumentNullException(nameof(text));
        }

        if (color.A == 0) {
            return builder;
        }

        string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        builder.Append("<text")
            .AppendNumberAttribute("x", x)
            .AppendNumberAttribute("y", y)
            .AppendAttribute("font-family", string.IsNullOrWhiteSpace(fontFamily) ? "Arial, sans-serif" : fontFamily)
            .AppendNumberAttribute("font-size", fontSize)
            .AppendAttribute("text-anchor", GetSvgTextAnchor(horizontalAlignment))
            .AppendPaintAttribute("fill", color);

        if (bold) {
            builder.Append(" font-weight=\"700\"");
        }

        if (italic) {
            builder.Append(" font-style=\"italic\"");
        }

        AppendSvgTextDecorationAttribute(builder, underline, strikethrough);

        if (Math.Abs(rotationDegrees) > 0.000001D) {
            builder.AppendRotateTransformAttribute(rotationDegrees, rotationCenterX, rotationCenterY);
        }

        builder.Append('>');
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

    /// <summary>
    /// Appends one SVG <c>text</c> element for a measured rich text segment.
    /// </summary>
    /// <param name="builder">SVG markup builder.</param>
    /// <param name="segment">Measured rich text segment.</param>
    /// <param name="x">Resolved segment x-coordinate.</param>
    /// <param name="baseline">Resolved segment baseline y-coordinate.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendSvgRichTextSegment(this StringBuilder builder, OfficeRichTextSegment segment, double x, double baseline) {
        if (segment == null) {
            throw new ArgumentNullException(nameof(segment));
        }

        return builder.AppendSvgTextElement(
            segment.Text,
            x,
            baseline,
            segment.FontSize,
            segment.Color,
            segment.FontFamily,
            segment.FontSize,
            OfficeTextAlignment.Left,
            segment.Bold,
            segment.Italic,
            segment.Underline,
            strikethrough: segment.Strikethrough);
    }

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
        bool strikethrough = false) {
        if (writer == null) {
            throw new ArgumentNullException(nameof(writer));
        }

        if (layout == null) {
            throw new ArgumentNullException(nameof(layout));
        }

        if (layout.Lines.Count == 0 || color.A == 0 || width <= 0D || height <= 0D) {
            return;
        }

        double textTop = OfficeTextPlacement.ResolveTop(top, height, layout.Height, verticalAlignment);
        double anchorX = OfficeTextPlacement.ResolveAnchorX(left, width, horizontalAlignment);
        writer.WriteStartElement("text", svgNamespace);
        configureTextAttributes?.Invoke(writer);
        writer.WriteNumberAttribute("x", anchorX);
        writer.WriteNumberAttribute("y", textTop + (layout.FontSize / 2D));
        writer.WriteAttributeString("font-family", string.IsNullOrWhiteSpace(fontFamily) ? "Arial, sans-serif" : fontFamily);
        writer.WriteNumberAttribute("font-size", layout.FontSize);
        writer.WriteAttributeString("text-anchor", GetSvgTextAnchor(horizontalAlignment));
        writer.WriteAttributeString("dominant-baseline", "middle");
        OfficeSvgFormatting.WriteColorAttribute(writer, "fill", color);
        if (bold) {
            writer.WriteAttributeString("font-weight", "700");
        }

        if (italic) {
            writer.WriteAttributeString("font-style", "italic");
        }

        WriteSvgTextDecorationAttribute(writer, underline, strikethrough);

        if (Math.Abs(rotationDegrees) > 0.000001D) {
            writer.WriteRotateTransformAttribute(rotationDegrees, rotationCenterX, rotationCenterY);
        }

        for (int i = 0; i < layout.Lines.Count; i++) {
            writer.WriteStartElement("tspan", svgNamespace);
            writer.WriteNumberAttribute("x", anchorX);
            writer.WriteNumberAttribute("dy", i == 0 ? 0D : layout.LineHeight);
            writer.WriteString(layout.Lines[i].Text);
            writer.WriteEndElement();
        }

        writer.WriteEndElement();
    }

    private static string GetSvgTextAnchor(OfficeTextAlignment alignment) {
        switch (alignment) {
            case OfficeTextAlignment.Right:
                return "end";
            case OfficeTextAlignment.Center:
                return "middle";
            default:
                return "start";
        }
    }

    private static void AppendSvgTextDecorationAttribute(StringBuilder builder, bool underline, bool strikethrough) {
        if (!underline && !strikethrough) {
            return;
        }

        builder.Append(" text-decoration=\"");
        if (underline) {
            builder.Append("underline");
        }

        if (underline && strikethrough) {
            builder.Append(' ');
        }

        if (strikethrough) {
            builder.Append("line-through");
        }

        builder.Append('"');
    }

    private static void WriteSvgTextDecorationAttribute(XmlWriter writer, bool underline, bool strikethrough) {
        if (!underline && !strikethrough) {
            return;
        }

        string value = underline && strikethrough
            ? "underline line-through"
            : underline ? "underline" : "line-through";
        writer.WriteAttributeString("text-decoration", value);
    }
}
