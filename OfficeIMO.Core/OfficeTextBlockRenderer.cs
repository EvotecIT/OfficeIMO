using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

namespace OfficeIMO.Drawing;

/// <summary>
/// Renders measured text blocks through the shared dependency-free Drawing primitives.
/// </summary>
public static partial class OfficeTextBlockRenderer {
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
    /// <param name="fontFamily">Requested font family fallback list.</param>
    /// <param name="flipHorizontal">Whether to mirror each rendered line horizontally around the rotation center before rotation.</param>
    /// <param name="flipVertical">Whether to mirror each rendered line vertically around the rotation center before rotation.</param>
    public static void DrawRasterTextBlock(
        OfficeRasterCanvas canvas,
        OfficeTextBlockLayout layout,
        double left,
        double top,
        double width,
        double height,
        OfficeColor color,
        OfficeTextAlignment horizontalAlignment,
        OfficeTextVerticalAlignment verticalAlignment,
        bool bold,
        bool italic,
        bool underline,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY,
        bool centerLineInLineHeight,
        double underlineOffsetFactor,
        bool strikethrough,
        string? fontFamily,
        bool flipHorizontal,
        bool flipVertical) =>
        DrawRasterTextBlock(
            canvas, layout, left, top, width, height, color, horizontalAlignment, verticalAlignment,
            bold, italic, underline, rotationDegrees, rotationCenterX, rotationCenterY,
            centerLineInLineHeight, underlineOffsetFactor, strikethrough, fontFamily,
            flipHorizontal, flipVertical,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, OfficeTextBaseline.Normal);

    /// <summary>Draws a measured text block with native decoration and baseline styling.</summary>
    /// <remarks><paramref name="underlineStyle"/> and <paramref name="strikethroughStyle"/> take precedence over their legacy Boolean switches; <paramref name="baseline"/> controls subscript and superscript placement.</remarks>
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
        bool strikethrough = false,
        string? fontFamily = null,
        bool flipHorizontal = false,
        bool flipVertical = false,
        OfficeTextDecorationStyle underlineStyle = OfficeTextDecorationStyle.None,
        OfficeTextDecorationStyle strikethroughStyle = OfficeTextDecorationStyle.None,
        OfficeTextBaseline baseline = OfficeTextBaseline.Normal,
        OfficeColor? decorationColor = null) {
        if (canvas == null) {
            throw new ArgumentNullException(nameof(canvas));
        }

        if (layout == null) {
            throw new ArgumentNullException(nameof(layout));
        }

        if (layout.Lines.Count == 0 || color.A == 0 || width <= 0D || height <= 0D) {
            return;
        }

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
            runTop += baseline == OfficeTextBaseline.Superscript
                ? -(layout.FontSize * 0.30D)
                : baseline == OfficeTextBaseline.Subscript ? layout.FontSize * 0.15D : 0D;
            if (ShouldJustifyLine(line, i, layout.Lines.Count, lineWidth, horizontalAlignment)) {
                DrawRasterJustifiedTextLine(
                    canvas,
                    line.Text,
                    lineLeft,
                    lineWidth,
                    runTop,
                    renderedFontSize,
                    color,
                    bold,
                    italic,
                    rotationDegrees,
                    rotationCenterX,
                    rotationCenterY,
                    underline,
                    strikethrough,
                    fontFamily,
                    flipHorizontal,
                    flipVertical,
                    underlineStyle,
                    strikethroughStyle,
                    decorationColor);
                continue;
            }

            canvas.DrawTextLine(line.Text, anchorX, runTop, renderedFontSize, color, bold, italic, horizontalAlignment, rotationDegrees, rotationCenterX, rotationCenterY, underline, strikethrough, fontFamily, flipHorizontal, flipVertical, underlineStyle, strikethroughStyle, decorationColor);
        }
    }

    /// <summary>
    /// Draws a measured text-box plan on a raster canvas, including an optional text background.
    /// </summary>
    /// <param name="canvas">Raster canvas receiving the text.</param>
    /// <param name="plan">Resolved text-box layout and placement.</param>
    /// <param name="color">Text color.</param>
    /// <param name="bold">Whether to render bold text.</param>
    /// <param name="italic">Whether to render italic text.</param>
    /// <param name="underline">Whether to render an underline for each visible line.</param>
    /// <param name="horizontalAlignment">Horizontal alignment override. Pass <c>null</c> to use <paramref name="plan"/>.</param>
    /// <param name="verticalAlignment">Vertical alignment override. Pass <c>null</c> to use <paramref name="plan"/>.</param>
    /// <param name="rotationDegrees">Clockwise rotation in degrees.</param>
    /// <param name="rotationCenterX">Rotation center X coordinate.</param>
    /// <param name="rotationCenterY">Rotation center Y coordinate.</param>
    /// <param name="backgroundColor">Optional background color around the measured text block.</param>
    /// <param name="backgroundPaddingX">Horizontal background padding.</param>
    /// <param name="backgroundPaddingY">Vertical background padding.</param>
    /// <param name="centerLineInLineHeight">Whether the text glyph box should be vertically centered inside each measured line height.</param>
    /// <param name="underlineOffsetFactor">Underline baseline offset as a factor of the resolved font size.</param>
    /// <param name="strikethrough">Whether to render a strikethrough for each visible line.</param>
    /// <param name="fontFamily">Requested font family fallback list.</param>
    public static void DrawRasterTextBox(
        OfficeRasterCanvas canvas,
        OfficeTextBlockRenderPlan plan,
        OfficeColor color,
        bool bold,
        bool italic,
        bool underline,
        OfficeTextAlignment? horizontalAlignment,
        OfficeTextVerticalAlignment? verticalAlignment,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY,
        OfficeColor? backgroundColor,
        double backgroundPaddingX,
        double backgroundPaddingY,
        bool centerLineInLineHeight,
        double underlineOffsetFactor,
        bool strikethrough,
        string? fontFamily) =>
        DrawRasterTextBox(
            canvas, plan, color, bold, italic, underline, horizontalAlignment, verticalAlignment,
            rotationDegrees, rotationCenterX, rotationCenterY, backgroundColor,
            backgroundPaddingX, backgroundPaddingY, centerLineInLineHeight,
            underlineOffsetFactor, strikethrough, fontFamily,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, OfficeTextBaseline.Normal);

    /// <summary>Draws a measured text-box plan with native decoration and baseline styling.</summary>
    /// <remarks><paramref name="underlineStyle"/> and <paramref name="strikethroughStyle"/> take precedence over their legacy Boolean switches; <paramref name="baseline"/> controls subscript and superscript placement.</remarks>
    public static void DrawRasterTextBox(
        OfficeRasterCanvas canvas,
        OfficeTextBlockRenderPlan plan,
        OfficeColor color,
        bool bold = false,
        bool italic = false,
        bool underline = false,
        OfficeTextAlignment? horizontalAlignment = null,
        OfficeTextVerticalAlignment? verticalAlignment = null,
        double rotationDegrees = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D,
        OfficeColor? backgroundColor = null,
        double backgroundPaddingX = 0D,
        double backgroundPaddingY = 0D,
        bool centerLineInLineHeight = true,
        double underlineOffsetFactor = 0.86D,
        bool strikethrough = false,
        string? fontFamily = null,
        OfficeTextDecorationStyle underlineStyle = OfficeTextDecorationStyle.None,
        OfficeTextDecorationStyle strikethroughStyle = OfficeTextDecorationStyle.None,
        OfficeTextBaseline baseline = OfficeTextBaseline.Normal) {
        if (canvas == null) {
            throw new ArgumentNullException(nameof(canvas));
        }

        if (plan == null) {
            throw new ArgumentNullException(nameof(plan));
        }

        if (backgroundColor.HasValue && backgroundColor.Value.A > 0) {
            OfficeTextBlockBackgroundBounds background = plan.CreateBackgroundBounds(backgroundPaddingX, backgroundPaddingY);
            if (Math.Abs(rotationDegrees) <= 0.000001D) {
                canvas.FillRectangle(background.Left, background.Top, background.Width, background.Height, backgroundColor.Value);
            } else {
                canvas.FillPolygon(background.GetRotatedCorners(rotationDegrees, rotationCenterX, rotationCenterY), backgroundColor.Value);
            }
        }

        DrawRasterTextBlock(
            canvas,
            plan.Layout,
            plan.Left,
            plan.Top,
            plan.Width,
            plan.Height,
            color,
            horizontalAlignment ?? plan.HorizontalAlignment,
            verticalAlignment ?? plan.VerticalAlignment,
            bold,
            italic,
            underline,
            rotationDegrees,
            rotationCenterX,
            rotationCenterY,
            centerLineInLineHeight,
            underlineOffsetFactor,
            strikethrough,
            fontFamily,
            underlineStyle: underlineStyle,
            strikethroughStyle: strikethroughStyle,
            baseline: baseline);
    }

    /// <summary>
    /// Draws a measured rich text block on a raster canvas.
    /// </summary>
    /// <param name="canvas">Raster canvas receiving the text.</param>
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
    /// <param name="flipHorizontal">Whether to mirror each rendered segment horizontally around the rotation center before rotation.</param>
    /// <param name="flipVertical">Whether to mirror each rendered segment vertically around the rotation center before rotation.</param>
    public static void DrawRasterRichTextBlock(
        OfficeRasterCanvas canvas,
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
        bool centerLineInLineHeight = true,
        bool flipHorizontal = false,
        bool flipVertical = false) {
        if (canvas == null) {
            throw new ArgumentNullException(nameof(canvas));
        }

        if (layout == null) {
            throw new ArgumentNullException(nameof(layout));
        }

        if (layout.Lines.Count == 0 || width <= 0D || height <= 0D) {
            return;
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
                DrawRasterJustifiedRichTextLine(
                    canvas,
                    line,
                    lineLeft,
                    lineWidth,
                    baseline,
                    rotationDegrees,
                    rotationCenterX,
                    rotationCenterY,
                    flipHorizontal,
                    flipVertical);
                lineTop += lineHeight;
                continue;
            }

            double cursor = OfficeTextPlacement.ResolveLineLeft(lineLeft, lineWidth, line.Width, horizontalAlignment);
            for (int segmentIndex = 0; segmentIndex < line.Segments.Count; segmentIndex++) {
                OfficeRichTextSegment segment = line.Segments[segmentIndex];
                double renderedFontSize = ResolveRichTextRenderedFontSize(segment);
                double renderedBaseline = ResolveRichTextRenderedBaseline(segment, baseline);
                double segmentTop = renderedBaseline - (renderedFontSize * 0.84D);
                DrawRasterRichTextSegmentBackground(canvas, segment, cursor, segmentTop, rotationDegrees, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
                if (segment.TabLinePaint != null) DrawRasterTabLineLeader(canvas, segment.TabLinePaint, cursor, renderedBaseline,
                    rotationDegrees, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
                canvas.DrawTextLine(
                    segment.Text,
                    cursor,
                    segmentTop,
                    renderedFontSize,
                    segment.Color,
                    segment.Bold,
                    segment.Italic,
                    OfficeTextAlignment.Left,
                    rotationDegrees,
                    rotationCenterX,
                    rotationCenterY,
                    segment.Underline,
                    segment.Strikethrough,
                    segment.FontFamily,
                    flipHorizontal,
                    flipVertical,
                    segment.UnderlineStyle,
                    segment.StrikethroughStyle);
                cursor += segment.Width;
            }

            lineTop += lineHeight;
        }
    }

    /// <summary>
    /// Appends one SVG <c>text</c> element for a measured rich text segment.
    /// </summary>
    /// <param name="builder">SVG markup builder.</param>
    /// <param name="segment">Measured rich text segment.</param>
    /// <param name="x">Resolved segment x-coordinate.</param>
    /// <param name="baseline">Resolved segment baseline y-coordinate.</param>
    /// <param name="rotationDegrees">Clockwise rotation in degrees.</param>
    /// <param name="rotationCenterX">Rotation center X coordinate.</param>
    /// <param name="rotationCenterY">Rotation center Y coordinate.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendSvgRichTextSegment(
        this StringBuilder builder,
        OfficeRichTextSegment segment,
        double x,
        double baseline,
        double rotationDegrees = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D) {
        if (builder == null) throw new ArgumentNullException(nameof(builder));
        if (segment == null) {
            throw new ArgumentNullException(nameof(segment));
        }

        if (segment.TabLinePaint != null) return AppendSvgTabLineLeader(builder, segment.TabLinePaint, x,
            ResolveRichTextRenderedBaseline(segment, baseline), rotationDegrees, rotationCenterX, rotationCenterY);

        bool linked = AppendSvgRichTextLinkStart(builder, segment.LinkUri);
        builder.AppendSvgTextElement(
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
            rotationDegrees,
            rotationCenterX,
            rotationCenterY,
            strikethrough: segment.Strikethrough,
            underlineStyle: segment.UnderlineStyle,
            strikethroughStyle: segment.StrikethroughStyle,
            baseline: segment.Baseline);
        if (linked) builder.Append("</a>");
        return builder;
    }

    private static void DrawRasterRichTextSegmentBackground(
        OfficeRasterCanvas canvas,
        OfficeRichTextSegment segment,
        double x,
        double top,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY,
        bool flipHorizontal,
        bool flipVertical) {
        if (!segment.BackgroundColor.HasValue || segment.BackgroundColor.Value.A == 0 || segment.Width <= 0D || segment.FontSize <= 0D) {
            return;
        }

        double height = ResolveRichTextSegmentBackgroundHeight(segment);
        if (Math.Abs(rotationDegrees) <= 0.000001D && !flipHorizontal && !flipVertical) {
            canvas.FillRectangle(x, top, segment.Width, height, segment.BackgroundColor.Value);
            return;
        }

        canvas.FillPolygon(
            CreateTransformedTextRectangle(x, top, segment.Width, height, rotationDegrees, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical),
            segment.BackgroundColor.Value);
    }

    private static StringBuilder AppendSvgRichTextSegmentBackground(
        this StringBuilder builder,
        OfficeRichTextSegment segment,
        double x,
        double baseline,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY) {
        if (!segment.BackgroundColor.HasValue || segment.BackgroundColor.Value.A == 0 || segment.Width <= 0D || segment.FontSize <= 0D) {
            return builder;
        }

        double renderedFontSize = ResolveRichTextRenderedFontSize(segment);
        double renderedBaseline = ResolveRichTextRenderedBaseline(segment, baseline);
        double top = renderedBaseline - (renderedFontSize * 0.84D);
        double height = Math.Max(1D, renderedFontSize * 1.05D);
        builder.Append("<rect")
            .AppendNumberAttribute("x", x)
            .AppendNumberAttribute("y", top)
            .AppendNumberAttribute("width", segment.Width)
            .AppendNumberAttribute("height", height);
        if (Math.Abs(rotationDegrees) > 0.000001D) {
            builder.AppendRotateTransformAttribute(rotationDegrees, rotationCenterX, rotationCenterY);
        }

        builder.AppendPaintAttribute("fill", segment.BackgroundColor.Value)
            .Append("/>");
        return builder;
    }

    internal static double ResolveRichTextSegmentBackgroundHeight(OfficeRichTextSegment segment) =>
        Math.Max(1D, ResolveRichTextRenderedFontSize(segment) * 1.05D);

    internal static double ResolveRichTextRenderedFontSize(OfficeRichTextSegment segment) =>
        segment.Baseline == OfficeTextBaseline.Normal ? segment.FontSize : segment.FontSize * 0.65D;

    internal static double ResolveRichTextRenderedBaseline(OfficeRichTextSegment segment, double baseline) =>
        segment.Baseline == OfficeTextBaseline.Superscript
            ? baseline - (segment.FontSize * 0.30D)
            : segment.Baseline == OfficeTextBaseline.Subscript ? baseline + (segment.FontSize * 0.15D) : baseline;

    private static IReadOnlyList<OfficePoint> CreateTransformedTextRectangle(
        double x,
        double y,
        double width,
        double height,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY,
        bool flipHorizontal,
        bool flipVertical) {
        double radians = OfficeGeometry.DegreesToRadians(rotationDegrees);
        return new[] {
            TransformTextRectanglePoint(new OfficePoint(x, y), radians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical),
            TransformTextRectanglePoint(new OfficePoint(x + width, y), radians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical),
            TransformTextRectanglePoint(new OfficePoint(x + width, y + height), radians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical),
            TransformTextRectanglePoint(new OfficePoint(x, y + height), radians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical)
        };
    }

    private static OfficePoint TransformTextRectanglePoint(
        OfficePoint point,
        double rotationRadians,
        double centerX,
        double centerY,
        bool flipHorizontal,
        bool flipVertical) {
        double x = flipHorizontal ? centerX - (point.X - centerX) : point.X;
        double y = flipVertical ? centerY - (point.Y - centerY) : point.Y;
        if (Math.Abs(rotationRadians) <= 0.000001D) {
            return new OfficePoint(x, y);
        }

        double dx = x - centerX;
        double dy = y - centerY;
        double cos = Math.Cos(rotationRadians);
        double sin = Math.Sin(rotationRadians);
        return new OfficePoint(
            centerX + (dx * cos) - (dy * sin),
            centerY + (dx * sin) + (dy * cos));
    }

    /// <summary>Writes an SVG text-box plan with native decoration and baseline styling.</summary>
    /// <remarks><paramref name="underlineStyle"/> and <paramref name="strikethroughStyle"/> take precedence over their legacy Boolean switches; <paramref name="baseline"/> controls subscript and superscript placement.</remarks>
    public static void WriteSvgTextBox(
        XmlWriter writer,
        OfficeTextBlockRenderPlan plan,
        OfficeColor color,
        string? fontFamily,
        bool bold = false,
        bool italic = false,
        bool underline = false,
        double rotationDegrees = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D,
        string? svgNamespace = null,
        OfficeColor? backgroundColor = null,
        double backgroundPaddingX = 0D,
        double backgroundPaddingY = 0D,
        Action<XmlWriter>? configureTextAttributes = null,
        Action<XmlWriter>? configureBackgroundAttributes = null,
        bool strikethrough = false,
        OfficeTextDecorationStyle underlineStyle = OfficeTextDecorationStyle.None,
        OfficeTextDecorationStyle strikethroughStyle = OfficeTextDecorationStyle.None,
        OfficeTextBaseline baseline = OfficeTextBaseline.Normal) {
        if (writer == null) {
            throw new ArgumentNullException(nameof(writer));
        }

        if (plan == null) {
            throw new ArgumentNullException(nameof(plan));
        }

        if (backgroundColor.HasValue && backgroundColor.Value.A > 0) {
            OfficeTextBlockBackgroundBounds background = plan.CreateBackgroundBounds(backgroundPaddingX, backgroundPaddingY);
            writer.WriteStartElement("rect", svgNamespace);
            configureBackgroundAttributes?.Invoke(writer);
            writer.WriteNumberAttribute("x", background.Left);
            writer.WriteNumberAttribute("y", background.Top);
            writer.WriteNumberAttribute("width", background.Width);
            writer.WriteNumberAttribute("height", background.Height);
            if (Math.Abs(rotationDegrees) > 0.000001D) {
                writer.WriteRotateTransformAttribute(rotationDegrees, rotationCenterX, rotationCenterY);
            }

            OfficeSvgFormatting.WriteColorAttribute(writer, "fill", backgroundColor.Value);
            writer.WriteEndElement();
        }

        WriteSvgTextBlock(
            writer,
            plan.Layout,
            plan.Left,
            plan.Top,
            plan.Width,
            plan.Height,
            color,
            fontFamily,
            plan.HorizontalAlignment,
            plan.VerticalAlignment,
            bold,
            italic,
            underline,
            rotationDegrees,
            rotationCenterX,
            rotationCenterY,
            svgNamespace,
            configureTextAttributes,
            strikethrough,
            underlineStyle,
            strikethroughStyle,
            baseline);
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

    private static string GetSvgTextAnchor(OfficeTextAlignment alignment, OfficeTextDirection direction) {
        if (direction != OfficeTextDirection.RightToLeft) return GetSvgTextAnchor(alignment);
        switch (alignment) {
            case OfficeTextAlignment.Left:
                return "end";
            case OfficeTextAlignment.Right:
                return "start";
            default:
                return GetSvgTextAnchor(alignment);
        }
    }

    private static bool ShouldJustifyLine(OfficeTextLine line, int lineIndex, int lineCount, double availableWidth, OfficeTextAlignment alignment) {
        return alignment == OfficeTextAlignment.Justify &&
            lineIndex < lineCount - 1 &&
            availableWidth > line.Width + 0.01D &&
            CountJustifiableWords(line.Text) > 1;
    }

    private static void DrawRasterJustifiedTextLine(
        OfficeRasterCanvas canvas,
        string text,
        double left,
        double availableWidth,
        double top,
        double fontSize,
        OfficeColor color,
        bool bold,
        bool italic,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY,
        bool underline,
        bool strikethrough,
        string? fontFamily,
        bool flipHorizontal,
        bool flipVertical,
        OfficeTextDecorationStyle underlineStyle = OfficeTextDecorationStyle.None,
        OfficeTextDecorationStyle strikethroughStyle = OfficeTextDecorationStyle.None,
        OfficeColor? decorationColor = null) {
        string[] words = SplitJustifiableWords(text);
        if (words.Length <= 1) {
            canvas.DrawTextLine(text, left, top, fontSize, color, bold, italic, OfficeTextAlignment.Left, rotationDegrees, rotationCenterX, rotationCenterY, underline, strikethrough, fontFamily, flipHorizontal, flipVertical, underlineStyle, strikethroughStyle, decorationColor);
            return;
        }

        double wordsWidth = 0D;
        var widths = new double[words.Length];
        for (int i = 0; i < words.Length; i++) {
            widths[i] = canvas.MeasureText(words[i], fontSize, fontFamily);
            wordsWidth += widths[i];
        }

        double gap = Math.Max(0D, (availableWidth - wordsWidth) / Math.Max(1, words.Length - 1));
        double cursor = left;
        for (int i = 0; i < words.Length; i++) {
            canvas.DrawTextLine(words[i], cursor, top, fontSize, color, bold, italic, OfficeTextAlignment.Left, rotationDegrees, rotationCenterX, rotationCenterY, underline, strikethrough, fontFamily, flipHorizontal, flipVertical, underlineStyle, strikethroughStyle, decorationColor);
            cursor += widths[i] + gap;
        }
    }

    private static int CountJustifiableWords(string text) => SplitJustifiableWords(text).Length;

    private static string[] SplitJustifiableWords(string text) {
        if (string.IsNullOrWhiteSpace(text)) {
            return Array.Empty<string>();
        }

        return text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
    }

    internal static double ResolveRichTextRenderLineHeight(OfficeRichTextLine line, double fallbackLineHeight) =>
        line.LineHeight > 0D ? line.LineHeight : fallbackLineHeight;

    internal static double ResolveRichTextRenderBaseline(
        OfficeRichTextLine line,
        double lineTop,
        double lineHeight,
        bool centerLineInLineHeight) {
        OfficeTextLayoutEngine.ResolveRichTextVerticalExtents(line.Segments, out double contentTop, out double contentBottom);
        double contentHeight = Math.Max(0D, contentBottom - contentTop);
        double leading = centerLineInLineHeight ? Math.Max(0D, lineHeight - contentHeight) / 2D : 0D;
        return lineTop + leading - contentTop;
    }

    private static bool RequiresSvgWhitespacePreserve(OfficeTextBlockLayout layout) {
        for (int i = 0; i < layout.Lines.Count; i++) {
            if (RequiresSvgWhitespacePreserve(layout.Lines[i].Text)) {
                return true;
            }
        }

        return false;
    }

    private static bool RequiresSvgWhitespacePreserve(string text) {
        if (string.IsNullOrEmpty(text)) {
            return false;
        }

        if (char.IsWhiteSpace(text[0]) || char.IsWhiteSpace(text[text.Length - 1])) {
            return true;
        }

        for (int i = 1; i < text.Length; i++) {
            if (char.IsWhiteSpace(text[i]) && char.IsWhiteSpace(text[i - 1])) {
                return true;
            }
        }

        return false;
    }

    private static void AppendSvgTextDecorationAttribute(StringBuilder builder, bool underline, bool strikethrough) {
        AppendSvgTextDecorationAttribute(
            builder,
            underline ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None,
            strikethrough ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None);
    }

    private static bool RequiresSeparateSvgDecorations(OfficeTextDecorationStyle underlineStyle, OfficeTextDecorationStyle strikethroughStyle) =>
        underlineStyle != OfficeTextDecorationStyle.None &&
        strikethroughStyle != OfficeTextDecorationStyle.None &&
        underlineStyle != strikethroughStyle;

    private static void AppendSvgTextDecorationAttribute(StringBuilder builder, OfficeTextDecorationStyle underlineStyle, OfficeTextDecorationStyle strikethroughStyle, OfficeColor? decorationColor = null) {
        bool underline = underlineStyle != OfficeTextDecorationStyle.None;
        bool strikethrough = strikethroughStyle != OfficeTextDecorationStyle.None;
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
        if (decorationColor.HasValue) {
            builder.AppendPaintAttribute("text-decoration-color", decorationColor.Value);
        }
        OfficeTextDecorationStyle pattern = underline ? underlineStyle : strikethroughStyle;
        if (pattern == OfficeTextDecorationStyle.Single) {
            return;
        }
        string svgStyle = pattern switch {
            OfficeTextDecorationStyle.Double => "double",
            OfficeTextDecorationStyle.Dotted => "dotted",
            OfficeTextDecorationStyle.Dashed => "dashed",
            OfficeTextDecorationStyle.Wavy => "wavy",
            _ => "solid"
        };
        builder.AppendAttribute("text-decoration-style", svgStyle);
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

    private static void WriteSvgTextDecorationAttribute(XmlWriter writer, OfficeTextDecorationStyle underlineStyle, OfficeTextDecorationStyle strikethroughStyle) {
        bool underline = underlineStyle != OfficeTextDecorationStyle.None;
        bool strikethrough = strikethroughStyle != OfficeTextDecorationStyle.None;
        if (!underline && !strikethrough) return;

        writer.WriteAttributeString("text-decoration", underline && strikethrough
            ? "underline line-through"
            : underline ? "underline" : "line-through");
        OfficeTextDecorationStyle pattern = underline ? underlineStyle : strikethroughStyle;
        if (pattern == OfficeTextDecorationStyle.Single) return;
        writer.WriteAttributeString("text-decoration-style", pattern switch {
            OfficeTextDecorationStyle.Double => "double",
            OfficeTextDecorationStyle.Dotted => "dotted",
            OfficeTextDecorationStyle.Dashed => "dashed",
            OfficeTextDecorationStyle.Wavy => "wavy",
            _ => "solid"
        });
    }
}
