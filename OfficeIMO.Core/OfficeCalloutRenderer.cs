using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Drawing;

/// <summary>
/// Shared dependency-free renderer for annotation/comment-style callouts.
/// </summary>
public static class OfficeCalloutRenderer {
    /// <summary>
    /// Draws a callout to a raster canvas.
    /// </summary>
    public static void DrawRaster(OfficeRasterCanvas canvas, OfficeCallout callout, OfficeCalloutStyle style, double scale = 1D) {
        if (canvas == null) {
            throw new ArgumentNullException(nameof(canvas));
        }

        if (callout == null) {
            throw new ArgumentNullException(nameof(callout));
        }

        if (style == null) {
            throw new ArgumentNullException(nameof(style));
        }

        if (scale <= 0D || callout.Width <= 0D || callout.Height <= 0D) {
            return;
        }

        double x = callout.X * scale;
        double y = callout.Y * scale;
        double width = callout.Width * scale;
        double height = callout.Height * scale;
        double headerHeight = Math.Min(height, style.HeaderHeight * scale);

        DrawRasterBodyShadow(canvas, style, scale, x, y, width, height);
        DrawRasterPointer(canvas, callout, style, scale, shadow: true);
        DrawRasterPointer(canvas, callout, style, scale, shadow: false);
        OfficeDrawing drawing = CreateBodyDrawing(style, scale, width, height, headerHeight);
        OfficeRasterImage bodyImage = OfficeDrawingRasterRenderer.Render(drawing);
        canvas.DrawImage(bodyImage, x, y, width, height);
        DrawRasterText(canvas, callout, style, scale, x, y, width, height);
    }

    /// <summary>
    /// Appends SVG elements for a callout.
    /// </summary>
    /// <param name="builder">SVG builder.</param>
    /// <param name="callout">Callout geometry and text.</param>
    /// <param name="style">Callout style.</param>
    /// <param name="measureText">Text measurement delegate used for wrapping body text.</param>
    /// <param name="scale">Renderer scale.</param>
    /// <param name="idPrefix">Stable SVG id prefix for clip paths.</param>
    /// <returns>The supplied builder for call chaining.</returns>
    public static StringBuilder AppendSvg(
        StringBuilder builder,
        OfficeCallout callout,
        OfficeCalloutStyle style,
        Func<string?, double, string?, double> measureText,
        double scale = 1D,
        string idPrefix = "office-callout") {
        if (builder == null) {
            throw new ArgumentNullException(nameof(builder));
        }

        if (callout == null) {
            throw new ArgumentNullException(nameof(callout));
        }

        if (style == null) {
            throw new ArgumentNullException(nameof(style));
        }

        if (measureText == null) {
            throw new ArgumentNullException(nameof(measureText));
        }

        if (scale <= 0D || callout.Width <= 0D || callout.Height <= 0D) {
            return builder;
        }

        double x = callout.X * scale;
        double y = callout.Y * scale;
        double width = callout.Width * scale;
        double height = callout.Height * scale;
        string safePrefix = string.IsNullOrWhiteSpace(idPrefix) ? "office-callout" : idPrefix;

        AppendSvgBodyShadow(builder, style, scale, x, y, width, height);
        AppendSvgPointer(builder, callout, style, scale, shadow: true);
        AppendSvgPointer(builder, callout, style, scale, shadow: false);
        builder.AppendNestedSvgStart(x, y, width, height);
        AppendSvgBody(builder, style, scale, width, height, safePrefix);
        AppendSvgText(builder, callout, style, scale, width, height, measureText, safePrefix);
        builder.AppendNestedSvgEnd();
        return builder;
    }

    private static void DrawRasterPointer(OfficeRasterCanvas canvas, OfficeCallout callout, OfficeCalloutStyle style, double scale, bool shadow) {
        IReadOnlyList<OfficePoint> points = GetPointerPoints(callout, style, scale, shadow);
        if (points.Count < 3) {
            return;
        }

        if (shadow) {
            canvas.FillPolygon(points, style.ShadowColor);
            return;
        }

        canvas.FillPolygon(points, style.FillColor);
        canvas.DrawPolygon(points, style.StrokeColor, Math.Max(1D, scale));
    }

    private static void AppendSvgPointer(StringBuilder builder, OfficeCallout callout, OfficeCalloutStyle style, double scale, bool shadow) {
        IReadOnlyList<OfficePoint> points = GetPointerPoints(callout, style, scale, shadow);
        if (points.Count < 3) {
            return;
        }

        builder.AppendPolygonElement(
            points,
            shadow ? style.ShadowColor : style.FillColor,
            shadow ? null : style.StrokeColor,
            shadow ? 0D : Math.Max(1D, scale));
    }

    private static IReadOnlyList<OfficePoint> GetPointerPoints(OfficeCallout callout, OfficeCalloutStyle style, double scale, bool shadow) {
        double x = callout.X * scale;
        double y = callout.Y * scale;
        double width = callout.Width * scale;
        double height = callout.Height * scale;
        double anchorX = callout.AnchorX * scale;
        double anchorY = callout.AnchorY * scale;
        double offset = shadow ? 2D * scale : 0D;
        double headerHeight = style.HeaderHeight * scale;
        double inset = Math.Min(height - (8D * scale), Math.Max(10D * scale, headerHeight * 0.65D));
        double half = Math.Max(5D * scale, Math.Min(10D * scale, height / 7D));

        if (anchorX <= x) {
            double sideY = Math.Min(y + height - half - (4D * scale), Math.Max(y + half + (4D * scale), y + inset));
            return new[] {
                new OfficePoint(anchorX + offset, anchorY + offset),
                new OfficePoint(x + offset, sideY - half + offset),
                new OfficePoint(x + offset, sideY + half + offset)
            };
        }

        if (anchorX >= x + width) {
            double sideY = Math.Min(y + height - half - (4D * scale), Math.Max(y + half + (4D * scale), y + inset));
            return new[] {
                new OfficePoint(anchorX + offset, anchorY + offset),
                new OfficePoint(x + width + offset, sideY - half + offset),
                new OfficePoint(x + width + offset, sideY + half + offset)
            };
        }

        if (anchorY <= y) {
            double sideX = Math.Min(x + width - half - (4D * scale), Math.Max(x + half + (4D * scale), anchorX));
            return new[] {
                new OfficePoint(anchorX + offset, anchorY + offset),
                new OfficePoint(sideX - half + offset, y + offset),
                new OfficePoint(sideX + half + offset, y + offset)
            };
        }

        double bottomSideX = Math.Min(x + width - half - (4D * scale), Math.Max(x + half + (4D * scale), anchorX));
        return new[] {
            new OfficePoint(anchorX + offset, anchorY + offset),
            new OfficePoint(bottomSideX - half + offset, y + height + offset),
            new OfficePoint(bottomSideX + half + offset, y + height + offset)
        };
    }

    private static void DrawRasterText(OfficeRasterCanvas canvas, OfficeCallout callout, OfficeCalloutStyle style, double scale, double x, double y, double width, double height) {
        double padding = style.Padding * scale;
        double titleHeight = style.HeaderHeight * scale;
        double titleFontSize = style.TitleFontSize * scale;
        double bodyFontSize = style.TextFontSize * scale;
        double textX = x + padding + (2D * scale);
        double textWidth = Math.Max(1D, width - (padding * 2D) - (2D * scale));

        using (canvas.PushClipRectangle(x, y, width, height)) {
            canvas.DrawTextLine(
                callout.Title,
                textX,
                y + Math.Max(2D * scale, (titleHeight - titleFontSize) / 2D),
                titleFontSize,
                style.TitleColor,
                bold: true,
                alignment: OfficeTextAlignment.Left,
                fontFamily: style.FontFamily);

        double bodyTop = y + titleHeight + (4D * scale);
        double bodyHeight = Math.Max(1D, height - titleHeight - (padding * 1.4D));
        if (callout.RichTextRuns.Count > 0) {
            IReadOnlyList<OfficeRichTextRun> scaledRuns = ScaleRichTextRuns(callout.RichTextRuns, scale);
            OfficeRichTextBlockLayout richLayout = OfficeTextLayoutEngine.LayoutRichTextBlock(
                scaledRuns,
                textWidth,
                bodyHeight,
                style.LineHeightFactor,
                (text, size, family) => canvas.MeasureText(text, size, family),
                wrap: true,
                shrinkToFit: false,
                minimumFontSize: Math.Max(1D, scale),
                overflowBehavior: OfficeTextOverflowBehavior.Clip);
            OfficeTextBlockRenderer.DrawRasterRichTextBlock(
                canvas,
                richLayout,
                textX,
                bodyTop,
                textWidth,
                bodyHeight,
                OfficeTextAlignment.Left,
                OfficeTextVerticalAlignment.Top,
                centerLineInLineHeight: false);
            return;
        }

        OfficeTextBlockLayout layout = OfficeTextLayoutEngine.LayoutTextBlock(
            callout.Text,
            bodyFontSize,
                textWidth,
                bodyHeight,
                style.LineHeightFactor,
                Math.Max(6D * scale, scale),
                (text, size) => canvas.MeasureText(text, size, style.FontFamily),
                wrap: true,
                forceSingleLine: false,
                shrinkToFit: false);
            OfficeTextBlockRenderer.DrawRasterTextBlock(
                canvas,
                layout,
                textX,
                bodyTop,
                textWidth,
                bodyHeight,
                style.TextColor,
                OfficeTextAlignment.Left,
                OfficeTextVerticalAlignment.Top,
                centerLineInLineHeight: false,
                fontFamily: style.FontFamily);
        }
    }

    private static void AppendSvgBody(StringBuilder builder, OfficeCalloutStyle style, double scale, double width, double height, string idPrefix) {
        double radius = ResolveCornerRadius(width, height, scale);
        double headerHeight = Math.Min(height, style.HeaderHeight * scale);
        string clipId = idPrefix + "-body";

        builder.AppendRectClipPathDefinition(clipId, 0D, 0D, width, height);
        builder.AppendRectElement(
            0D,
            0D,
            width,
            height,
            radius,
            radius,
            new StringBuilder()
                .AppendPaintAttribute("fill", style.FillColor)
                .AppendPaintAttribute("stroke", style.StrokeColor)
                .AppendNumberAttribute("stroke-width", Math.Max(1D, scale))
                .ToString());
        builder.Append("<g").AppendClipPathReference(clipId).Append(">");
        builder.AppendRectElement(0D, 0D, width, headerHeight, new StringBuilder().AppendPaintAttribute("fill", style.HeaderFillColor).ToString());
        builder.AppendRectElement(0D, 0D, Math.Max(2D, 3D * scale), height, new StringBuilder().AppendPaintAttribute("fill", style.AccentColor).ToString());
        builder.Append("</g>");
    }

    private static void DrawRasterBodyShadow(OfficeRasterCanvas canvas, OfficeCalloutStyle style, double scale, double x, double y, double width, double height) {
        if (style.ShadowColor.A == 0) {
            return;
        }

        double spread = Math.Max(0D, style.ShadowSpread * scale);
        double offsetX = style.ShadowOffsetX * scale;
        double offsetY = style.ShadowOffsetY * scale;
        double radius = ResolveCornerRadius(width, height, scale);

        if (spread > 0D) {
            OfficeDrawing ambient = CreateShadowDrawing(ScaleAlpha(style.ShadowColor, 0.45D), width + (spread * 2D), height + (spread * 2D), radius + spread);
            OfficeRasterImage ambientImage = OfficeDrawingRasterRenderer.Render(ambient);
            canvas.DrawImage(ambientImage, x + offsetX - spread, y + offsetY - spread, width + (spread * 2D), height + (spread * 2D));
        }

        OfficeDrawing shadow = CreateShadowDrawing(style.ShadowColor, width, height, radius);
        OfficeRasterImage shadowImage = OfficeDrawingRasterRenderer.Render(shadow);
        canvas.DrawImage(shadowImage, x + offsetX, y + offsetY, width, height);
    }

    private static void AppendSvgBodyShadow(StringBuilder builder, OfficeCalloutStyle style, double scale, double x, double y, double width, double height) {
        if (style.ShadowColor.A == 0) {
            return;
        }

        double spread = Math.Max(0D, style.ShadowSpread * scale);
        double offsetX = style.ShadowOffsetX * scale;
        double offsetY = style.ShadowOffsetY * scale;
        double radius = ResolveCornerRadius(width, height, scale);

        if (spread > 0D) {
            builder.AppendRectElement(
                x + offsetX - spread,
                y + offsetY - spread,
                width + (spread * 2D),
                height + (spread * 2D),
                radius + spread,
                radius + spread,
                new StringBuilder().AppendPaintAttribute("fill", ScaleAlpha(style.ShadowColor, 0.45D)).ToString());
        }

        builder.AppendRectElement(
            x + offsetX,
            y + offsetY,
            width,
            height,
            radius,
            radius,
            new StringBuilder().AppendPaintAttribute("fill", style.ShadowColor).ToString());
    }

    private static OfficeDrawing CreateBodyDrawing(OfficeCalloutStyle style, double scale, double width, double height, double headerHeight) {
        double radius = ResolveCornerRadius(width, height, scale);
        var drawing = new OfficeDrawing(width, height);

        OfficeShape background = OfficeShape.RoundedRectangle(width, height, radius);
        background.FillColor = style.FillColor;
        background.StrokeColor = style.StrokeColor;
        background.StrokeWidth = Math.Max(1D, scale);
        drawing.AddShape(background, 0D, 0D);

        OfficeShape header = OfficeShape.Rectangle(width, headerHeight);
        header.FillColor = style.HeaderFillColor;
        header.StrokeColor = null;
        header.ClipPath = OfficeClipPath.RoundedRectangle(width, height, radius);
        drawing.AddShape(header, 0D, 0D);

        OfficeShape accent = OfficeShape.Rectangle(Math.Max(2D, 3D * scale), height);
        accent.FillColor = style.AccentColor;
        accent.StrokeColor = null;
        accent.ClipPath = OfficeClipPath.RoundedRectangle(width, height, radius);
        drawing.AddShape(accent, 0D, 0D);

        return drawing;
    }

    private static OfficeDrawing CreateShadowDrawing(OfficeColor shadowColor, double width, double height, double radius) {
        var drawing = new OfficeDrawing(width, height);
        OfficeShape shadow = OfficeShape.RoundedRectangle(width, height, radius);
        shadow.FillColor = shadowColor;
        shadow.StrokeColor = null;
        drawing.AddShape(shadow, 0D, 0D);
        return drawing;
    }

    private static double ResolveCornerRadius(double width, double height, double scale) =>
        Math.Min(7D * scale, Math.Min(width, height) / 7D);

    private static OfficeColor ScaleAlpha(OfficeColor color, double factor) {
        double clamped = Math.Max(0D, Math.Min(1D, factor));
        return OfficeColor.FromRgba(color.R, color.G, color.B, (byte)Math.Round(color.A * clamped));
    }

    private static void AppendSvgText(
        StringBuilder builder,
        OfficeCallout callout,
        OfficeCalloutStyle style,
        double scale,
        double width,
        double height,
        Func<string?, double, string?, double> measureText,
        string idPrefix) {
        double padding = style.Padding * scale;
        double titleHeight = style.HeaderHeight * scale;
        double titleFontSize = style.TitleFontSize * scale;
        double bodyFontSize = style.TextFontSize * scale;
        double textX = padding + (2D * scale);
        double textWidth = Math.Max(1D, width - (padding * 2D) - (2D * scale));
        string clipId = idPrefix + "-text";

        builder.AppendRectClipPathDefinition(clipId, 0D, 0D, width, height);
        builder.Append("<g").AppendClipPathReference(clipId).Append(">");
        builder.AppendSvgTextElement(
            callout.Title,
            textX,
            Math.Max(titleFontSize + (3D * scale), titleHeight - (5D * scale)),
            titleFontSize,
            style.TitleColor,
            style.FontFamily,
            titleFontSize,
            OfficeTextAlignment.Left,
            bold: true);

        double bodyTop = titleHeight + (4D * scale);
        double bodyHeight = Math.Max(1D, height - titleHeight - (padding * 1.4D));
        if (callout.RichTextRuns.Count > 0) {
            IReadOnlyList<OfficeRichTextRun> scaledRuns = ScaleRichTextRuns(callout.RichTextRuns, scale);
            OfficeRichTextBlockLayout richLayout = OfficeTextLayoutEngine.LayoutRichTextBlock(
                scaledRuns,
                textWidth,
                bodyHeight,
                style.LineHeightFactor,
                measureText,
                wrap: true,
                shrinkToFit: false,
                minimumFontSize: Math.Max(1D, scale),
                overflowBehavior: OfficeTextOverflowBehavior.Clip);
            builder.AppendSvgRichTextBlock(
                richLayout,
                textX,
                bodyTop,
                textWidth,
                bodyHeight,
                OfficeTextAlignment.Left,
                OfficeTextVerticalAlignment.Top,
                centerLineInLineHeight: false);
            builder.Append("</g>");
            return;
        }

        OfficeTextBlockLayout layout = OfficeTextLayoutEngine.LayoutTextBlock(
            callout.Text,
            bodyFontSize,
            textWidth,
            bodyHeight,
            style.LineHeightFactor,
            Math.Max(6D * scale, scale),
            (text, size) => measureText(text, size, style.FontFamily),
            wrap: true,
            forceSingleLine: false,
            shrinkToFit: false);
        builder.AppendSvgTextBlock(
            layout,
            textX,
            bodyTop,
            textWidth,
            bodyHeight,
            style.TextColor,
            style.FontFamily,
            OfficeTextAlignment.Left,
            OfficeTextVerticalAlignment.Top,
            centerLineInLineHeight: false);

        builder.Append("</g>");
    }

    private static IReadOnlyList<OfficeRichTextRun> ScaleRichTextRuns(IReadOnlyList<OfficeRichTextRun> runs, double scale) {
        if (runs.Count == 0 || Math.Abs(scale - 1D) < 0.000001D) {
            return runs;
        }

        var scaled = new List<OfficeRichTextRun>(runs.Count);
        for (int i = 0; i < runs.Count; i++) {
            OfficeRichTextRun run = runs[i];
            scaled.Add(new OfficeRichTextRun(
                run.Text,
                run.FontSize * scale,
                run.Color,
                run.Bold,
                run.Italic,
                run.Underline,
                run.FontFamily,
                run.Strikethrough,
                run.BackgroundColor));
        }

        return scaled.AsReadOnly();
    }
}
