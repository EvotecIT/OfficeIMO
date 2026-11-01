using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    /// <summary>
    /// Draws a single anchored text line with optional bold, italic, alignment, and rotation.
    /// </summary>
    /// <remarks>
    /// This overload retains the original CLR signature for applications compiled against earlier OfficeIMO.Core releases.
    /// </remarks>
    /// <param name="text">Text to render.</param>
    /// <param name="anchorX">Horizontal anchor. Left, center, or right interpretation depends on <paramref name="alignment"/>.</param>
    /// <param name="top">Top coordinate of the text line box.</param>
    /// <param name="height">Text line height in canvas pixels.</param>
    /// <param name="color">Text color.</param>
    /// <param name="bold">Whether to simulate bold rendering.</param>
    /// <param name="italic">Whether to simulate italic rendering.</param>
    /// <param name="alignment">Horizontal alignment relative to <paramref name="anchorX"/>.</param>
    /// <param name="rotationDegrees">Clockwise rotation in degrees.</param>
    /// <param name="rotationCenterX">Rotation center X coordinate.</param>
    /// <param name="rotationCenterY">Rotation center Y coordinate.</param>
    /// <param name="underline">Whether to draw an underline using the measured text width.</param>
    /// <param name="strikethrough">Whether to draw a strikethrough using the measured text width.</param>
    /// <param name="fontFamily">Requested font family fallback list.</param>
    /// <param name="flipHorizontal">Whether to mirror the rendered line horizontally around the rotation center before rotation.</param>
    /// <param name="flipVertical">Whether to mirror the rendered line vertically around the rotation center before rotation.</param>
    public void DrawTextLine(
        string? text,
        double anchorX,
        double top,
        double height,
        OfficeColor color,
        bool bold = false,
        bool italic = false,
        OfficeTextAlignment alignment = OfficeTextAlignment.Center,
        double rotationDegrees = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D,
        bool underline = false,
        bool strikethrough = false,
        string? fontFamily = null,
        bool flipHorizontal = false,
        bool flipVertical = false) =>
        DrawTextLine(
            text,
            anchorX,
            top,
            height,
            color,
            bold,
            italic,
            alignment,
            rotationDegrees,
            rotationCenterX,
            rotationCenterY,
            underline,
            strikethrough,
            fontFamily,
            flipHorizontal,
            flipVertical,
            OfficeTextDecorationStyle.None,
            OfficeTextDecorationStyle.None);

    /// <summary>
    /// Draws a single anchored text line with independent underline and strikethrough patterns.
    /// </summary>
    /// <param name="text">Text to render.</param>
    /// <param name="anchorX">Horizontal anchor. Left, center, or right interpretation depends on <paramref name="alignment"/>.</param>
    /// <param name="top">Top coordinate of the text line box.</param>
    /// <param name="height">Text line height in canvas pixels.</param>
    /// <param name="color">Text color.</param>
    /// <param name="bold">Whether to simulate bold rendering.</param>
    /// <param name="italic">Whether to simulate italic rendering.</param>
    /// <param name="alignment">Horizontal alignment relative to <paramref name="anchorX"/>.</param>
    /// <param name="rotationDegrees">Clockwise rotation in degrees.</param>
    /// <param name="rotationCenterX">Rotation center X coordinate.</param>
    /// <param name="rotationCenterY">Rotation center Y coordinate.</param>
    /// <param name="underline">Whether to draw an underline using the measured text width.</param>
    /// <param name="strikethrough">Whether to draw a strikethrough using the measured text width.</param>
    /// <param name="fontFamily">Requested font family fallback list.</param>
    /// <param name="flipHorizontal">Whether to mirror the rendered line horizontally around the rotation center before rotation.</param>
    /// <param name="flipVertical">Whether to mirror the rendered line vertically around the rotation center before rotation.</param>
    /// <param name="underlineStyle">Underline pattern. A non-none value takes precedence over <paramref name="underline"/>.</param>
    /// <param name="strikethroughStyle">Strikethrough pattern. A non-none value takes precedence over <paramref name="strikethrough"/>.</param>
    /// <param name="decorationColor">Optional underline and strikethrough color. Null uses <paramref name="color"/>.</param>
    public void DrawTextLine(
        string? text,
        double anchorX,
        double top,
        double height,
        OfficeColor color,
        bool bold,
        bool italic,
        OfficeTextAlignment alignment,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY,
        bool underline,
        bool strikethrough,
        string? fontFamily,
        bool flipHorizontal,
        bool flipVertical,
        OfficeTextDecorationStyle underlineStyle,
        OfficeTextDecorationStyle strikethroughStyle,
        OfficeColor? decorationColor = null) =>
        DrawTextLineWithFeatures(null, text, anchorX, top, height, color, bold, italic, alignment, rotationDegrees, rotationCenterX, rotationCenterY, underline, strikethrough, fontFamily, flipHorizontal, flipVertical, underlineStyle, strikethroughStyle, decorationColor);

    internal void DrawTextLineWithFeatures(
        OfficeTextFeatureSettings? features,
        string? text,
        double anchorX,
        double top,
        double height,
        OfficeColor color,
        bool bold,
        bool italic,
        OfficeTextAlignment alignment,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY,
        bool underline,
        bool strikethrough,
        string? fontFamily,
        bool flipHorizontal,
        bool flipVertical,
        OfficeTextDecorationStyle underlineStyle,
        OfficeTextDecorationStyle strikethroughStyle,
        OfficeColor? decorationColor = null) {
        if (string.IsNullOrEmpty(text) || color.A == 0 || height <= 0D) {
            return;
        }

        string value = text!;
        double fontHeight = Math.Max(1D, height);
        OfficeFontStyle fontStyle = (bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
            | (italic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
        if (TryDrawMixedTextLine(
            value,
            anchorX,
            top,
            fontHeight,
            color,
            bold,
            italic,
            alignment,
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
            decorationColor, features)) {
            return;
        }
        IOfficeFontProgram? font = ResolveTextFont(value, fontFamily, fontStyle, fontHeight, out OfficeFontStyle resolvedStyle);
        bool simulateBold = bold && (resolvedStyle & OfficeFontStyle.Bold) != OfficeFontStyle.Bold;
        bool simulateItalic = italic && (resolvedStyle & OfficeFontStyle.Italic) != OfficeFontStyle.Italic;
        double width = MeasureTextWithFeatures(value, fontHeight, fontFamily, fontStyle, features);
        double x = ResolveAnchoredTextX(anchorX, width, alignment);
        double rotationRadians = OfficeGeometry.DegreesToRadians(rotationDegrees);
        if (font != null) {
            double outlineTop = ResolveRasterTextLineOutlineTop(font, fontHeight, top);
            double bottom = top + fontHeight;
            IReadOnlyList<List<OfficePoint>> contours = TransformTextContours(
                GetResolvedTextContours(value, font, x, outlineTop, fontHeight, features),
                bottom,
                simulateItalic,
                rotationRadians,
                rotationCenterX,
                rotationCenterY,
                flipHorizontal,
                flipVertical);
            if (simulateBold) {
                var shifted = TransformTextContours(
                    GetResolvedTextContours(value, font, x + fontHeight / 24D, outlineTop, fontHeight, features),
                    bottom,
                    simulateItalic,
                    rotationRadians,
                    rotationCenterX,
                    rotationCenterY,
                    flipHorizontal,
                    flipVertical);
                FillTextContourUnion(contours, shifted, color);
            } else {
                FillTextContours(contours, color);
            }

            DrawTextLineDecorations(x, width, top, fontHeight, decorationColor ?? color, rotationRadians, rotationCenterX, rotationCenterY, underlineStyle != OfficeTextDecorationStyle.None ? underlineStyle : underline ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None, strikethroughStyle != OfficeTextDecorationStyle.None ? strikethroughStyle : strikethrough ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None, flipHorizontal, flipVertical);
            return;
        }

        if (_textInkObserver != null) { _textInkObserver(null); return; }
        DrawStrokeText(value, anchorX, top + (fontHeight / 2D), fontHeight, color, bold, italic, alignment, rotationRadians, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
        DrawTextLineDecorations(x, width, top, fontHeight, decorationColor ?? color, rotationRadians, rotationCenterX, rotationCenterY, underlineStyle != OfficeTextDecorationStyle.None ? underlineStyle : underline ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None, strikethroughStyle != OfficeTextDecorationStyle.None ? strikethroughStyle : strikethrough ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None, flipHorizontal, flipVertical);
    }

    /// <summary>
    /// Draws a single text line through an arbitrary affine transform.
    /// </summary>
    /// <param name="text">Text to draw.</param>
    /// <param name="anchorX">Untransformed text anchor X coordinate.</param>
    /// <param name="top">Untransformed top coordinate.</param>
    /// <param name="height">Untransformed text height.</param>
    /// <param name="color">Text fill color.</param>
    /// <param name="transform">Affine transform applied to text contours.</param>
    /// <param name="bold">Whether to draw a bold approximation.</param>
    /// <param name="italic">Whether to skew text contours before applying the transform.</param>
    /// <param name="alignment">Anchor alignment.</param>
    /// <param name="underline">Whether to draw an underline.</param>
    /// <param name="strikethrough">Whether to draw a strikethrough.</param>
    /// <param name="fontFamily">Requested font family fallback list.</param>
    public void DrawTextLineTransformed(
        string? text,
        double anchorX,
        double top,
        double height,
        OfficeColor color,
        OfficeTransform transform,
        bool bold = false,
        bool italic = false,
        OfficeTextAlignment alignment = OfficeTextAlignment.Center,
        bool underline = false,
        bool strikethrough = false,
        string? fontFamily = null) {
        if (string.IsNullOrEmpty(text) || color.A == 0 || height <= 0D) {
            return;
        }

        string value = text!;
        double fontHeight = Math.Max(1D, height);
        OfficeFontStyle fontStyle = (bold ? OfficeFontStyle.Bold : OfficeFontStyle.Regular)
            | (italic ? OfficeFontStyle.Italic : OfficeFontStyle.Regular);
        if (TryDrawMixedTransformedTextLine(
            value,
            anchorX,
            top,
            fontHeight,
            color,
            transform,
            bold,
            italic,
            alignment,
            underline,
            strikethrough,
            fontFamily)) {
            return;
        }
        IOfficeFontProgram? font = ResolveTextFont(value, fontFamily, fontStyle, fontHeight, out OfficeFontStyle resolvedStyle);
        bool simulateBold = bold && (resolvedStyle & OfficeFontStyle.Bold) != OfficeFontStyle.Bold;
        bool simulateItalic = italic && (resolvedStyle & OfficeFontStyle.Italic) != OfficeFontStyle.Italic;
        double width = MeasureText(value, fontHeight, fontFamily, fontStyle);
        double x = ResolveAnchoredTextX(anchorX, width, alignment);
        if (font != null) {
            double outlineTop = ResolveRasterTextLineOutlineTop(font, fontHeight, top);
            IReadOnlyList<List<OfficePoint>> contours = TransformTextContours(
                GetResolvedTextContours(value, font, x, outlineTop, fontHeight),
                top + fontHeight,
                simulateItalic,
                transform);
            if (simulateBold) {
                var shifted = TransformTextContours(
                    GetResolvedTextContours(value, font, x + fontHeight / 24D, outlineTop, fontHeight),
                    top + fontHeight,
                    simulateItalic,
                    transform);
                FillTextContourUnion(contours, shifted, color);
            } else {
                FillTextContours(contours, color);
            }

            DrawAffineTextLineDecorations(x, width, top, fontHeight, color, transform, underline, strikethrough);
            return;
        }

        DrawAffineStrokeText(value, anchorX, top + (fontHeight / 2D), fontHeight, color, bold, italic, alignment, transform);
        DrawAffineTextLineDecorations(x, width, top, fontHeight, color, transform, underline, strikethrough);
    }

}
