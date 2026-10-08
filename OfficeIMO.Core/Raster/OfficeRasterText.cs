using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Text layout and raster effects shared by managed image-editing consumers.</summary>
public static class OfficeRasterText {
    /// <summary>Measures authored lines or wraps them to a requested width using managed font metrics.</summary>
    public static OfficeTextBlockLayout Measure(string? text, double fontSize = 16, string? fontFamily = null, double? wrapWidth = null) {
        if (fontSize <= 0 || double.IsNaN(fontSize) || double.IsInfinity(fontSize)) {
            throw new ArgumentOutOfRangeException(nameof(fontSize));
        }
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(1, 1));
        Func<string?, double, double> measure = (value, size) => canvas.MeasureText(value, size, fontFamily);
        IReadOnlyList<OfficeTextLine> lines = wrapWidth.HasValue
            ? OfficeTextLayoutEngine.WrapLines(text, fontSize, wrapWidth.Value, measure)
            : OfficeTextLayoutEngine.MeasureUnwrappedLines(text, fontSize, measure);
        return new OfficeTextBlockLayout(lines, fontSize, fontSize * 1.2D, OfficeTextLayoutEngine.MeasureMaxLineWidth(lines), lines.Count * fontSize * 1.2D);
    }

    /// <summary>Draws measured text, optional shadow, and a rounded raster outline over existing pixels.</summary>
    /// <remarks>The input image is mutated. Effects use the same glyph coverage and clipping rectangle as the text.</remarks>
    public static void Draw(OfficeRasterImage target, string? text, double x, double y, double width, double height, OfficeColor color,
        double fontSize = 16, string? fontFamily = null, OfficeTextAlignment horizontalAlignment = OfficeTextAlignment.Left,
        OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top, bool wrap = false, bool clip = false,
        OfficeColor? shadowColor = null, double shadowOffsetX = 0, double shadowOffsetY = 0, OfficeColor? outlineColor = null,
        double outlineWidth = 0, CancellationToken cancellationToken = default) {
        if (target == null) {
            throw new ArgumentNullException(nameof(target));
        }
        if (double.IsNaN(outlineWidth) || double.IsInfinity(outlineWidth) || outlineWidth < 0 || outlineWidth > 64) {
            throw new ArgumentOutOfRangeException(nameof(outlineWidth), "Raster text outlines must be between zero and sixty-four pixels.");
        }
        if (width <= 0 || height <= 0 || !Finite(x) || !Finite(y) || !Finite(width) || !Finite(height)) { throw new ArgumentOutOfRangeException(nameof(width)); }
        if (!Finite(fontSize) || fontSize <= 0) { throw new ArgumentOutOfRangeException(nameof(fontSize)); }
        if (!Finite(shadowOffsetX) || !Finite(shadowOffsetY)) { throw new ArgumentOutOfRangeException(nameof(shadowOffsetX)); }
        long bufferBytes = checked((long)target.Width * target.Height * 4);
        if (bufferBytes * (outlineColor.HasValue && outlineWidth > 0 ? 3 : 2) + (long)target.Width * 4 + 65536 > OfficeRasterGuards.MaximumDecodedBytes) {
            throw new ArgumentException("Text effects exceed the managed working-set limit.", nameof(target));
        }
        cancellationToken.ThrowIfCancellationRequested();
        var mask = new OfficeRasterImage(target.Width, target.Height);
        var canvas = new OfficeRasterCanvas(mask, font: null, fonts: null, cancellationToken: cancellationToken);
        var plan = OfficeTextBlockRenderPlan.CreateTextBlockFromRectangle(text, fontSize, x, y, width, height,
            (value, size) => canvas.MeasureText(value, size, fontFamily), horizontalAlignment, verticalAlignment,
            1.2D, fontSize, wrap, false, false, OfficeTextOverflowBehavior.Clip);
        using (clip ? canvas.PushClipRectangle(x, y, width, height) : null) {
            OfficeTextBlockRenderer.DrawRasterTextBox(canvas, plan, OfficeColor.White, horizontalAlignment: horizontalAlignment, verticalAlignment: verticalAlignment, fontFamily: fontFamily);
        }
        if (shadowColor.HasValue) CompositeMask(target, mask, shadowColor.Value, shadowOffsetX, shadowOffsetY, clip, x, y, width, height, cancellationToken);
        if (outlineColor.HasValue && outlineWidth > 0) {
            OfficeRasterImage outline = ExpandCoverage(mask, outlineWidth, cancellationToken);
            CompositeMask(target, outline, outlineColor.Value, 0, 0, clip, x, y, width, height, cancellationToken);
        }
        CompositeMask(target, mask, color, 0, 0, clip, x, y, width, height, cancellationToken);
    }

    private static OfficeRasterImage ExpandCoverage(OfficeRasterImage mask, double radius, CancellationToken cancellationToken) {
        var result = new OfficeRasterImage(mask.Width, mask.Height);
        int range = (int)Math.Ceiling(radius);
        double maximumDistance = (radius + 0.5D) * (radius + 0.5D);
        int[] deque = new int[mask.Width];
        for (int y = 0; y < mask.Height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int dy = -range; dy <= range; dy++) {
                int row = y + dy;
                if ((uint)row >= (uint)mask.Height || dy * dy > maximumDistance) { continue; }
                int reach = Math.Min(range, (int)Math.Floor(Math.Sqrt(maximumDistance - dy * dy)));
                int head = 0, tail = 0, next = 0;
                for (int x = 0; x < mask.Width; x++) {
                    if ((x & 1023) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
                    int end = Math.Min(mask.Width - 1, x + reach);
                    while (next <= end) {
                        byte alpha = mask.PixelBuffer[(row * mask.Width + next) * 4 + 3];
                        while (tail > head && mask.PixelBuffer[(row * mask.Width + deque[tail - 1]) * 4 + 3] <= alpha) { tail--; }
                        deque[tail++] = next++;
                    }
                    while (head < tail && deque[head] < x - reach) { head++; }
                    byte coverage = mask.PixelBuffer[(row * mask.Width + deque[head]) * 4 + 3];
                    int offset = (y * mask.Width + x) * 4 + 3;
                    if (coverage > result.PixelBuffer[offset]) { result.PixelBuffer[offset] = coverage; }
                }
            }
        }
        return result;
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    private static void CompositeMask(OfficeRasterImage target, OfficeRasterImage mask, OfficeColor color, double offsetX, double offsetY,
        bool clip, double clipX, double clipY, double clipWidth, double clipHeight, CancellationToken cancellationToken) {
        if (double.IsNaN(offsetX) || double.IsInfinity(offsetX) || double.IsNaN(offsetY) || double.IsInfinity(offsetY)) {
            throw new ArgumentOutOfRangeException(nameof(offsetX));
        }
        if (offsetX <= -target.Width || offsetX >= target.Width || offsetY <= -target.Height || offsetY >= target.Height) return;
        int shiftX = (int)Math.Round(offsetX), shiftY = (int)Math.Round(offsetY);
        for (int y = 0; y < mask.Height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            int destinationY = y + shiftY;
            if ((uint)destinationY >= (uint)target.Height) {
                continue;
            }
            for (int x = 0; x < mask.Width; x++) {
                if ((x & 1023) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
                int destinationX = x + shiftX;
                if ((uint)destinationX >= (uint)target.Width) {
                    continue;
                }
                if (clip && (destinationX < clipX || destinationX >= clipX + clipWidth || destinationY < clipY || destinationY >= clipY + clipHeight)) {
                    continue;
                }
                byte coverage = mask.PixelBuffer[(y * mask.Width + x) * 4 + 3];
                if (coverage != 0) target.BlendPixel(destinationX, destinationY, OfficeColor.FromRgba(color.R, color.G, color.B, (byte)((coverage * color.A + 127) / 255)));
            }
        }
    }
}
