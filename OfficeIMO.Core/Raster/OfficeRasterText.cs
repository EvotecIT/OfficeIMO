using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Text layout and raster effects shared by managed image-editing consumers.</summary>
public static class OfficeRasterText {
    /// <summary>Estimates peak managed text scratch storage beyond the target RGBA pixel buffer.</summary>
    /// <remarks>Includes the full-image glyph mask, an optional outline surface, row scratch and fixed working overhead.
    /// The estimate is conservative even for empty text or a clipped rectangle. It excludes caller-retained source and result images,
    /// which frame operations account for separately. The options are captured and validated before calculating the estimate.</remarks>
    /// <param name="width">Target image width in pixels.</param>
    /// <param name="height">Target image height in pixels.</param>
    /// <param name="options">Shared text and effect settings used by drawing.</param>
    /// <returns>Additional bytes to retain in a composed operation's working-set budget.</returns>
    public static long EstimateAdditionalWorkingBytes(int width, int height, OfficeRasterTextOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        OfficeRasterTextOptions settings = options.Clone();
        settings.Validate();
        return CalculateAdditionalWorkingBytes(width, height, settings);
    }

    /// <summary>Measures authored lines or wraps them to a requested width using managed font metrics.</summary>
    public static OfficeTextBlockLayout Measure(string? text, double fontSize = 16, string? fontFamily = null,
        double? wrapWidth = null, CancellationToken cancellationToken = default) =>
        Measure(text, new OfficeRasterTextOptions { FontSize = fontSize, FontFamily = fontFamily, Wrap = wrapWidth.HasValue }, wrapWidth, cancellationToken);

    /// <summary>Measures text using the same captured font and layout settings as raster drawing.</summary>
    /// <param name="text">Text to measure.</param>
    /// <param name="options">Shared font and layout settings.</param>
    /// <param name="wrapWidth">Finite positive wrap width, required when Wrap is enabled.</param>
    /// <param name="cancellationToken">Observes cancellation during measurement and layout.</param>
    public static OfficeTextBlockLayout Measure(string? text, OfficeRasterTextOptions options,
        double? wrapWidth = null, CancellationToken cancellationToken = default) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        cancellationToken.ThrowIfCancellationRequested();
        OfficeRasterTextOptions settings = options.Clone();
        settings.Validate();
        if (wrapWidth.HasValue && (!Finite(wrapWidth.Value) || wrapWidth.Value <= 0)) throw new ArgumentOutOfRangeException(nameof(wrapWidth));
        if (settings.Wrap && !wrapWidth.HasValue) throw new ArgumentException("Wrapped text measurement requires a width.", nameof(wrapWidth));
        var canvas = CreateCanvas(new OfficeRasterImage(1, 1), settings, cancellationToken);
        return MeasureLayout(text, settings, canvas, wrapWidth, cancellationToken);
    }

    /// <summary>Draws measured text, optional shadow, and a rounded raster outline over existing pixels.</summary>
    /// <remarks>The input image is mutated. Effects use the same glyph coverage and clipping rectangle as the text.</remarks>
    public static void Draw(OfficeRasterImage target, string? text, double x, double y, double width, double height, OfficeColor color,
        double fontSize = 16, string? fontFamily = null, CancellationToken cancellationToken = default) =>
        Draw(target, text, x, y, width, height, color, new OfficeRasterTextOptions { FontSize = fontSize, FontFamily = fontFamily }, cancellationToken);

    /// <summary>Draws text and optional effects with the same shared font and layout settings used by measurement.</summary>
    /// <remarks>The target is mutated. The rectangle places the measured block and supplies its wrap width when wrapping is enabled.
    /// Rectangle height does not remove lines; glyphs and effects are cut at its edges only when Clip is enabled.
    /// Cancellation can leave painted pixels.</remarks>
    public static void Draw(OfficeRasterImage target, string? text, double x, double y, double width, double height, OfficeColor color,
        OfficeRasterTextOptions options, CancellationToken cancellationToken = default) {
        if (target == null) {
            throw new ArgumentNullException(nameof(target));
        }
        if (options == null) throw new ArgumentNullException(nameof(options));
        cancellationToken.ThrowIfCancellationRequested();
        OfficeRasterTextOptions settings = options.Clone();
        settings.Validate();
        if (width <= 0 || height <= 0 || !Finite(x) || !Finite(y) || !Finite(width) || !Finite(height)) { throw new ArgumentOutOfRangeException(nameof(width)); }
        long bufferBytes = checked((long)target.Width * target.Height * 4);
        if (bufferBytes + CalculateAdditionalWorkingBytes(target.Width, target.Height, settings) > OfficeRasterGuards.MaximumDecodedBytes) {
            throw new ArgumentException("Text effects exceed the managed working-set limit.", nameof(target));
        }
        cancellationToken.ThrowIfCancellationRequested();
        var mask = new OfficeRasterImage(target.Width, target.Height);
        var canvas = CreateCanvas(mask, settings, cancellationToken);
        OfficeTextBlockLayout layout = MeasureLayout(text, settings, canvas, settings.Wrap ? width : null, cancellationToken);
        var plan = OfficeTextBlockRenderPlan.CreateFromRectangle(layout, x, y, width, height,
            settings.HorizontalAlignment, settings.VerticalAlignment);
        using (settings.Clip ? canvas.PushClipRectangle(x, y, width, height) : null) {
            OfficeTextBlockRenderer.DrawRasterTextBox(canvas, plan, OfficeColor.White,
                bold: (settings.FontStyle & OfficeFontStyle.Bold) != 0, italic: (settings.FontStyle & OfficeFontStyle.Italic) != 0,
                underline: (settings.FontStyle & OfficeFontStyle.Underline) != 0, strikethrough: (settings.FontStyle & OfficeFontStyle.Strikethrough) != 0,
                horizontalAlignment: settings.HorizontalAlignment, verticalAlignment: settings.VerticalAlignment, fontFamily: settings.FontFamily);
        }
        if (settings.ShadowColor.HasValue) CompositeMask(target, mask, settings.ShadowColor.Value, settings.ShadowOffsetX, settings.ShadowOffsetY, settings.Clip, x, y, width, height, cancellationToken);
        if (settings.OutlineColor.HasValue && settings.OutlineWidth > 0) {
            OfficeRasterImage outline = ExpandCoverage(mask, settings.OutlineWidth, cancellationToken);
            CompositeMask(target, outline, settings.OutlineColor.Value, 0, 0, settings.Clip, x, y, width, height, cancellationToken);
        }
        CompositeMask(target, mask, color, 0, 0, settings.Clip, x, y, width, height, cancellationToken);
    }

    /// <summary>Measures complete authored or wrapped lines with the captured canvas and fractional raster line height.</summary>
    private static OfficeTextBlockLayout MeasureLayout(string? text, OfficeRasterTextOptions settings, OfficeRasterCanvas canvas,
        double? wrapWidth, CancellationToken cancellationToken) {
        Func<string?, double, double> measure = (value, size) => {
            cancellationToken.ThrowIfCancellationRequested();
            return canvas.MeasureText(value, size, settings.FontFamily, settings.FontStyle);
        };
        IReadOnlyList<OfficeTextLine> lines = settings.Wrap
            ? OfficeTextLayoutEngine.WrapLines(text, settings.FontSize, wrapWidth!.Value, measure)
            : OfficeTextLayoutEngine.MeasureUnwrappedLines(text, settings.FontSize, measure);
        cancellationToken.ThrowIfCancellationRequested();
        double lineHeight = settings.FontSize * settings.LineHeight;
        return new OfficeTextBlockLayout(lines, settings.FontSize, lineHeight, OfficeTextLayoutEngine.MeasureMaxLineWidth(lines), lines.Count * lineHeight);
    }

    private static OfficeRasterCanvas CreateCanvas(OfficeRasterImage image, OfficeRasterTextOptions settings, CancellationToken token) =>
        new OfficeRasterCanvas(image, settings.Font, settings.Fonts, settings.TextShapingProvider, settings.TextShapingLanguage,
            settings.DiagnosticSink, settings.DiagnosticSource, token);

    private static long CalculateAdditionalWorkingBytes(int width, int height, OfficeRasterTextOptions settings) {
        long pixels = OfficeRasterGuards.EnsureOutputPixels(width, height, "Text target dimensions exceed the managed image limit.");
        int surfaces = settings.OutlineColor.HasValue && settings.OutlineWidth > 0 ? 2 : 1;
        return checked(pixels * 4L * surfaces + (long)width * 4L + 65536L);
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
