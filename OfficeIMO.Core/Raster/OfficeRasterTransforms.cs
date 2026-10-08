using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Reusable geometric edits of managed raster pixels. Every operation returns independent pixels.</summary>
public static class OfficeRasterTransforms {
    /// <summary>Crops an exact in-bounds rectangular pixel region.</summary>
    public static OfficeRasterImage Crop(OfficeRasterImage source, int x, int y, int width, int height, CancellationToken cancellationToken = default) {
        if (source == null) {
            throw new ArgumentNullException(nameof(source));
        }
        if (x < 0 || y < 0 || width <= 0 || height <= 0 || (long)x + width > source.Width || (long)y + height > source.Height) {
            throw new ArgumentOutOfRangeException(nameof(width), "Crop bounds must be inside the source image.");
        }
        OfficeRasterGuards.EnsureOutputPixels(width, height, "Crop dimensions exceed the managed pixel limit.");
        EnsureWorkingSet(source, (long)width * height * 4, cancellationToken);
        var result = new OfficeRasterImage(width, height);
        int rowBytes = checked(width * 4);
        for (int row = 0; row < height; row++) {
            cancellationToken.ThrowIfCancellationRequested();
            Buffer.BlockCopy(source.PixelBuffer, checked(((y + row) * source.Width + x) * 4), result.PixelBuffer, row * rowBytes, rowBytes);
        }
        return result;
    }

    /// <summary>Mirrors an image horizontally, vertically, or along both axes.</summary>
    public static OfficeRasterImage Flip(OfficeRasterImage source, bool horizontal, bool vertical, CancellationToken cancellationToken = default) {
        if (source == null) {
            throw new ArgumentNullException(nameof(source));
        }
        EnsureWorkingSet(source, (long)source.Width * source.Height * 4, cancellationToken);
        var result = new OfficeRasterImage(source.Width, source.Height);
        for (int y = 0; y < source.Height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < source.Width; x++) {
                if ((x & 1023) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
                int fromX = horizontal ? source.Width - 1 - x : x;
                int fromY = vertical ? source.Height - 1 - y : y;
                Buffer.BlockCopy(source.PixelBuffer, (fromY * source.Width + fromX) * 4, result.PixelBuffer, (y * source.Width + x) * 4, 4);
            }
        }
        return result;
    }

    /// <summary>Rotates clockwise and expands the canvas to retain the entire source.</summary>
    public static OfficeRasterImage Rotate(OfficeRasterImage source, double degrees, OfficeColor? background = null, CancellationToken cancellationToken = default) {
        double normalized = ((Finite(degrees, nameof(degrees)) % 360) + 360) % 360;
        if (normalized == 0) { return AutoOrient(source, 1, cancellationToken); }
        if (normalized == 90) { return AutoOrient(source, 6, cancellationToken); }
        if (normalized == 180) { return AutoOrient(source, 3, cancellationToken); }
        if (normalized == 270) { return AutoOrient(source, 8, cancellationToken); }
        return TransformExpanded(source, OfficeTransform.RotateDegrees(normalized), background, cancellationToken);
    }

    /// <summary>Skews each axis by its angle in degrees and expands the canvas to retain all pixels.</summary>
    public static OfficeRasterImage Skew(OfficeRasterImage source, double degreesX, double degreesY, OfficeColor? background = null, CancellationToken cancellationToken = default) {
        return TransformExpanded(source, GetSkewTransform(degreesX, degreesY), background, cancellationToken);
    }

    /// <summary>Plans the exact canvas dimensions produced by a clockwise rotation without allocating pixels.</summary>
    public static (int Width, int Height) GetRotatedSize(OfficeRasterImage source, double degrees) {
        if (source == null) { throw new ArgumentNullException(nameof(source)); }
        double normalized = ((Finite(degrees, nameof(degrees)) % 360) + 360) % 360;
        if (normalized == 0 || normalized == 180) { return (source.Width, source.Height); }
        if (normalized == 90 || normalized == 270) { return (source.Height, source.Width); }
        var bounds = GetExpandedBounds(source, OfficeTransform.RotateDegrees(normalized));
        return (bounds.Width, bounds.Height);
    }

    /// <summary>Plans the exact expanded canvas dimensions produced by a skew without allocating pixels.</summary>
    public static (int Width, int Height) GetSkewedSize(OfficeRasterImage source, double degreesX, double degreesY) {
        var bounds = GetExpandedBounds(source, GetSkewTransform(degreesX, degreesY));
        return (bounds.Width, bounds.Height);
    }

    private static OfficeTransform GetSkewTransform(double degreesX, double degreesY) {
        double x = Math.Tan((Finite(degreesX, nameof(degreesX)) % 180D) * Math.PI / 180D);
        double y = Math.Tan((Finite(degreesY, nameof(degreesY)) % 180D) * Math.PI / 180D);
        if (Math.Abs(x) > 1E6 || Math.Abs(y) > 1E6) {
            throw new ArgumentOutOfRangeException(nameof(degreesX), "The skew angle approaches a singular transform.");
        }
        var transform = new OfficeTransform(1, y, x, 1, 0, 0);
        if (!transform.TryInvert(out _)) {
            throw new ArgumentOutOfRangeException(nameof(degreesX), "The combined skew angles produce a singular transform.");
        }
        return transform;
    }

    /// <summary>Applies an EXIF orientation value, including mirrored orientations.</summary>
    public static OfficeRasterImage AutoOrient(OfficeRasterImage source, int orientation, CancellationToken cancellationToken = default) {
        if (source == null) { throw new ArgumentNullException(nameof(source)); }
        if (orientation < 1 || orientation > 8) { throw new ArgumentOutOfRangeException(nameof(orientation), "EXIF orientation must be between one and eight."); }
        EnsureWorkingSet(source, (long)source.Width * source.Height * 4, cancellationToken);
        int width = orientation < 5 ? source.Width : source.Height;
        int height = orientation < 5 ? source.Height : source.Width;
        var result = new OfficeRasterImage(width, height);
        for (int y = 0; y < height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < width; x++) {
                if ((x & 1023) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
                int sx, sy;
                switch (orientation) {
                    case 1: sx = x; sy = y; break;
                    case 2: sx = source.Width - 1 - x; sy = y; break;
                    case 3: sx = source.Width - 1 - x; sy = source.Height - 1 - y; break;
                    case 4: sx = x; sy = source.Height - 1 - y; break;
                    case 5: sx = y; sy = x; break;
                    case 6: sx = y; sy = source.Height - 1 - x; break;
                    case 7: sx = source.Width - 1 - y; sy = source.Height - 1 - x; break;
                    default: sx = source.Width - 1 - y; sy = x; break;
                }
                Buffer.BlockCopy(source.PixelBuffer, (sy * source.Width + sx) * 4, result.PixelBuffer, (y * width + x) * 4, 4);
            }
        }
        return result;
    }

    /// <summary>Keeps pixels inside an antialiased ellipse and clears pixels outside it.</summary>
    public static OfficeRasterImage MaskEllipse(OfficeRasterImage source, double x, double y, double width, double height, CancellationToken cancellationToken = default) {
        if (source == null) {
            throw new ArgumentNullException(nameof(source));
        }
        if (Finite(width, nameof(width)) <= 0 || Finite(height, nameof(height)) <= 0) {
            throw new ArgumentOutOfRangeException(nameof(width));
        }
        Finite(x, nameof(x));
        Finite(y, nameof(y));
        EnsureWorkingSet(source, (long)source.Width * source.Height * 8, cancellationToken);
        var mask = new OfficeRasterImage(source.Width, source.Height);
        new OfficeRasterCanvas(mask, font: null, fonts: null, cancellationToken: cancellationToken).FillEllipse(Finite(x, nameof(x)), Finite(y, nameof(y)), width, height, OfficeColor.White);
        return ApplyMask(source, mask, cancellationToken);
    }

    /// <summary>Keeps pixels inside an antialiased polygon.</summary>
    public static OfficeRasterImage MaskPolygon(OfficeRasterImage source, IReadOnlyList<OfficePoint> points, CancellationToken cancellationToken = default) {
        if (source == null) {
            throw new ArgumentNullException(nameof(source));
        }
        if (points == null) {
            throw new ArgumentNullException(nameof(points));
        }
        if (points.Count < 3) {
            throw new ArgumentException("At least three polygon vertices are required.", nameof(points));
        }
        EnsureWorkingSet(source, (long)source.Width * source.Height * 8, cancellationToken);
        var mask = new OfficeRasterImage(source.Width, source.Height);
        new OfficeRasterCanvas(mask, font: null, fonts: null, cancellationToken: cancellationToken).FillPolygon(points, OfficeColor.White);
        return ApplyMask(source, mask, cancellationToken);
    }

    /// <summary>Clears pixels outside the rounded corners of the image canvas.</summary>
    public static OfficeRasterImage MaskRoundedRectangle(OfficeRasterImage source, double radius, CancellationToken cancellationToken = default) {
        if (source == null) {
            throw new ArgumentNullException(nameof(source));
        }
        if (Finite(radius, nameof(radius)) < 0) {
            throw new ArgumentOutOfRangeException(nameof(radius));
        }
        radius = Math.Min(radius, Math.Min(source.Width, source.Height) / 2D);
        if (radius == 0) { return AutoOrient(source, 1, cancellationToken); }
        EnsureWorkingSet(source, (long)source.Width * source.Height * 8, cancellationToken);
        var mask = new OfficeRasterImage(source.Width, source.Height);
        var canvas = new OfficeRasterCanvas(mask, font: null, fonts: null, cancellationToken: cancellationToken);
        canvas.FillRectangle(radius, 0, source.Width - radius * 2, source.Height, OfficeColor.White);
        canvas.FillRectangle(0, radius, source.Width, source.Height - radius * 2, OfficeColor.White);
        double diameter = radius * 2;
        canvas.FillEllipse(0, 0, diameter, diameter, OfficeColor.White);
        canvas.FillEllipse(source.Width - diameter, 0, diameter, diameter, OfficeColor.White);
        canvas.FillEllipse(0, source.Height - diameter, diameter, diameter, OfficeColor.White);
        canvas.FillEllipse(source.Width - diameter, source.Height - diameter, diameter, diameter, OfficeColor.White);
        return ApplyMask(source, mask, cancellationToken);
    }

    private static OfficeRasterImage ApplyMask(OfficeRasterImage source, OfficeRasterImage mask, CancellationToken cancellationToken) {
        var result = source.Clone();
        for (int y = 0; y < source.Height; y++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int x = 0; x < source.Width; x++) {
                if ((x & 1023) == 0) { cancellationToken.ThrowIfCancellationRequested(); }
                int index = (y * source.Width + x) * 4;
                result.PixelBuffer[index + 3] = (byte)((source.PixelBuffer[index + 3] * mask.PixelBuffer[index + 3] + 127) / 255);
            }
        }
        return result;
    }

    private static OfficeRasterImage TransformExpanded(OfficeRasterImage source, OfficeTransform transform, OfficeColor? background, CancellationToken cancellationToken) {
        var bounds = GetExpandedBounds(source, transform);
        return OfficeRasterResampler.Transform(source, transform.Then(OfficeTransform.Translate(-bounds.Left, -bounds.Top)), bounds.Width, bounds.Height, background ?? OfficeColor.Transparent, cancellationToken);
    }

    private static (double Left, double Top, int Width, int Height) GetExpandedBounds(OfficeRasterImage source, OfficeTransform transform) {
        if (source == null) {
            throw new ArgumentNullException(nameof(source));
        }
        OfficePoint[] corners = {
            transform.TransformPoint(new OfficePoint(0, 0)),
            transform.TransformPoint(new OfficePoint(source.Width, 0)),
            transform.TransformPoint(new OfficePoint(0, source.Height)),
            transform.TransformPoint(new OfficePoint(source.Width, source.Height))
        };
        double left = corners[0].X, top = corners[0].Y, right = left, bottom = top;
        for (int i = 1; i < corners.Length; i++) {
            left = Math.Min(left, corners[i].X);
            top = Math.Min(top, corners[i].Y);
            right = Math.Max(right, corners[i].X);
            bottom = Math.Max(bottom, corners[i].Y);
        }
        double width = Math.Ceiling(right - left - 1E-10), height = Math.Ceiling(bottom - top - 1E-10);
        if (width < 1 || height < 1 || width > int.MaxValue || height > int.MaxValue) {
            throw new ArgumentException("The transformed dimensions are not representable.", nameof(transform));
        }
        OfficeRasterGuards.EnsureOutputPixels((int)width, (int)height, "Transformed dimensions exceed the managed image limit.");
        return (left, top, (int)width, (int)height);
    }

    private static double Finite(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(name);
        }
        return value;
    }

    private static void EnsureWorkingSet(OfficeRasterImage source, long additionalBytes, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if ((long)source.Width * source.Height * 4 + additionalBytes + 65536 > OfficeRasterGuards.MaximumDecodedBytes) {
            throw new ArgumentException("The raster transformation exceeds the managed working-set limit.", nameof(source));
        }
    }
}
