using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free RGBA raster image buffer.
/// </summary>
public sealed class OfficeRasterImage {
    private readonly byte[] _pixels;

    internal byte[] PixelBuffer => _pixels;

    /// <summary>
    /// Creates a new RGBA image buffer.
    /// </summary>
    public OfficeRasterImage(int width, int height, OfficeColor? background = null) {
        if (width <= 0) {
            throw new ArgumentOutOfRangeException(nameof(width), "Width must be positive.");
        }

        if (height <= 0) {
            throw new ArgumentOutOfRangeException(nameof(height), "Height must be positive.");
        }

        Width = width;
        Height = height;
        _pixels = new byte[checked(width * height * 4)];
        if (background.HasValue) {
            Fill(background.Value);
        }
    }

    /// <summary>Image width in pixels.</summary>
    public int Width { get; }

    /// <summary>Image height in pixels.</summary>
    public int Height { get; }

    /// <summary>Returns a copy of the RGBA pixels.</summary>
    public byte[] GetPixels() {
        byte[] copy = new byte[_pixels.Length];
        Buffer.BlockCopy(_pixels, 0, copy, 0, _pixels.Length);
        return copy;
    }

    internal static OfficeRasterImage FromRgba32(int width, int height, byte[] pixels) {
        if (pixels == null) throw new ArgumentNullException(nameof(pixels));
        if (pixels.Length != checked(width * height * 4)) {
            throw new ArgumentException("RGBA buffer length does not match image dimensions.", nameof(pixels));
        }

        var image = new OfficeRasterImage(width, height);
        Buffer.BlockCopy(pixels, 0, image._pixels, 0, pixels.Length);
        return image;
    }

    /// <summary>Gets the color of a pixel.</summary>
    public OfficeColor GetPixel(int x, int y) {
        ValidateCoordinates(x, y);
        int offset = ((y * Width) + x) * 4;
        return OfficeColor.FromRgba(_pixels[offset], _pixels[offset + 1], _pixels[offset + 2], _pixels[offset + 3]);
    }

    /// <summary>Sets a pixel without alpha blending.</summary>
    public void SetPixel(int x, int y, OfficeColor color) {
        if ((uint)x >= (uint)Width || (uint)y >= (uint)Height) {
            return;
        }

        int offset = ((y * Width) + x) * 4;
        _pixels[offset] = color.R;
        _pixels[offset + 1] = color.G;
        _pixels[offset + 2] = color.B;
        _pixels[offset + 3] = color.A;
    }

    /// <summary>Alpha blends a pixel over the current pixel.</summary>
    public void BlendPixel(int x, int y, OfficeColor color) {
        if ((uint)x >= (uint)Width || (uint)y >= (uint)Height || color.A == 0) {
            return;
        }

        if (color.A == 255) {
            SetPixel(x, y, color);
            return;
        }

        int offset = ((y * Width) + x) * 4;
        double sourceAlpha = color.A / 255D;
        double targetAlpha = _pixels[offset + 3] / 255D;
        double outputAlpha = sourceAlpha + (targetAlpha * (1D - sourceAlpha));
        if (outputAlpha <= 0D) {
            _pixels[offset] = 0;
            _pixels[offset + 1] = 0;
            _pixels[offset + 2] = 0;
            _pixels[offset + 3] = 0;
            return;
        }

        _pixels[offset] = BlendChannel(color.R, _pixels[offset], sourceAlpha, targetAlpha, outputAlpha);
        _pixels[offset + 1] = BlendChannel(color.G, _pixels[offset + 1], sourceAlpha, targetAlpha, outputAlpha);
        _pixels[offset + 2] = BlendChannel(color.B, _pixels[offset + 2], sourceAlpha, targetAlpha, outputAlpha);
        _pixels[offset + 3] = (byte)Math.Round(outputAlpha * 255D);
    }

    /// <summary>Fills the whole image with a color.</summary>
    public void Fill(OfficeColor color) {
        for (int i = 0; i < _pixels.Length; i += 4) {
            _pixels[i] = color.R;
            _pixels[i + 1] = color.G;
            _pixels[i + 2] = color.B;
            _pixels[i + 3] = color.A;
        }
    }

    private void ValidateCoordinates(int x, int y) {
        if ((uint)x >= (uint)Width) {
            throw new ArgumentOutOfRangeException(nameof(x));
        }

        if ((uint)y >= (uint)Height) {
            throw new ArgumentOutOfRangeException(nameof(y));
        }
    }

    private static byte BlendChannel(byte source, byte target, double sourceAlpha, double targetAlpha, double outputAlpha) {
        double value = ((source * sourceAlpha) + (target * targetAlpha * (1D - sourceAlpha))) / outputAlpha;
        return (byte)Math.Round(value);
    }
}
