using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Supersampled RGBA render target used by dependency-free raster renderers.
/// </summary>
public sealed class OfficeRasterRenderTarget {
    private readonly byte[] _pixels;

    /// <summary>
    /// Creates a render target with an optional supersampling factor.
    /// </summary>
    public OfficeRasterRenderTarget(int width, int height, int supersampling = 1, OfficeColor? background = null) {
        if (width <= 0) {
            throw new ArgumentOutOfRangeException(nameof(width), "Width must be positive.");
        }

        if (height <= 0) {
            throw new ArgumentOutOfRangeException(nameof(height), "Height must be positive.");
        }

        if (supersampling <= 0) {
            throw new ArgumentOutOfRangeException(nameof(supersampling), "Supersampling must be positive.");
        }

        Width = width;
        Height = height;
        Supersampling = supersampling;
        RenderWidth = checked(width * supersampling);
        RenderHeight = checked(height * supersampling);
        _pixels = new byte[checked(RenderWidth * RenderHeight * 4)];
        if (background.HasValue) {
            Fill(background.Value);
        }
    }

    /// <summary>Resolved output width in pixels.</summary>
    public int Width { get; }

    /// <summary>Resolved output height in pixels.</summary>
    public int Height { get; }

    /// <summary>Supersampling factor applied to render coordinates.</summary>
    public int Supersampling { get; }

    /// <summary>Internal render-buffer width in pixels.</summary>
    public int RenderWidth { get; }

    /// <summary>Internal render-buffer height in pixels.</summary>
    public int RenderHeight { get; }

    /// <summary>Gets a render-buffer pixel.</summary>
    public OfficeColor GetPixel(int x, int y) {
        ValidateRenderCoordinates(x, y);
        int offset = ((y * RenderWidth) + x) * 4;
        return OfficeColor.FromRgba(_pixels[offset], _pixels[offset + 1], _pixels[offset + 2], _pixels[offset + 3]);
    }

    /// <summary>Sets a render-buffer pixel without alpha blending.</summary>
    public void SetPixel(int x, int y, OfficeColor color) {
        if ((uint)x >= (uint)RenderWidth || (uint)y >= (uint)RenderHeight) {
            return;
        }

        int offset = ((y * RenderWidth) + x) * 4;
        _pixels[offset] = color.R;
        _pixels[offset + 1] = color.G;
        _pixels[offset + 2] = color.B;
        _pixels[offset + 3] = color.A;
    }

    /// <summary>Alpha blends a render-buffer pixel over the current pixel.</summary>
    public void BlendPixel(int x, int y, OfficeColor color) {
        if ((uint)x >= (uint)RenderWidth || (uint)y >= (uint)RenderHeight || color.A == 0) {
            return;
        }

        int offset = ((y * RenderWidth) + x) * 4;
        int sourceAlpha = color.A;
        if (sourceAlpha == 255 || _pixels[offset + 3] == 0) {
            _pixels[offset] = color.R;
            _pixels[offset + 1] = color.G;
            _pixels[offset + 2] = color.B;
            _pixels[offset + 3] = color.A;
            return;
        }

        int targetAlpha = _pixels[offset + 3];
        int outputAlpha = sourceAlpha + ((targetAlpha * (255 - sourceAlpha)) / 255);
        if (outputAlpha == 0) {
            return;
        }

        _pixels[offset] = (byte)(((color.R * sourceAlpha) + (_pixels[offset] * targetAlpha * (255 - sourceAlpha) / 255)) / outputAlpha);
        _pixels[offset + 1] = (byte)(((color.G * sourceAlpha) + (_pixels[offset + 1] * targetAlpha * (255 - sourceAlpha) / 255)) / outputAlpha);
        _pixels[offset + 2] = (byte)(((color.B * sourceAlpha) + (_pixels[offset + 2] * targetAlpha * (255 - sourceAlpha) / 255)) / outputAlpha);
        _pixels[offset + 3] = (byte)outputAlpha;
    }

    /// <summary>Fills the internal render buffer.</summary>
    public void Fill(OfficeColor color) {
        for (int i = 0; i < _pixels.Length; i += 4) {
            _pixels[i] = color.R;
            _pixels[i + 1] = color.G;
            _pixels[i + 2] = color.B;
            _pixels[i + 3] = color.A;
        }
    }

    /// <summary>Resolves the render target into output-size RGBA bytes.</summary>
    public byte[] ResolveRgba() {
        if (Supersampling == 1) {
            byte[] clone = new byte[_pixels.Length];
            Buffer.BlockCopy(_pixels, 0, clone, 0, _pixels.Length);
            return clone;
        }

        byte[] output = new byte[Width * Height * 4];
        int samples = Supersampling * Supersampling;
        for (int y = 0; y < Height; y++) {
            for (int x = 0; x < Width; x++) {
                int alpha = 0;
                long red = 0;
                long green = 0;
                long blue = 0;
                for (int sy = 0; sy < Supersampling; sy++) {
                    for (int sx = 0; sx < Supersampling; sx++) {
                        int source = ((((y * Supersampling) + sy) * RenderWidth) + ((x * Supersampling) + sx)) * 4;
                        int sampleAlpha = _pixels[source + 3];
                        red += _pixels[source] * sampleAlpha;
                        green += _pixels[source + 1] * sampleAlpha;
                        blue += _pixels[source + 2] * sampleAlpha;
                        alpha += sampleAlpha;
                    }
                }

                int target = ((y * Width) + x) * 4;
                if (alpha > 0) {
                    output[target] = (byte)((red + (alpha / 2L)) / alpha);
                    output[target + 1] = (byte)((green + (alpha / 2L)) / alpha);
                    output[target + 2] = (byte)((blue + (alpha / 2L)) / alpha);
                }

                output[target + 3] = (byte)(alpha / samples);
            }
        }

        return output;
    }

    private void ValidateRenderCoordinates(int x, int y) {
        if ((uint)x >= (uint)RenderWidth) {
            throw new ArgumentOutOfRangeException(nameof(x));
        }

        if ((uint)y >= (uint)RenderHeight) {
            throw new ArgumentOutOfRangeException(nameof(y));
        }
    }
}
