using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    /// <summary>Draws an affine image with managed opacity and blend-mode compositing.</summary>
    public void DrawAffineImage(OfficeRasterImage image, OfficeTransform transform, double opacity, OfficeBlendMode blendMode) {
        if (blendMode == OfficeBlendMode.Normal) {
            DrawAffineImage(image, transform, opacity);
            return;
        }

        if (image == null) throw new ArgumentNullException(nameof(image));
        if (double.IsNaN(opacity) || double.IsInfinity(opacity) || opacity < 0D || opacity > 1D) {
            throw new ArgumentOutOfRangeException(nameof(opacity), "Image opacity must be between zero and one.");
        }
        if (opacity <= 0D || !transform.TryInvert(out OfficeTransform inverse)) return;

        (double minX, double minY, double maxX, double maxY) = transform.TransformRectangleBounds(0D, 0D, image.Width, image.Height);
        int left = Clamp((int)Math.Floor(minX), 0, Width - 1);
        int top = Clamp((int)Math.Floor(minY), 0, Height - 1);
        int right = Clamp((int)Math.Ceiling(maxX), 0, Width - 1);
        int bottom = Clamp((int)Math.Ceiling(maxY), 0, Height - 1);
        for (int y = top; y <= bottom; y++) {
            for (int x = left; x <= right; x++) {
                if (!IsPixelInsideClip(x, y)) continue;
                OfficePoint sourcePoint = inverse.TransformPoint(new OfficePoint(x + 0.5D, y + 0.5D));
                if (sourcePoint.X < 0D || sourcePoint.X >= image.Width || sourcePoint.Y < 0D || sourcePoint.Y >= image.Height) continue;
                OfficeColor source = SampleBilinear(image, sourcePoint.X - 0.5D, sourcePoint.Y - 0.5D);
                if (opacity < 1D) source = OfficeColor.FromRgba(source.R, source.G, source.B, (byte)Math.Round(source.A * opacity));
                CompositeBlendPixel(x, y, source, blendMode);
            }
        }
    }

    private void CompositeBlendPixel(int x, int y, OfficeColor source, OfficeBlendMode blendMode) {
        if (source.A == 0) return;
        OfficeColor backdrop = _image != null ? _image.GetPixel(x, y) : _target!.GetPixel(x, y);
        double sourceAlpha = source.A / 255D;
        double backdropAlpha = backdrop.A / 255D;
        double outputAlpha = sourceAlpha + (backdropAlpha * (1D - sourceAlpha));
        if (outputAlpha <= 0D) return;

        double[] cs = { source.R / 255D, source.G / 255D, source.B / 255D };
        double[] cb = { backdrop.R / 255D, backdrop.G / 255D, backdrop.B / 255D };
        double[] blended = Blend(cb, cs, blendMode);
        byte red = CompositeChannel(cs[0], cb[0], blended[0], sourceAlpha, backdropAlpha, outputAlpha);
        byte green = CompositeChannel(cs[1], cb[1], blended[1], sourceAlpha, backdropAlpha, outputAlpha);
        byte blue = CompositeChannel(cs[2], cb[2], blended[2], sourceAlpha, backdropAlpha, outputAlpha);
        OfficeColor output = OfficeColor.FromRgba(red, green, blue, (byte)Math.Round(outputAlpha * 255D));
        if (_image != null) _image.SetPixel(x, y, output); else _target!.SetPixel(x, y, output);
    }

    private static byte CompositeChannel(double source, double backdrop, double blended, double sourceAlpha, double backdropAlpha, double outputAlpha) {
        double adjustedSource = ((1D - backdropAlpha) * source) + (backdropAlpha * blended);
        double premultiplied = (sourceAlpha * adjustedSource) + (backdropAlpha * (1D - sourceAlpha) * backdrop);
        return ToChannel(premultiplied / outputAlpha);
    }

    private static double[] Blend(double[] backdrop, double[] source, OfficeBlendMode mode) {
        switch (mode) {
            case OfficeBlendMode.Hue:
                return SetLuminosity(SetSaturation(source, Saturation(backdrop)), Luminosity(backdrop));
            case OfficeBlendMode.Saturation:
                return SetLuminosity(SetSaturation(backdrop, Saturation(source)), Luminosity(backdrop));
            case OfficeBlendMode.Color:
                return SetLuminosity(source, Luminosity(backdrop));
            case OfficeBlendMode.Luminosity:
                return SetLuminosity(backdrop, Luminosity(source));
            default:
                return new[] {
                    BlendComponent(backdrop[0], source[0], mode),
                    BlendComponent(backdrop[1], source[1], mode),
                    BlendComponent(backdrop[2], source[2], mode)
                };
        }
    }

    private static double BlendComponent(double backdrop, double source, OfficeBlendMode mode) {
        switch (mode) {
            case OfficeBlendMode.Multiply: return backdrop * source;
            case OfficeBlendMode.Screen: return backdrop + source - (backdrop * source);
            case OfficeBlendMode.Overlay: return HardLight(source, backdrop);
            case OfficeBlendMode.Darken: return Math.Min(backdrop, source);
            case OfficeBlendMode.Lighten: return Math.Max(backdrop, source);
            case OfficeBlendMode.ColorDodge: return source >= 1D ? 1D : Math.Min(1D, backdrop / (1D - source));
            case OfficeBlendMode.ColorBurn: return source <= 0D ? 0D : 1D - Math.Min(1D, (1D - backdrop) / source);
            case OfficeBlendMode.HardLight: return HardLight(backdrop, source);
            case OfficeBlendMode.SoftLight: return SoftLight(backdrop, source);
            case OfficeBlendMode.Difference: return Math.Abs(backdrop - source);
            case OfficeBlendMode.Exclusion: return backdrop + source - (2D * backdrop * source);
            default: return source;
        }
    }

    private static double HardLight(double backdrop, double source) =>
        source <= 0.5D ? 2D * backdrop * source : 1D - (2D * (1D - backdrop) * (1D - source));

    private static double SoftLight(double backdrop, double source) {
        if (source <= 0.5D) return backdrop - ((1D - (2D * source)) * backdrop * (1D - backdrop));
        double d = backdrop <= 0.25D
            ? (((16D * backdrop - 12D) * backdrop) + 4D) * backdrop
            : Math.Sqrt(backdrop);
        return backdrop + ((2D * source - 1D) * (d - backdrop));
    }

    private static double Luminosity(double[] color) => (0.3D * color[0]) + (0.59D * color[1]) + (0.11D * color[2]);

    private static double Saturation(double[] color) => Math.Max(color[0], Math.Max(color[1], color[2])) - Math.Min(color[0], Math.Min(color[1], color[2]));

    private static double[] SetLuminosity(double[] color, double luminosity) {
        double delta = luminosity - Luminosity(color);
        return ClipColor(new[] { color[0] + delta, color[1] + delta, color[2] + delta });
    }

    private static double[] SetSaturation(double[] color, double saturation) {
        var result = new double[3];
        int min = color[0] <= color[1] ? (color[0] <= color[2] ? 0 : 2) : (color[1] <= color[2] ? 1 : 2);
        int max = color[0] >= color[1] ? (color[0] >= color[2] ? 0 : 2) : (color[1] >= color[2] ? 1 : 2);
        int mid = 3 - min - max;
        if (color[max] > color[min]) {
            result[mid] = ((color[mid] - color[min]) * saturation) / (color[max] - color[min]);
            result[max] = saturation;
        }
        return result;
    }

    private static double[] ClipColor(double[] color) {
        double luminosity = Luminosity(color);
        double min = Math.Min(color[0], Math.Min(color[1], color[2]));
        double max = Math.Max(color[0], Math.Max(color[1], color[2]));
        if (min < 0D) {
            for (int i = 0; i < 3; i++) color[i] = luminosity + (((color[i] - luminosity) * luminosity) / (luminosity - min));
        }
        if (max > 1D) {
            for (int i = 0; i < 3; i++) color[i] = luminosity + (((color[i] - luminosity) * (1D - luminosity)) / (max - luminosity));
        }
        for (int i = 0; i < 3; i++) color[i] = Clamp01(color[i]);
        return color;
    }

    private static byte ToChannel(double value) => (byte)Math.Round(Clamp01(value) * 255D);

    private static double Clamp01(double value) => value < 0D ? 0D : value > 1D ? 1D : value;
}
