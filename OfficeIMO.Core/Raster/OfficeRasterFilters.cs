using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Luminance coefficients used by general raster color conversion.</summary>
public enum OfficeRasterGrayscaleMode {
    /// <summary>ITU-R BT.709 encoded-channel weights: 0.2126, 0.7152 and 0.0722.</summary>
    Bt709,
    /// <summary>ITU-R BT.601 encoded-channel weights: 0.299, 0.587 and 0.114.</summary>
    Bt601
}

/// <summary>An immutable normalized-channel affine color transform with four output rows and five input columns.</summary>
/// <remarks>Each row multiplies normalized R, G, B, A and a constant one. Results are clamped to zero through one.</remarks>
public sealed class OfficeRasterColorMatrix {
    private readonly double[] _values;

    /// <summary>Creates a matrix by copying exactly twenty finite row-major coefficients.</summary>
    public OfficeRasterColorMatrix(params double[] coefficients) {
        if (coefficients == null) throw new ArgumentNullException(nameof(coefficients));
        if (coefficients.Length != 20) throw new ArgumentException("A color matrix requires twenty coefficients.", nameof(coefficients));
        foreach (double coefficient in coefficients) {
            if (double.IsNaN(coefficient) || double.IsInfinity(coefficient) || Math.Abs(coefficient) > 1_000_000D) {
                throw new ArgumentOutOfRangeException(nameof(coefficients), "Color coefficients must be finite and bounded.");
            }
        }
        _values = (double[])coefficients.Clone();
    }

    /// <summary>Gets the identity color transform.</summary>
    public static OfficeRasterColorMatrix Identity { get; } = new OfficeRasterColorMatrix(
        1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 1, 0);

    /// <summary>Gets one coefficient by output-channel row and input-channel/bias column.</summary>
    public double GetCoefficient(int row, int column) {
        if ((uint)row >= 4U) throw new ArgumentOutOfRangeException(nameof(row));
        if ((uint)column >= 5U) throw new ArgumentOutOfRangeException(nameof(column));
        return _values[row * 5 + column];
    }

    internal double Apply(int row, double r, double g, double b, double a) {
        int offset = row * 5;
        return _values[offset] * r + _values[offset + 1] * g + _values[offset + 2] * b + _values[offset + 3] * a + _values[offset + 4];
    }
}

/// <summary>Bounded, dependency-free operations that return a separately owned raster without modifying the source.</summary>
/// <remarks>Color operations preserve straight alpha unless explicitly changing opacity or applying a color matrix.
/// Neighborhood operations use premultiplied alpha so invisible RGB values do not bleed into visible pixels.</remarks>
public static partial class OfficeRasterFilters {
    private const long MaximumOperations = 1_000_000_000L;

    private static void ValidateSource(OfficeRasterImage source, long additionalBytes = 0L, long operationsPerPixel = 1L) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        long pixels = OfficeRasterGuards.EnsureOutputPixels(source.Width, source.Height, "Raster filtering dimensions exceed the managed image limit.");
        if (additionalBytes < 0L || additionalBytes > OfficeRasterGuards.MaximumDecodedBytes - pixels * 8L - 64L * 1024L) {
            throw new ArgumentException("Raster filtering working set exceeds the managed image limit.", nameof(source));
        }
        if (operationsPerPixel <= 0L || pixels > MaximumOperations / operationsPerPixel) {
            throw new ArgumentException("Raster filtering exceeds the bounded operation count.", nameof(source));
        }
    }

    private static void ValidateAmount(double value, string name, double minimum = 0D, double maximum = 100D) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < minimum || value > maximum) {
            throw new ArgumentOutOfRangeException(name, value, "The amount is outside the supported finite range.");
        }
    }

    private static OfficeRasterImage Map(OfficeRasterImage source, Func<OfficeColor, int, int, OfficeColor> operation, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        ValidateSource(source);
        var result = new OfficeRasterImage(source.Width, source.Height);
        byte[] input = source.PixelBuffer, output = result.PixelBuffer;
        for (int y = 0; y < source.Height; y++) {
            token.ThrowIfCancellationRequested();
            for (int x = 0; x < source.Width; x++) {
                if ((x & 1023) == 0) token.ThrowIfCancellationRequested();
                int offset = (y * source.Width + x) * 4;
                OfficeColor color = operation(OfficeColor.FromRgba(input[offset], input[offset + 1], input[offset + 2], input[offset + 3]), x, y);
                output[offset] = color.R; output[offset + 1] = color.G; output[offset + 2] = color.B; output[offset + 3] = color.A;
            }
        }
        return result;
    }

    private static byte Channel(double value) => (byte)Math.Max(0D, Math.Min(255D, Math.Round(value)));
    private static double Luminance(OfficeColor color) => color.R * .2126D + color.G * .7152D + color.B * .0722D;
    private static OfficeColor Rgb(double r, double g, double b, byte alpha) => OfficeColor.FromRgba(Channel(r), Channel(g), Channel(b), alpha);
}
