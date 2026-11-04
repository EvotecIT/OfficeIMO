using System;

namespace OfficeIMO;

public readonly partial struct OfficeColor {
    private OfficeColor(byte r, byte g, byte b, byte alpha, double normalizedR, double normalizedG, double normalizedB)
        : this(r, g, b, alpha) {
        _hasNormalizedRgb = true;
        _normalizedR = normalizedR;
        _normalizedG = normalizedG;
        _normalizedB = normalizedB;
    }

    /// <summary>Creates a literal color retaining normalized RGB channels and a byte alpha channel.</summary>
    /// <param name="r">Finite red channel in the inclusive 0..1 range.</param>
    /// <param name="g">Finite green channel in the inclusive 0..1 range.</param>
    /// <param name="b">Finite blue channel in the inclusive 0..1 range.</param>
    /// <param name="alpha">Alpha byte, from 0 (transparent) to 255 (opaque).</param>
    /// <returns>A literal color whose normalized RGB channels preserve the supplied values.</returns>
    /// <remarks>
    /// RGB byte getters round each channel multiplied by 255 to the nearest integer, with midpoint ties to even.
    /// Hexadecimal/CSS formatting uses that byte projection and does not round-trip normalized precision.
    /// Exact byte fractions (byte divided by 255D) use the same representation, equality and hash as byte literals.
    /// </remarks>
    /// <exception cref="ArgumentOutOfRangeException">An RGB channel is not finite or is outside 0..1.</exception>
    public static OfficeColor FromNormalizedRgb(double r, double g, double b, byte alpha = 255) {
        ValidateNormalizedRgbChannel(r, nameof(r));
        ValidateNormalizedRgbChannel(g, nameof(g));
        ValidateNormalizedRgbChannel(b, nameof(b));

        byte red = (byte)Math.Round(r * 255D);
        byte green = (byte)Math.Round(g * 255D);
        byte blue = (byte)Math.Round(b * 255D);
        if (r == red / 255D && g == green / 255D && b == blue / 255D) {
            return new OfficeColor(red, green, blue, alpha);
        }

        return new OfficeColor(red, green, blue, alpha, r, g, b);
    }

    /// <summary>Returns literal RGB channels in the 0..1 range without projecting retained normalized values to bytes.</summary>
    /// <returns>The retained normalized channels, or byte channels divided by 255D for byte literals.</returns>
    /// <remarks>Alpha remains available separately through the byte <see cref="A"/> property. Resolve theme roles first.</remarks>
    /// <exception cref="InvalidOperationException">This color is an unresolved theme role.</exception>
    public (double R, double G, double B) ToNormalizedRgb() {
        _ = LiteralComponent(_r);
        return !_hasNormalizedRgb
            ? (_r / 255D, _g / 255D, _b / 255D)
            : (_normalizedR, _normalizedG, _normalizedB);
    }

    private static void ValidateNormalizedRgbChannel(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D || value > 1D) {
            throw new ArgumentOutOfRangeException(paramName, value, "Normalized RGB channels must be finite values between 0 and 1.");
        }
    }

    // Exact byte fractions use the byte representation, so precise values cannot alias a byte-only color.
    private bool EqualsNormalizedRgb(OfficeColor other) =>
        _hasNormalizedRgb == other._hasNormalizedRgb
        && (!_hasNormalizedRgb || _normalizedR.Equals(other._normalizedR)
            && _normalizedG.Equals(other._normalizedG)
            && _normalizedB.Equals(other._normalizedB));

    private int GetNormalizedRgbHashCode(int hash) {
        unchecked {
            hash = (hash * 397) ^ _normalizedR.GetHashCode();
            hash = (hash * 397) ^ _normalizedG.GetHashCode();
            return (hash * 397) ^ _normalizedB.GetHashCode();
        }
    }
}
