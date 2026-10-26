using System;
using System.Globalization;
using System.Numerics;
using System.Text.RegularExpressions;

namespace OfficeIMO;

/// <summary>A signed document length stored as whole English Metric Units (EMUs).</summary>
/// <remarks>Fractional EMUs round to nearest, with midpoint ties away from zero. Arithmetic is checked.</remarks>
public readonly struct OfficeLength : IEquatable<OfficeLength>, IComparable<OfficeLength>, IComparable, IFormattable {
    private readonly long _emus;
    private OfficeLength(long emus) { _emus = emus; }
    private static readonly Regex Syntax = new Regex(@"^([+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?)\s*(pt|cm|mm|in|px|emu)?$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    /// <summary>Gets a zero length.</summary>
    public static OfficeLength Zero => default;
    /// <summary>Creates a length from whole EMUs.</summary>
    public static OfficeLength Emus(long value) => new OfficeLength(value);
    /// <summary>Creates a length in points (12,700 EMUs per point).</summary>
    public static OfficeLength Points(double value) => From(value, 12700);
    /// <summary>Creates a length in centimeters (360,000 EMUs per centimeter).</summary>
    public static OfficeLength Centimeters(double value) => From(value, 360000);
    /// <summary>Creates a length in millimeters (36,000 EMUs per millimeter).</summary>
    public static OfficeLength Millimeters(double value) => From(value, 36000);
    /// <summary>Creates a length in inches (914,400 EMUs per inch).</summary>
    public static OfficeLength Inches(double value) => From(value, 914400);
    /// <summary>Creates a length in pixels at 96 DPI (9,525 EMUs per pixel).</summary>
    public static OfficeLength Pixels(double value) => From(value, 9525);
    /// <summary>Returns the whole EMU value.</summary>
    public long ToEmus() => _emus;
    /// <summary>Returns the length in points.</summary>
    public double ToPoints() => _emus / 12700d;
    /// <summary>Returns the length in centimeters.</summary>
    public double ToCentimeters() => _emus / 360000d;
    /// <summary>Returns the length in millimeters.</summary>
    public double ToMillimeters() => _emus / 36000d;
    /// <summary>Returns the length in inches.</summary>
    public double ToInches() => _emus / 914400d;
    /// <summary>Returns the length in 96-DPI pixels.</summary>
    public double ToPixels() => _emus / 9525d;
    /// <summary>Rounds to whole twips at write time; midpoint ties round away from zero.</summary>
    public long ToTwips() => Round(_emus / 635m);
    /// <summary>Rounds to whole half-points at write time; midpoint ties round away from zero.</summary>
    public long ToHalfPoints() => Round(_emus / 6350m);

    /// <summary>Parses an invariant number and optional pt, cm, mm, in, px or emu suffix. A bare number means points.</summary>
    public static OfficeLength Parse(string value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        if (TryParse(value, out OfficeLength result)) return result;
        throw new FormatException("Expected an invariant length such as 12pt, 2.5cm or 914400emu within the EMU range.");
    }

    /// <summary>Tries to parse an invariant length. Commas, nonfinite values and overflowing lengths are rejected.</summary>
    public static bool TryParse(string? value, out OfficeLength result) {
        result = Zero;
        if (value == null) return false;
        Match match = Syntax.Match(value.Trim());
        if (!match.Success || !decimal.TryParse(match.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out decimal number)) return false;
        try {
            result = new OfficeLength(Round(number * Factor(match.Groups[2].Value)));
            return true;
        } catch (OverflowException) { return false; }
    }

    /// <summary>Formats exactly as whole EMUs for an invariant round trip.</summary>
    public override string ToString() => ToString(null, CultureInfo.InvariantCulture);
    /// <summary>Formats using a unit token (pt, cm, mm, in, px, emu), or G for exact EMUs. Output is always invariant.</summary>
    /// <remarks>The provider is intentionally ignored so the result remains valid input to <see cref="Parse"/>.</remarks>
    public string ToString(string? format, IFormatProvider? formatProvider) {
        string unit = string.IsNullOrEmpty(format) || string.Equals(format, "G", StringComparison.OrdinalIgnoreCase) ? "emu" : format!.ToLowerInvariant();
        decimal factor = Factor(unit);
        return (_emus / factor).ToString("G29", CultureInfo.InvariantCulture) + unit;
    }

    /// <inheritdoc />
    public bool Equals(OfficeLength other) => _emus == other._emus;
    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficeLength other && Equals(other);
    /// <inheritdoc />
    public override int GetHashCode() => _emus.GetHashCode();
    /// <inheritdoc />
    public int CompareTo(OfficeLength other) => _emus.CompareTo(other._emus);
    /// <inheritdoc />
    public int CompareTo(object? obj) => obj == null ? 1 : obj is OfficeLength other ? CompareTo(other) : throw new ArgumentException("Expected an OfficeLength.", nameof(obj));
    /// <summary>Adds lengths, throwing on overflow.</summary>
    public static OfficeLength operator +(OfficeLength left, OfficeLength right) => Emus(checked(left._emus + right._emus));
    /// <summary>Subtracts lengths, throwing on overflow.</summary>
    public static OfficeLength operator -(OfficeLength left, OfficeLength right) => Emus(checked(left._emus - right._emus));
    /// <summary>Multiplies a length by a finite scalar and rounds to whole EMUs.</summary>
    public static OfficeLength operator *(OfficeLength length, double scalar) => Emus(Scale(length._emus, scalar));
    /// <summary>Multiplies a length by a finite scalar and rounds to whole EMUs.</summary>
    public static OfficeLength operator *(double scalar, OfficeLength length) => length * scalar;
    /// <summary>Compares lengths for equality.</summary>
    public static bool operator ==(OfficeLength left, OfficeLength right) => left.Equals(right);
    /// <summary>Compares lengths for inequality.</summary>
    public static bool operator !=(OfficeLength left, OfficeLength right) => !left.Equals(right);
    /// <summary>Compares signed lengths.</summary>
    public static bool operator <(OfficeLength left, OfficeLength right) => left._emus < right._emus;
    /// <summary>Compares signed lengths.</summary>
    public static bool operator >(OfficeLength left, OfficeLength right) => left._emus > right._emus;
    /// <summary>Compares signed lengths.</summary>
    public static bool operator <=(OfficeLength left, OfficeLength right) => left._emus <= right._emus;
    /// <summary>Compares signed lengths.</summary>
    public static bool operator >=(OfficeLength left, OfficeLength right) => left._emus >= right._emus;

    private static OfficeLength From(double value, long factor) => Emus(Scale(factor, value));
    private static long Round(decimal value) => checked((long)decimal.Round(value, 0, MidpointRounding.AwayFromZero));
    // Decode the finite IEEE-754 value rather than narrowing it to decimal before rounding.
    // The numerator/denominator stay exact even next to a half EMU or the Int64 boundary.
    private static long Scale(long emus, double scalar) {
        if (double.IsNaN(scalar) || double.IsInfinity(scalar))
            throw new ArgumentOutOfRangeException(nameof(scalar), "Length and scalar values must be finite.");
        if (emus == 0 || scalar == 0) return 0;
        long bits = BitConverter.DoubleToInt64Bits(scalar);
        int exponentBits = (int)((bits >> 52) & 0x7ff);
        long significand = bits & 0x000fffffffffffffL;
        if (exponentBits != 0) significand |= 1L << 52;
        int shift = exponentBits == 0 ? -1074 : exponentBits - 1023 - 52;
        BigInteger numerator = (BigInteger)emus * significand;
        if (bits < 0) numerator = -numerator;
        if (shift >= 0) return checked((long)(numerator << shift));
        BigInteger denominator = BigInteger.One << -shift;
        BigInteger rounded = BigInteger.DivRem(BigInteger.Abs(numerator), denominator, out BigInteger remainder);
        if (remainder * 2 >= denominator) rounded++;
        return checked((long)(numerator.Sign < 0 ? -rounded : rounded));
    }
    private static decimal Factor(string unit) {
        switch (unit.ToLowerInvariant()) {
            case "": case "pt": return 12700;
            case "cm": return 360000;
            case "mm": return 36000;
            case "in": return 914400;
            case "px": return 9525;
            case "emu": return 1;
            default: throw new FormatException("Unknown length unit: " + unit);
        }
    }
}
