using System;
using System.Runtime.CompilerServices;

namespace OfficeIMO;

/// <summary>
/// Converts binary floating-point values to decimal with the same result on every runtime.
/// </summary>
/// <remarks>
/// Up to .NET 10, and on .NET Framework, an explicit <c>(decimal)</c> conversion rounds a double to
/// 15 significant digits and a float to 7, so 0.1 becomes 0.1m. From .NET 11 the conversion keeps the
/// full binary value, so the same double becomes 0.1000000000000000055511151231m. Values read from
/// documents and user input must not change with the runtime, so this type reproduces the rounding
/// conversion arithmetically: the same digits, scale, sign and <see cref="OverflowException"/> cases,
/// without formatting or allocation.
/// </remarks>
internal static class OfficeDecimalConversion {
    private const int MaxScale = 28;

    private static readonly double[] DoublePowersOf10 = {
        1e0, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9,
        1e10, 1e11, 1e12, 1e13, 1e14, 1e15, 1e16, 1e17, 1e18, 1e19,
        1e20, 1e21, 1e22, 1e23, 1e24, 1e25, 1e26, 1e27, 1e28
    };

    private static readonly decimal[] DecimalPowersOf10 = {
        1m, 1e1m, 1e2m, 1e3m, 1e4m, 1e5m, 1e6m, 1e7m, 1e8m, 1e9m,
        1e10m, 1e11m, 1e12m, 1e13m, 1e14m, 1e15m, 1e16m, 1e17m, 1e18m, 1e19m,
        1e20m, 1e21m, 1e22m
    };

    /// <summary>Converts a double to decimal rounded to 15 significant digits.</summary>
    /// <exception cref="OverflowException">The value is NaN, infinite, or outside the decimal range.</exception>
    internal static decimal FromDouble(double value) {
        if (!TryFromDouble(value, out decimal result)) {
            throw new OverflowException("Value was either too large or too small for a Decimal.");
        }

        return result;
    }

    /// <summary>Converts a float to decimal rounded to 7 significant digits.</summary>
    /// <exception cref="OverflowException">The value is NaN, infinite, or outside the decimal range.</exception>
    internal static decimal FromSingle(float value) {
        if (!TryFromSingle(value, out decimal result)) {
            throw new OverflowException("Value was either too large or too small for a Decimal.");
        }

        return result;
    }

    /// <summary>
    /// Converts a value to decimal like <see cref="Convert.ToDecimal(object, IFormatProvider)"/>, rounding
    /// a boxed double or float the same way on every runtime.
    /// </summary>
    internal static decimal FromObject(object? value, IFormatProvider? provider) => value switch {
        double number => FromDouble(number),
        float number => FromSingle(number),
        _ => Convert.ToDecimal(value, provider)
    };

    /// <summary>
    /// Converts a value like <see cref="Convert.ChangeType(object, Type, IFormatProvider)"/>, rounding a
    /// double or float converted to decimal the same way on every runtime.
    /// </summary>
    internal static object? ChangeType(object? value, Type conversionType, IFormatProvider? provider) =>
        conversionType == typeof(decimal) && value is double or float
            ? FromObject(value, provider)
            : Convert.ChangeType(value, conversionType, provider);

    /// <summary>Converts a double to decimal rounded to 15 significant digits.</summary>
    /// <returns><see langword="false"/> when the value is NaN, infinite, or outside the decimal range.</returns>
    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static bool TryFromDouble(double value, out decimal result) {
        // Biased binary exponent shifted so that 2^exponent is just above the value.
        int exponent = (int)((BitConverter.DoubleToInt64Bits(value) >> 52) & 0x7FF) - 1022;
        if (exponent < -94) {
            // Even scaled by 10^28 the value stays below 0.5 (this includes zero and subnormals).
            result = 0m;
            return true;
        }

        if (exponent > 96) {
            result = default;
            return false;
        }

        return TryRound(Math.Abs(value), value < 0, exponent, digits: 15, out result);
    }

    /// <summary>Converts a float to decimal rounded to 7 significant digits.</summary>
    /// <returns><see langword="false"/> when the value is NaN, infinite, or outside the decimal range.</returns>
    internal static bool TryFromSingle(float value, out decimal result) {
        // Widening is exact, and for a normal float the double exponent matches the float one; smaller
        // floats and non-finite values land in the same zero and overflow cases.
        double widened = value;
        int exponent = (int)((BitConverter.DoubleToInt64Bits(widened) >> 52) & 0x7FF) - 1022;
        if (exponent < -94) {
            result = 0m;
            return true;
        }

        if (exponent > 96) {
            result = default;
            return false;
        }

        return TryRound(Math.Abs(widened), value < 0, exponent, digits: 7, out result);
    }

    private static bool TryRound(double magnitude, bool negative, int exponent, int digits, out decimal result) {
        // Scale the value to an integer of the requested number of significant digits. The power is
        // estimated from the binary exponent with log10(2) * 2^16 = 19728.
        double upper = DoublePowersOf10[digits];
        double lower = DoublePowersOf10[digits - 1];
        int power = digits - 1 - ((exponent * 19728) >> 16);
        if (power >= 0) {
            if (power > MaxScale) {
                power = MaxScale;
            }

            magnitude *= DoublePowersOf10[power];
        } else if (power != -1 || magnitude >= upper) {
            magnitude /= DoublePowersOf10[-power];
        } else {
            power = 0;
        }

        if (magnitude < lower && power < MaxScale) {
            magnitude *= 10;
            power++;
        }

        // Round half to even.
        ulong mantissa = (ulong)magnitude;
        double fraction = magnitude - mantissa;
        if (fraction > 0.5 || (fraction == 0.5 && (mantissa & 1) != 0)) {
            mantissa++;
        }

        if (mantissa == 0) {
            result = 0m;
            return true;
        }

        if (power < 0) {
            // Large values: the integer times a power of ten, at scale 0. Any value below 2^96 rounded to
            // 15 (or 7) digits stays within the decimal range, so the product cannot overflow.
            decimal integer = new decimal(mantissa) * DecimalPowersOf10[-power];
            result = negative ? -integer : integer;
            return true;
        }

        // Drop trailing zeros the scaling introduced, never below scale 0.
        int removable = Math.Min(power, digits - 1);
        while (removable > 0 && mantissa % 10 == 0) {
            mantissa /= 10;
            power--;
            removable--;
        }

        result = new decimal(unchecked((int)(uint)mantissa), unchecked((int)(uint)(mantissa >> 32)), 0, negative, (byte)power);
        return true;
    }
}
