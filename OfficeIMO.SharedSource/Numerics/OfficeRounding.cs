using System;

namespace OfficeIMO;

/// <summary>
/// Rounds a double to a number of fractional digits with the same result on every runtime.
/// </summary>
/// <remarks>
/// Up to .NET 10, and on .NET Framework, <see cref="Math.Round(double, int, MidpointRounding)"/> scales
/// the value by a power of ten and rounds the scaled double, so 2.675 rounds to 2.68 at two digits.
/// From .NET 11 it rounds the exact binary value, which lies just below 2.675, and returns 2.67. Values
/// written to documents must not change with the runtime, so this type keeps the scaling behaviour.
/// </remarks>
internal static class OfficeRounding {
    private static readonly double[] PowersOf10 = {
        1e0, 1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9, 1e10, 1e11, 1e12, 1e13, 1e14, 1e15
    };

    /// <summary>Rounds to <paramref name="digits"/> fractional digits, as Math.Round did before .NET 11.</summary>
    internal static double Round(double value, int digits, MidpointRounding mode = MidpointRounding.ToEven) {
        if ((uint)digits >= (uint)PowersOf10.Length) {
            throw new ArgumentOutOfRangeException(nameof(digits), "Rounding digits must be between 0 and 15, inclusive.");
        }

        // Doubles of this size have no fractional digits left to round.
        if (Math.Abs(value) >= 1e16) {
            return value;
        }

        double power = PowersOf10[digits];
        double scaled = value * power;
        switch (mode) {
            case MidpointRounding.ToEven:
                scaled = Math.Round(scaled);
                break;
            case MidpointRounding.AwayFromZero:
                double integral = Math.Truncate(scaled);
                double fraction = scaled - integral;
                scaled = Math.Abs(fraction) >= 0.5 ? integral + Math.Sign(fraction) : integral;
                break;
            default:
                return Math.Round(value, digits, mode);
        }

        return scaled / power;
    }
}
