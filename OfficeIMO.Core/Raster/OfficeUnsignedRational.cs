using System;

namespace OfficeIMO.Drawing;

/// <summary>Converts authored positive density to the unsigned TIFF/Exif rational representation.</summary>
internal static class OfficeUnsignedRational {
    internal static OfficeRational FromPositiveDouble(double value, string parameterName = "value") {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 1D / uint.MaxValue || value > uint.MaxValue) throw OutsideRange();
        // Continued-fraction convergents preserve small values without a fixed decimal
        // denominator. Both words remain bounded before multiplication. Reject values
        // that cannot achieve one part in a trillion instead of silently losing density.
        ulong priorNumerator = 0, numerator = 1, priorDenominator = 1, denominator = 0;
        double remainder = value;
        for (int step = 0; step < 64; step++) {
            double whole = Math.Floor(remainder);
            ulong limit = uint.MaxValue;
            if (numerator != 0) limit = Math.Min(limit, (uint.MaxValue - priorNumerator) / numerator);
            if (denominator != 0) limit = Math.Min(limit, (uint.MaxValue - priorDenominator) / denominator);
            // The largest feasible intermediate convergent can still represent
            // the value when the next full convergent exceeds either uint word.
            bool bounded = whole > limit;
            ulong coefficient = bounded ? limit : (ulong)whole;
            ulong nextNumerator = coefficient * numerator + priorNumerator;
            ulong nextDenominator = coefficient * denominator + priorDenominator;
            if (nextNumerator != 0 && nextDenominator != 0) {
                var result = new OfficeRational((uint)nextNumerator, (uint)nextDenominator);
                if (Math.Abs(result.ToDouble() - value) <= value * 1E-12) return result;
            }
            if (bounded) break;
            double fraction = remainder - whole;
            if (fraction == 0D) break;
            priorNumerator = numerator; numerator = nextNumerator;
            priorDenominator = denominator; denominator = nextDenominator;
            remainder = 1D / fraction;
        }
        throw OutsideRange();

        ArgumentOutOfRangeException OutsideRange() => new ArgumentOutOfRangeException(parameterName, "Resolution cannot be represented as a positive unsigned rational within one part in a trillion.");
    }
}
