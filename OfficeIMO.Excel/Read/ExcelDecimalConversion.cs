using System.Runtime.CompilerServices;

namespace OfficeIMO.Excel {
    /// <summary>Converts common binary worksheet numbers to decimal without the general IEEE conversion.</summary>
    internal static class ExcelDecimalConversion {
        /// <summary>
        /// Produces exactly <c>(decimal)value</c> when the double equals an integer number of hundredths
        /// within the Int32 range, which covers binary RK cells. Other values return false so callers keep
        /// the ordinary conversion. Equivalence was checked for every hundredth in that range.
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static bool TryConvertHundredths(double value, out decimal result) {
            double scaled = value * 100;
            if (Math.Abs(scaled) <= int.MaxValue) {
                long hundredths = (long)Math.Round(scaled);
                if (hundredths / 100.0 == value) {
                    if (hundredths == 0) {
                        result = 0m;
                        return true;
                    }

                    bool negative = hundredths < 0;
                    ulong magnitude = (ulong)(negative ? -hundredths : hundredths);
                    byte scale = 2;
                    if (magnitude % 10 == 0) {
                        magnitude /= 10;
                        scale = 1;
                        if (magnitude % 10 == 0) {
                            magnitude /= 10;
                            scale = 0;
                        }
                    }
                    result = new decimal(unchecked((int)(uint)magnitude), (int)(magnitude >> 32), 0, negative, scale);
                    return true;
                }
            }

            result = default;
            return false;
        }
    }
}
