#if !NET5_0_OR_GREATER
using System.Runtime.InteropServices;
#endif

namespace OfficeIMO.Excel {
    /// <summary>
    /// Decides whether a number survives as an Excel number. Excel keeps 15 significant digits, so wider values are
    /// written as their exact invariant text instead of a numeric cell that every reader turns into a rounded double.
    /// </summary>
    internal static class ExcelExactNumber {
        private const long SignificantDigitLimit = 1_000_000_000_000_000L; // 10^15
        private const decimal DecimalSignificantDigitLimit = 1_000_000_000_000_000m;

        internal static bool FitsNumberCell(long value) => value > -SignificantDigitLimit && value < SignificantDigitLimit;

        internal static bool FitsNumberCell(ulong value) => value < (ulong)SignificantDigitLimit;

        /// <summary>Counts the coefficient's significant digits, ignoring trailing zeros, without allocating.</summary>
        internal static bool FitsNumberCell(decimal value) {
            GetCoefficient(value, out int low, out int middle, out int high);
            if (high == 0 && (((ulong)(uint)middle << 32) | (uint)low) < (ulong)SignificantDigitLimit) {
                return true;
            }

            decimal coefficient = new decimal(low, middle, high, false, 0);
            while (coefficient % 10m == 0m) {
                coefficient /= 10m;
            }

            return coefficient < DecimalSignificantDigitLimit;
        }

#if NET5_0_OR_GREATER
        private static void GetCoefficient(decimal value, out int low, out int middle, out int high) {
            Span<int> bits = stackalloc int[4];
            decimal.GetBits(value, bits);
            low = bits[0];
            middle = bits[1];
            high = bits[2];
        }
#else
        // decimal.GetBits allocates on these targets; both runtimes store flags, high, low and middle in this order.
        [StructLayout(LayoutKind.Explicit)]
        private struct DecimalLayout {
            [FieldOffset(0)] internal decimal Value;
            [FieldOffset(4)] internal int High;
            [FieldOffset(8)] internal int Low;
            [FieldOffset(12)] internal int Middle;
        }

        private static void GetCoefficient(decimal value, out int low, out int middle, out int high) {
            var layout = new DecimalLayout { Value = value };
            low = layout.Low;
            middle = layout.Middle;
            high = layout.High;
        }
#endif
    }
}
