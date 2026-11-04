using System.Runtime.CompilerServices;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffRkNumberReader {
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal static double ReadRkNumber(uint rawValue) {
            bool divideBy100 = (rawValue & 0x01) != 0;
            bool isInteger = (rawValue & 0x02) != 0;
            double value;

            if (isInteger) {
                value = unchecked((int)rawValue) >> 2;
            } else {
                ulong doubleBits = ((ulong)(rawValue & 0xfffffffc)) << 32;
                value = BitConverter.Int64BitsToDouble(unchecked((long)doubleBits));
            }

            return divideBy100 ? value / 100d : value;
        }

        /// <summary>
        /// Encodes a number as a 32-bit RK value when one exists whose decoded value has the same
        /// IEEE bits: a 30-bit integer or the upper 30 bits of the double, optionally scaled by 100.
        /// </summary>
        internal static bool TryEncodeRkNumber(double value, out uint rawValue) {
            long bits = BitConverter.DoubleToInt64Bits(value);
            if (value >= -536_870_912d && value <= 536_870_911d) {
                int integer = (int)value;
                if (integer == value && bits != NegativeZeroBits) {
                    rawValue = unchecked((uint)(integer << 2)) | 0x02U;
                    return true;
                }
            }
            if ((bits & LowDoubleBitsMask) == 0) {
                rawValue = unchecked((uint)((ulong)bits >> 32)) & 0xFFFFFFFCU;
                return true;
            }

            double scaled = value * 100d;
            if (scaled >= -536_870_912d && scaled <= 536_870_911d) {
                int integer = (int)scaled;
                if (integer == scaled
                    && TryAccept(unchecked((uint)(integer << 2)) | 0x03U, bits, out rawValue)) {
                    return true;
                }
            }
            long scaledBits = BitConverter.DoubleToInt64Bits(scaled);
            if ((scaledBits & LowDoubleBitsMask) == 0
                && TryAccept((unchecked((uint)((ulong)scaledBits >> 32)) & 0xFFFFFFFCU) | 0x01U, bits, out rawValue)) {
                return true;
            }

            rawValue = 0;
            return false;
        }

        private const long LowDoubleBitsMask = 0x3_FFFF_FFFFL;
        private const long NegativeZeroBits = unchecked((long)0x8000_0000_0000_0000UL);

        private static bool TryAccept(uint candidate, long expectedBits, out uint rawValue) {
            rawValue = candidate;
            return BitConverter.DoubleToInt64Bits(ReadRkNumber(candidate)) == expectedBits;
        }
    }
}
