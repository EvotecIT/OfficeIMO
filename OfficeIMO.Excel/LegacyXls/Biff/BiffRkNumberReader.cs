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
    }
}
