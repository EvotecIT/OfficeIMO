namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffRkNumberReader {
        internal static double ReadRkNumber(uint rawValue) {
            bool divideBy100 = (rawValue & 0x01) != 0;
            bool isInteger = (rawValue & 0x02) != 0;
            double value;

            if (isInteger) {
                int integerValue = (int)(rawValue >> 2);
                if ((integerValue & 0x20000000) != 0) {
                    integerValue |= unchecked((int)0xc0000000);
                }

                value = integerValue;
            } else {
                ulong doubleBits = ((ulong)(rawValue & 0xfffffffc)) << 32;
                value = BitConverter.ToDouble(BitConverter.GetBytes(doubleBits), 0);
            }

            return divideBy100 ? value / 100d : value;
        }
    }
}
