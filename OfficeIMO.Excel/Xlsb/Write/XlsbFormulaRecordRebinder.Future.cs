using OfficeIMO.Excel.Xlsb.Biff12;

namespace OfficeIMO.Excel.Xlsb.Write;

internal static partial class XlsbFormulaRecordRebinder {
    private static byte[] RebindFutureHeader(byte[] original, IReadOnlyDictionary<uint, uint?> names) {
        byte[] data = (byte[])original.Clone(); int offset = 4;
        uint flags = XlsbFormulaTokenScanner.ReadUInt32(data, 0);
        if ((flags & ~15U) != 0) throw new InvalidDataException("Invalid native XLSB future-record header flags.");
        if ((flags & 1) != 0) {
            int count = Count(data, ref offset, 20); offset = checked(offset + count * 20);
        }
        if ((flags & 2) != 0) {
            int count = Count(data, ref offset, 8);
            for (int index = 0; index < count; index++) {
                XlsbFormulaTokenScanner.ReadUInt32(data, offset); offset += 4;
                int ranges = Count(data, ref offset, 16); offset = checked(offset + ranges * 16);
            }
        }
        if ((flags & 4) != 0) {
            int count = Count(data, ref offset, 12);
            for (int index = 0; index < count; index++) {
                XlsbFormulaTokenScanner.ReadUInt32(data, offset); offset += 4;
                int tokens = checked((int)XlsbFormulaTokenScanner.ReadUInt32(data, offset));
                int extra = checked((int)XlsbFormulaTokenScanner.ReadUInt32(data, offset + 4)); offset += 8;
                XlsbFormulaTokenScanner.RewriteNames(data, offset, tokens, names);
                offset = checked(offset + tokens + extra);
                if (offset > data.Length) throw new InvalidDataException("A native XLSB future formula is truncated.");
            }
        }
        return data;
    }

    private static int Count(byte[] data, ref int offset, int itemMinimum) {
        uint count = XlsbFormulaTokenScanner.ReadUInt32(data, offset); offset += 4;
        if (count > (data.Length - offset) / itemMinimum) throw new InvalidDataException("A native XLSB future-record count exceeds its payload.");
        return checked((int)count);
    }
}
