using OfficeIMO.Excel.Xlsb.Biff12;

namespace OfficeIMO.Excel.Xlsb.Write;

/// <summary>Locates name operands in known native formula carriers without changing the surrounding record layout.</summary>
internal static partial class XlsbFormulaRecordRebinder {
    internal static byte[] Rebind(int type, byte[] original, IReadOnlyDictionary<uint, uint?> names) {
        if (type is 64 or 463 or 471) return RebindValidationOrCondition(type, original, names);
        int parsedFormula;
        switch (type) {
            case 8: parsedFormula = checked(12 + (int)XlsbFormulaTokenScanner.ReadUInt32(original, 8) * 2 + 2); break;
            case 9: parsedFormula = 18; break;
            case 10 or 11: parsedFormula = 11; break;
            case 39: parsedFormula = checked(13 + (int)XlsbFormulaTokenScanner.ReadUInt32(original, 9) * 2); break;
            case 426: parsedFormula = 17; break;
            case 427: parsedFormula = 16; break;
            case 351 or 352: parsedFormula = 1; break;
            default: return original;
        }
        byte[] data = (byte[])original.Clone();
        int count = checked((int)XlsbFormulaTokenScanner.ReadUInt32(data, parsedFormula));
        XlsbFormulaTokenScanner.RewriteNames(data, parsedFormula + 4, count, names);
        return data;
    }

    internal static byte[] RebindCellFormulaPayload(byte[] original, IReadOnlyDictionary<uint, uint?> names) {
        byte[] payload = (byte[])original.Clone(); int count = checked((int)XlsbFormulaTokenScanner.ReadUInt32(payload, 2));
        XlsbFormulaTokenScanner.RewriteNames(payload, 6, count, names); return payload;
    }

    internal static byte[] RebindPart(byte[] original, IReadOnlyDictionary<uint, uint?> names, long maximumBytes) {
        if (!names.Any(item => item.Value != item.Key)) return original;
        using var input = new MemoryStream(original); var records = XlsbRecordReader.ReadAll(input);
        using var output = new OfficeIMO.Core.Internal.OfficeBoundedMemoryStream(maximumBytes);
        int futureDepth = 0;
        foreach (var record in records) {
            if (record.Type == 35) futureDepth++;
            if (record.Type == 36 && --futureDepth < 0) throw new InvalidDataException("Unbalanced native XLSB future-record block.");
            byte[] data = futureDepth > 0 && record.Type is not (35 or 36 or 37 or 38) && record.Data.Length >= 4
                ? RebindFutureHeader(record.Data, names) : Rebind(record.Type, record.Data, names);
            XlsbRecordWriter.Write(output, record.Type, data);
        }
        if (futureDepth != 0) throw new InvalidDataException("Unbalanced native XLSB future-record block.");
        return output.ToArray();
    }

    private static byte[] RebindValidationOrCondition(int type, byte[] original, IReadOnlyDictionary<uint, uint?> names) {
        byte[] data = (byte[])original.Clone(); int offset;
        if (type == 64) {
            int ranges = checked((int)XlsbFormulaTokenScanner.ReadUInt32(data, 4));
            offset = checked(8 + ranges * 16);
            for (int index = 0; index < 4; index++) SkipString(data, ref offset);
            Parsed(data, ref offset, names); Parsed(data, ref offset, names);
        } else if (type == 463) {
            offset = 42; SkipString(data, ref offset);
            for (int index = 0; index < 3; index++)
                if (XlsbFormulaTokenScanner.ReadUInt32(data, 30 + index * 4) != 0) Parsed(data, ref offset, names);
        } else {
            offset = 24;
            if (XlsbFormulaTokenScanner.ReadUInt32(data, 20) != 0) Parsed(data, ref offset, names);
        }
        return data;
    }

    private static void Parsed(byte[] data, ref int offset, IReadOnlyDictionary<uint, uint?> names) {
        int count = checked((int)XlsbFormulaTokenScanner.ReadUInt32(data, offset)); offset = checked(offset + 4);
        XlsbFormulaTokenScanner.RewriteNames(data, offset, count, names); offset = checked(offset + count);
        int extra = checked((int)XlsbFormulaTokenScanner.ReadUInt32(data, offset)); offset = checked(offset + 4 + extra);
        if (offset > data.Length) throw new InvalidDataException("A native XLSB formula extra-data range is truncated.");
    }

    private static void SkipString(byte[] data, ref int offset) {
        uint count = XlsbFormulaTokenScanner.ReadUInt32(data, offset); offset = checked(offset + 4);
        if (count != uint.MaxValue) offset = checked(offset + checked((int)count * 2));
        if (offset > data.Length) throw new InvalidDataException("A native XLSB formula carrier string is truncated.");
    }
}
