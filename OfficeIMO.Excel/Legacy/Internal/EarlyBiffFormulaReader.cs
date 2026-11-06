using OfficeIMO.Excel.LegacyXls.Biff;

namespace OfficeIMO.Excel.Legacy;

/// <summary>Normalizes the generation-specific token widths before using the shared BIFF expression decoder.</summary>
internal static class EarlyBiffFormulaReader {
    internal static bool TryRead(byte[] source, int start, int length, ushort bof, Encoding encoding,
        int row, int column, out string? formula) {
        formula = null;
        if (length == 0 || length > 8192) return false;
        bool biff2 = bof == 0x0009;
        bool narrowFunctions = bof != 0x0409;
        using var normalized = new MemoryStream();
        using var output = new BinaryWriter(normalized, Encoding.Unicode, leaveOpen: true);
        output.Write((ushort)0);
        int position = start;
        int end = start + length;
        while (position < end) {
            byte token = source[position++];
            int kind = token < 0x20 ? token : (token & 0x1F) | 0x20;
            output.Write(token);
            switch (kind) {
                case >= 0x03 and <= 0x16:
                    break;
                case 0x17:
                    if (position == end) return false;
                    int characters = source[position++];
                    if (characters > end - position) return false;
                    string text = encoding.GetString(source, position, characters);
                    output.Write((byte)text.Length);
                    output.Write((byte)1);
                    output.Write(Encoding.Unicode.GetBytes(text));
                    position += characters;
                    break;
                case 0x19:
                    int width = biff2 ? 1 : 2;
                    if (position + 1 + width > end) return false;
                    byte attribute = source[position++];
                    int value = width == 1 ? source[position] : BiffRecordReader.ReadUInt16(source, position);
                    position += width;
                    output.Write(attribute);
                    output.Write((ushort)value);
                    if (attribute == 4) {
                        if (value > 29 || (value + 1) * width > end - position) return false;
                        for (int i = 0; i <= value; i++) {
                            output.Write((ushort)0);
                            position += width;
                        }
                    }
                    break;
                case 0x1C:
                case 0x1D:
                case 0x1E:
                case 0x1F:
                    int count = kind == 0x1F ? 8 : kind == 0x1E ? 2 : 1;
                    if (count > end - position) return false;
                    if (kind == 0x1D && source[position] > 1) return false;
                    if (kind == 0x1F && !IsFinite(BiffRecordReader.ReadDouble(source, position))) return false;
                    output.Write(source, position, count);
                    position += count;
                    break;
                case 0x21:
                case 0x22:
                    if (kind == 0x22) {
                        if (position == end || source[position] > 127) return false;
                        output.Write(source[position++]);
                    }
                    int functionWidth = narrowFunctions ? 1 : 2;
                    if (functionWidth > end - position) return false;
                    ushort function = narrowFunctions ? source[position] : BiffRecordReader.ReadUInt16(source, position);
                    if (function == 255 || (function & 0x8000) != 0) return false;
                    output.Write(function);
                    position += functionWidth;
                    break;
                case 0x24:
                case 0x25:
                case 0x2A:
                case 0x2B:
                    bool area = kind is 0x25 or 0x2B;
                    int referenceLength = area ? 6 : 3;
                    if (referenceLength > end - position) return false;
                    ushort firstRow = BiffRecordReader.ReadUInt16(source, position);
                    ushort lastRow = area ? BiffRecordReader.ReadUInt16(source, position + 2) : firstRow;
                    int columnOffset = position + (area ? 4 : 2);
                    output.Write((ushort)(firstRow & 0x3FFF));
                    if (area) output.Write((ushort)(lastRow & 0x3FFF));
                    output.Write((ushort)(source[columnOffset] | (firstRow & 0xC000)));
                    if (area) output.Write((ushort)(source[columnOffset + 1] | (lastRow & 0xC000)));
                    position += referenceLength;
                    break;
                case 0x26:
                case 0x27:
                case 0x28:
                case 0x29:
                    // Memory tokens carry evaluation bookkeeping, never a new expression operand.
                    int memoryLength = kind == 0x29 ? (biff2 ? 1 : 2) : (biff2 ? 4 : 6);
                    if (memoryLength > end - position) return false;
                    position += memoryLength;
                    normalized.SetLength(normalized.Length - 1);
                    normalized.Position = normalized.Length;
                    break;
                default:
                    return false;
            }
            if (normalized.Length - 2 > ushort.MaxValue) return false;
        }
        byte[] bytes = normalized.ToArray();
        bytes[0] = (byte)(bytes.Length - 2);
        bytes[1] = (byte)((bytes.Length - 2) >> 8);
        if (!BiffFormulaTextReader.TryRead(bytes, 0, row, column,
                Array.Empty<BiffExternSheetReference>(), Array.Empty<string>(), Array.Empty<string?>(), out formula)
            || formula!.Length > 8192) {
            formula = null;
            return false;
        }
        return true;
    }

    private static bool IsFinite(double number) => !double.IsNaN(number) && !double.IsInfinity(number);
}
