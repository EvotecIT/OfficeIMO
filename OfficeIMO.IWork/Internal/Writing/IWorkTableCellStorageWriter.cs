namespace OfficeIMO.IWork.Internal;

/// <summary>Encodes the bounded scalar subset of modern iWork table cells.</summary>
internal static class IWorkTableCellStorageWriter {
    internal static byte[] Encode(IWorkTableCell? cell, IReadOnlyDictionary<string, uint> strings) {
        byte type = cell?.Kind switch { IWorkCellKind.Text => 3, IWorkCellKind.Number => 2,
            IWorkCellKind.Boolean => 6, _ => 0 };
        int length = type == 3 ? 16 : type == 2 ? 28 : type == 6 ? 20 : 12;
        var bytes = new byte[length];
        bytes[0] = 5; bytes[1] = type;
        if (type == 3) {
            WriteUInt32(bytes, 8, 1u << 3);
            WriteUInt32(bytes, 12, strings[(string)cell!.Value!]);
        } else if (type == 2) {
            // Apple's BNC numeric decoder requires Decimal128, even though the
            // shared source reader also understands older double-backed numbers.
            WriteUInt32(bytes, 8, 1u);
            byte[] scalar = IWorkDecimal128Writer.Encode((double)cell!.Value!);
            Buffer.BlockCopy(scalar, 0, bytes, 12, scalar.Length);
        } else if (type == 6) {
            WriteUInt32(bytes, 8, 1u << 1);
            double value = (bool)cell!.Value! ? 1d : 0d;
            byte[] scalar = BitConverter.GetBytes(value);
            if (!BitConverter.IsLittleEndian) Array.Reverse(scalar);
            Buffer.BlockCopy(scalar, 0, bytes, 12, scalar.Length);
        }
        return bytes;
    }

    internal static void WriteUInt32(byte[] bytes, int offset, uint value) {
        for (int i = 0; i < 4; i++) bytes[offset + i] = (byte)(value >> (8 * i));
    }

    internal static void WriteOffset(byte[] bytes, int column, int value) {
        if (value >= ushort.MaxValue) throw new InvalidDataException("An iWork scalar row exceeds the narrow cell-offset range.");
        bytes[column * 2] = (byte)value; bytes[column * 2 + 1] = (byte)(value >> 8);
    }
}
