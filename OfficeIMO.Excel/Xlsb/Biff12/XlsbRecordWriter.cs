namespace OfficeIMO.Excel.Xlsb.Biff12 {
    /// <summary>Writes canonically framed BIFF12 records.</summary>
    internal static class XlsbRecordWriter {
        internal static void Write(Stream stream, int recordType, byte[]? payload = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("The BIFF12 destination must be writable.", nameof(stream));
            if (recordType < 0 || recordType > 0x3FFF) throw new ArgumentOutOfRangeException(nameof(recordType));
            byte[] data = payload ?? Array.Empty<byte>();

            if (recordType < 0x80) {
                stream.WriteByte((byte)recordType);
            } else {
                stream.WriteByte((byte)((recordType & 0x7F) | 0x80));
                stream.WriteByte((byte)(recordType >> 7));
            }

            WriteVariableLengthValue(stream, data.Length);
            stream.Write(data, 0, data.Length);
        }

        internal static byte[] Encode(int recordType, byte[]? payload = null) {
            using var stream = new MemoryStream();
            Write(stream, recordType, payload);
            return stream.ToArray();
        }

        private static void WriteVariableLengthValue(Stream stream, int value) {
            if (value < 0 || value > 0x0FFFFFFF) throw new ArgumentOutOfRangeException(nameof(value));
            do {
                byte current = (byte)(value & 0x7F);
                value >>= 7;
                if (value != 0) current |= 0x80;
                stream.WriteByte(current);
            } while (value != 0);
        }
    }
}
