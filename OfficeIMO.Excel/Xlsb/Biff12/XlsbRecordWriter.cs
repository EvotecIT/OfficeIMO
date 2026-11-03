namespace OfficeIMO.Excel.Xlsb.Biff12 {
    /// <summary>Writes canonically framed BIFF12 records.</summary>
    internal static class XlsbRecordWriter {
        internal static void Write(Stream stream, int recordType, byte[]? payload = null) {
            ArgumentNullException.ThrowIfNull(stream);
            if (!stream.CanWrite) throw new ArgumentException("The BIFF12 destination must be writable.", nameof(stream));
            ArgumentOutOfRangeException.ThrowIfNegative(recordType);
            ArgumentOutOfRangeException.ThrowIfGreaterThan(recordType, 0x3FFF);
            byte[] data = payload ?? Array.Empty<byte>();

            WriteHeader(stream, recordType, data.Length);
            stream.Write(data, 0, data.Length);
        }

        internal static void WriteHeader(Stream stream, int recordType, int payloadLength) {
            ArgumentNullException.ThrowIfNull(stream);
            if (!stream.CanWrite) throw new ArgumentException("The BIFF12 destination must be writable.", nameof(stream));
            ArgumentOutOfRangeException.ThrowIfNegative(recordType);
            ArgumentOutOfRangeException.ThrowIfGreaterThan(recordType, 0x3FFF);
            ArgumentOutOfRangeException.ThrowIfNegative(payloadLength);
            ArgumentOutOfRangeException.ThrowIfGreaterThan(payloadLength, 0x0FFFFFFF);

            if (recordType < 0x80) {
                stream.WriteByte((byte)recordType);
            } else {
                stream.WriteByte((byte)((recordType & 0x7F) | 0x80));
                stream.WriteByte((byte)(recordType >> 7));
            }

            WriteVariableLengthValue(stream, payloadLength);
        }

        internal static int EncodeHeader(int recordType, int payloadLength, byte[] destination) {
            return EncodeHeader(recordType, payloadLength, destination, 0);
        }

        internal static int EncodeHeader(
            int recordType,
            int payloadLength,
            byte[] destination,
            int offset) {
            ArgumentOutOfRangeException.ThrowIfNegative(recordType);
            ArgumentOutOfRangeException.ThrowIfGreaterThan(recordType, 0x3FFF);
            ArgumentOutOfRangeException.ThrowIfNegative(payloadLength);
            ArgumentOutOfRangeException.ThrowIfGreaterThan(payloadLength, 0x0FFFFFFF);
            ArgumentNullException.ThrowIfNull(destination);
            if (offset < 0 || offset > destination.Length - 6) {
                throw new ArgumentOutOfRangeException(nameof(offset));
            }

            int index = offset;
            if (recordType < 0x80) {
                destination[index++] = (byte)recordType;
            } else {
                destination[index++] = (byte)((recordType & 0x7F) | 0x80);
                destination[index++] = (byte)(recordType >> 7);
            }

            int value = payloadLength;
            do {
                byte current = (byte)(value & 0x7F);
                value >>= 7;
                if (value != 0) current |= 0x80;
                destination[index++] = current;
            } while (value != 0);

            return index - offset;
        }

        internal static byte[] Encode(int recordType, byte[]? payload = null) {
            using var stream = new MemoryStream();
            Write(stream, recordType, payload);
            return stream.ToArray();
        }

        private static void WriteVariableLengthValue(Stream stream, int value) {
            ArgumentOutOfRangeException.ThrowIfNegative(value);
            ArgumentOutOfRangeException.ThrowIfGreaterThan(value, 0x0FFFFFFF);
            do {
                byte current = (byte)(value & 0x7F);
                value >>= 7;
                if (value != 0) current |= 0x80;
                stream.WriteByte(current);
            } while (value != 0);
        }
    }
}
