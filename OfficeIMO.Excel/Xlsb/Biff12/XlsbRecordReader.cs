namespace OfficeIMO.Excel.Xlsb.Biff12 {
    /// <summary>
    /// Frames BIFF12 records without interpreting their record-specific payloads.
    /// </summary>
    internal static class XlsbRecordReader {
        internal const int DefaultMaxRecordBytes = 64 * 1024 * 1024;

        /// <summary>Reads all records from a BIFF12 part while enforcing a per-record allocation limit.</summary>
        internal static IReadOnlyList<XlsbRecord> ReadAll(
            Stream stream,
            int maxRecordBytes = DefaultMaxRecordBytes,
            XlsbRecordReadBudget? budget = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("The BIFF12 stream must be readable.", nameof(stream));
            if (maxRecordBytes < 0) throw new ArgumentOutOfRangeException(nameof(maxRecordBytes));

            var records = new List<XlsbRecord>();
            while (TryRead(stream, maxRecordBytes, out XlsbRecord record)) {
                budget?.Consume();
                records.Add(record);
            }

            return records.AsReadOnly();
        }

        /// <summary>Reads the next BIFF12 record, returning false only at a clean record boundary at end of stream.</summary>
        internal static bool TryRead(Stream stream, int maxRecordBytes, out XlsbRecord record) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("The BIFF12 stream must be readable.", nameof(stream));
            if (maxRecordBytes < 0) throw new ArgumentOutOfRangeException(nameof(maxRecordBytes));

            long offset = stream.CanSeek ? stream.Position : -1;
            int firstTypeByte = stream.ReadByte();
            if (firstTypeByte < 0) {
                record = null!;
                return false;
            }

            int headerSize = 1;
            int type = firstTypeByte & 0x7F;
            if ((firstTypeByte & 0x80) != 0) {
                int secondTypeByte = ReadRequiredByte(stream, "record type");
                headerSize++;
                type |= (secondTypeByte & 0x7F) << 7;
                if (type < 128) {
                    throw new InvalidDataException("The BIFF12 record type uses a non-canonical two-byte encoding.");
                }
            }

            int size = ReadVariableLengthValue(stream, ref headerSize);
            if (size > maxRecordBytes) {
                throw new InvalidDataException($"The BIFF12 record at offset {offset} declares {size} payload bytes, exceeding the configured limit of {maxRecordBytes} bytes.");
            }

            byte[] data = new byte[size];
            ReadExactly(stream, data, offset);
            record = new XlsbRecord(offset, headerSize, type, data);
            return true;
        }

        private static int ReadVariableLengthValue(Stream stream, ref int headerSize) {
            int value = 0;
            for (int index = 0; index < 4; index++) {
                int current = ReadRequiredByte(stream, "record size");
                headerSize++;
                value |= (current & 0x7F) << (index * 7);
                if ((current & 0x80) == 0 || index == 3) {
                    return value;
                }
            }

            throw new InvalidDataException("The BIFF12 record size header is invalid.");
        }

        private static int ReadRequiredByte(Stream stream, string fieldName) {
            int value = stream.ReadByte();
            if (value < 0) {
                throw new EndOfStreamException($"The BIFF12 stream ended inside the {fieldName} header.");
            }

            return value;
        }

        private static void ReadExactly(Stream stream, byte[] data, long recordOffset) {
            int read = 0;
            while (read < data.Length) {
                int count = stream.Read(data, read, data.Length - read);
                if (count == 0) {
                    throw new EndOfStreamException($"The BIFF12 record at offset {recordOffset} ended after {read} of {data.Length} payload bytes.");
                }

                read += count;
            }
        }
    }
}
