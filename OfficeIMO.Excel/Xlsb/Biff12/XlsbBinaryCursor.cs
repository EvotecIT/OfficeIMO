namespace OfficeIMO.Excel.Xlsb.Biff12 {
    /// <summary>
    /// Provides bounds-checked little-endian reads over one BIFF12 record payload.
    /// </summary>
    internal sealed class XlsbBinaryCursor {
        private readonly byte[] _data;

        internal XlsbBinaryCursor(byte[] data) {
            _data = data ?? throw new ArgumentNullException(nameof(data));
        }

        internal int Position { get; private set; }

        internal int Remaining => _data.Length - Position;

        internal byte ReadByte() {
            EnsureAvailable(1);
            return _data[Position++];
        }

        internal ushort ReadUInt16() {
            EnsureAvailable(2);
            int offset = Position;
            Position += 2;
            return (ushort)(_data[offset] | (_data[offset + 1] << 8));
        }

        internal short ReadInt16() => unchecked((short)ReadUInt16());

        internal uint ReadUInt32() {
            EnsureAvailable(4);
            int offset = Position;
            Position += 4;
            return (uint)(_data[offset]
                | (_data[offset + 1] << 8)
                | (_data[offset + 2] << 16)
                | (_data[offset + 3] << 24));
        }

        internal int ReadInt32() => unchecked((int)ReadUInt32());

        internal double ReadDouble() {
            EnsureAvailable(8);
            byte[] bytes = new byte[8];
            Buffer.BlockCopy(_data, Position, bytes, 0, bytes.Length);
            Position += bytes.Length;
            return BitConverter.ToDouble(bytes, 0);
        }

        internal string ReadWideString(int maxCharacters) {
            if (maxCharacters < 0) throw new ArgumentOutOfRangeException(nameof(maxCharacters));

            uint count = ReadUInt32();
            if (count > maxCharacters) {
                throw new InvalidDataException($"The BIFF12 string declares {count} characters, exceeding the configured limit of {maxCharacters} characters.");
            }

            int byteCount;
            try {
                byteCount = checked((int)count * 2);
            } catch (OverflowException exception) {
                throw new InvalidDataException("The BIFF12 string length is too large.", exception);
            }

            EnsureAvailable(byteCount);
            string value = Encoding.Unicode.GetString(_data, Position, byteCount);
            Position += byteCount;
            return value;
        }

        internal byte[] ReadBytes(int count) {
            if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
            EnsureAvailable(count);
            byte[] bytes = new byte[count];
            Buffer.BlockCopy(_data, Position, bytes, 0, count);
            Position += count;
            return bytes;
        }

        internal void Skip(int count) {
            if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
            EnsureAvailable(count);
            Position += count;
        }

        private void EnsureAvailable(int count) {
            if (count < 0 || count > Remaining) {
                throw new EndOfStreamException($"The BIFF12 payload ended at byte {Position}; {count} additional bytes were required but only {Remaining} remain.");
            }
        }
    }
}
