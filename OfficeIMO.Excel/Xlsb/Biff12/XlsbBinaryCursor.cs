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
            int offset = Position;
            Position += 8;
            ulong bits = _data[offset]
                | ((ulong)_data[offset + 1] << 8)
                | ((ulong)_data[offset + 2] << 16)
                | ((ulong)_data[offset + 3] << 24)
                | ((ulong)_data[offset + 4] << 32)
                | ((ulong)_data[offset + 5] << 40)
                | ((ulong)_data[offset + 6] << 48)
                | ((ulong)_data[offset + 7] << 56);
            return BitConverter.Int64BitsToDouble(unchecked((long)bits));
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
