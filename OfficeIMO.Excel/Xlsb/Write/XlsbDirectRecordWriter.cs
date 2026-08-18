using System.Buffers;
using OfficeIMO.Excel.Xlsb.Biff12;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>
    /// Writes the dense direct-tabular BIFF12 lane through one reusable buffer.
    /// This avoids issuing a separate virtual stream call for every primitive byte.
    /// </summary>
    internal sealed class XlsbDirectRecordWriter : IDisposable {
        private const int BufferSize = 32 * 1024;

        private readonly Stream _stream;
        private byte[]? _buffer;
        private int _position;

        internal XlsbDirectRecordWriter(Stream stream) {
            _stream = stream ?? throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("The BIFF12 destination must be writable.", nameof(stream));
            _buffer = ArrayPool<byte>.Shared.Rent(BufferSize);
        }

        internal void WriteRecord(int recordType) => WriteHeader(recordType, payloadLength: 0);

        internal void WriteHeader(int recordType, int payloadLength) {
            byte[] buffer = EnsureAvailable(6);
            _position += XlsbRecordWriter.EncodeHeader(recordType, payloadLength, buffer, _position);
        }

        internal void WriteRowHeader(int recordType, int zeroBasedRow, int columnCount, byte[] defaultRowProperties) {
            int spanCount = checked((columnCount + 1023) / 1024);
            int payloadLength = checked(17 + spanCount * 8);
            byte[] buffer = EnsureAvailable(checked(6 + payloadLength));
            int offset = _position;
            offset += XlsbRecordWriter.EncodeHeader(recordType, payloadLength, buffer, offset);
            offset = AppendUInt32(buffer, offset, checked((uint)zeroBasedRow));
            Buffer.BlockCopy(defaultRowProperties, 0, buffer, offset, defaultRowProperties.Length);
            offset += defaultRowProperties.Length;
            offset = AppendUInt32(buffer, offset, checked((uint)spanCount));
            for (int span = 0; span < spanCount; span++) {
                uint first = checked((uint)(span * 1024));
                uint last = checked((uint)Math.Min(columnCount - 1, ((span + 1) * 1024) - 1));
                offset = AppendUInt32(buffer, offset, first);
                offset = AppendUInt32(buffer, offset, last);
            }
            _position = offset;
        }

        internal void WriteTextCell(int recordType, int zeroBasedColumn, string value) {
            int payloadLength = checked(12 + value.Length * 2);
            byte[] buffer = GetBuffer();
            int maximumRecordLength = checked(6 + payloadLength);
            if (maximumRecordLength > buffer.Length) {
                Flush();
                int headerLength = XlsbRecordWriter.EncodeHeader(recordType, payloadLength, buffer);
                _stream.Write(buffer, 0, headerLength);
                WriteUInt32(checked((uint)zeroBasedColumn));
                WriteUInt32(0U);
                WriteWideString(value);
                return;
            }

            buffer = EnsureAvailable(maximumRecordLength);
            int offset = _position;
            offset += XlsbRecordWriter.EncodeHeader(recordType, payloadLength, buffer, offset);
            offset = AppendUInt32(buffer, offset, checked((uint)zeroBasedColumn));
            offset = AppendUInt32(buffer, offset, 0U);
            offset = AppendUInt32(buffer, offset, checked((uint)value.Length));
            for (int index = 0; index < value.Length; index++) {
                ushort character = value[index];
                buffer[offset++] = (byte)character;
                buffer[offset++] = (byte)(character >> 8);
            }
            _position = offset;
        }

        internal void WriteNumberCell(int recordType, int zeroBasedColumn, double value) {
            byte[] buffer = EnsureAvailable(22);
            int offset = _position;
            offset += XlsbRecordWriter.EncodeHeader(recordType, payloadLength: 16, buffer, offset);
            offset = AppendUInt32(buffer, offset, checked((uint)zeroBasedColumn));
            offset = AppendUInt32(buffer, offset, 0U);
            ulong bits = unchecked((ulong)BitConverter.DoubleToInt64Bits(value));
            offset = AppendUInt64(buffer, offset, bits);
            _position = offset;
        }

        internal void WriteBooleanCell(int recordType, int zeroBasedColumn, bool value) {
            byte[] buffer = EnsureAvailable(15);
            int offset = _position;
            offset += XlsbRecordWriter.EncodeHeader(recordType, payloadLength: 9, buffer, offset);
            offset = AppendUInt32(buffer, offset, checked((uint)zeroBasedColumn));
            offset = AppendUInt32(buffer, offset, 0U);
            buffer[offset++] = value ? (byte)1 : (byte)0;
            _position = offset;
        }

        internal void WriteUInt32(uint value) {
            byte[] buffer = EnsureAvailable(sizeof(uint));
            _position = AppendUInt32(buffer, _position, value);
        }

        private void WriteWideString(string value) {
            if (value == null) throw new ArgumentNullException(nameof(value));
            WriteUInt32(checked((uint)value.Length));

            byte[] buffer = GetBuffer();
            int charsPerChunk = buffer.Length / 2;
            for (int offset = 0; offset < value.Length; offset += charsPerChunk) {
                int characterCount = Math.Min(charsPerChunk, value.Length - offset);
                int byteCount = checked(characterCount * 2);
                buffer = EnsureAvailable(byteCount);
                int targetOffset = _position;
                for (int index = 0; index < characterCount; index++) {
                    ushort character = value[offset + index];
                    int byteOffset = targetOffset + (index * 2);
                    buffer[byteOffset] = (byte)character;
                    buffer[byteOffset + 1] = (byte)(character >> 8);
                }
                _position += byteCount;
            }
        }

        public void Dispose() {
            byte[]? buffer = _buffer;
            _buffer = null;
            if (buffer != null) {
                try {
                    if (_position != 0) _stream.Write(buffer, 0, _position);
                } finally {
                    _position = 0;
                    ArrayPool<byte>.Shared.Return(buffer, clearArray: true);
                }
            }
        }

        private byte[] EnsureAvailable(int count) {
            byte[] buffer = GetBuffer();
            if (count > buffer.Length) {
                throw new ArgumentOutOfRangeException(nameof(count), "The BIFF12 direct-write chunk exceeds its buffer.");
            }
            if (buffer.Length - _position < count) Flush();
            return buffer;
        }

        internal void Flush() {
            if (_position == 0) return;
            _stream.Write(GetBuffer(), 0, _position);
            _position = 0;
        }

        private byte[] GetBuffer() =>
            _buffer ?? throw new ObjectDisposedException(nameof(XlsbDirectRecordWriter));

        private static int AppendUInt32(byte[] buffer, int offset, uint value) {
            buffer[offset++] = (byte)value;
            buffer[offset++] = (byte)(value >> 8);
            buffer[offset++] = (byte)(value >> 16);
            buffer[offset++] = (byte)(value >> 24);
            return offset;
        }

        private static int AppendUInt64(byte[] buffer, int offset, ulong value) {
            offset = AppendUInt32(buffer, offset, unchecked((uint)value));
            return AppendUInt32(buffer, offset, unchecked((uint)(value >> 32)));
        }
    }
}
