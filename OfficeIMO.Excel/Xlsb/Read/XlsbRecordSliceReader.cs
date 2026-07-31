using OfficeIMO.Excel.Xlsb.Biff12;
using System.Buffers;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Excel.Xlsb.Read {
    /// <summary>
    /// Frames BIFF12 records over one decompressed package part without allocating a payload
    /// array for every record.
    /// </summary>
    internal sealed class XlsbRecordSliceReader {
        private readonly byte[] _bytes;
        private readonly int _length;
        private readonly int _maxRecordBytes;
        private readonly XlsbRecordReadBudget _budget;

        internal XlsbRecordSliceReader(
            byte[] bytes,
            int maxRecordBytes,
            XlsbRecordReadBudget budget,
            int? length = null) {
            _bytes = bytes ?? throw new ArgumentNullException(nameof(bytes));
            _length = length ?? bytes.Length;
            if (_length < 0 || _length > bytes.Length) {
                throw new ArgumentOutOfRangeException(nameof(length));
            }
            _maxRecordBytes = maxRecordBytes;
            _budget = budget ?? throw new ArgumentNullException(nameof(budget));
        }

        internal int Position { get; set; }

        internal bool TryRead(out XlsbRecordSlice record) {
            if (Position == _length) {
                record = default;
                return false;
            }

            if (Position < 0 || Position > _length) {
                throw new InvalidDataException("The BIFF12 record cursor is outside its package part.");
            }

            int recordOffset = Position;
            int firstTypeByte = ReadRequiredByte("record type");
            int type = firstTypeByte & 0x7F;
            if ((firstTypeByte & 0x80) != 0) {
                int secondTypeByte = ReadRequiredByte("record type");
                type |= (secondTypeByte & 0x7F) << 7;
                if (type < 128) {
                    throw new InvalidDataException("The BIFF12 record type uses a non-canonical two-byte encoding.");
                }
            }

            int size = ReadVariableLengthValue();
            if (size > _maxRecordBytes) {
                throw new InvalidDataException(
                    $"The BIFF12 record at offset {recordOffset} declares {size} payload bytes, exceeding the configured limit of {_maxRecordBytes} bytes.");
            }

            if (size > _length - Position) {
                throw new EndOfStreamException(
                    $"The BIFF12 record at offset {recordOffset} declares {size} payload bytes but only {_length - Position} remain.");
            }

            int payloadOffset = Position;
            Position += size;
            _budget.Consume();
            record = new XlsbRecordSlice(_bytes, recordOffset, type, payloadOffset, size);
            return true;
        }

        private int ReadVariableLengthValue() {
            int value = 0;
            for (int index = 0; index < 4; index++) {
                int current = ReadRequiredByte("record size");
                value |= (current & 0x7F) << (index * 7);
                if ((current & 0x80) == 0) {
                    return value;
                }
            }

            throw new InvalidDataException("The BIFF12 record size header is invalid.");
        }

        private int ReadRequiredByte(string fieldName) {
            if (Position >= _length) {
                throw new EndOfStreamException($"The BIFF12 stream ended inside the {fieldName} header.");
            }

            return _bytes[Position++];
        }
    }

    internal sealed class XlsbStreamRecordSliceReader : IDisposable {
        private const int InputBufferSize = 128 * 1024;
        private readonly Stream _stream;
        private readonly int _maxRecordBytes;
        private readonly XlsbRecordReadBudget _budget;
        private readonly bool _consumeRecordBudget;
        private readonly bool _leaveOpen;
        private byte[] _inputBuffer;
        private byte[] _payloadBuffer = new byte[256];
        private int _inputOffset;
        private int _inputLength;
        private int _offset;
        private bool _disposed;

        internal XlsbStreamRecordSliceReader(
            Stream stream,
            int maxRecordBytes,
            XlsbRecordReadBudget budget,
            bool leaveOpen = false,
            bool consumeRecordBudget = true) {
            _stream = stream ?? throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) {
                throw new ArgumentException("The BIFF12 stream must be readable.", nameof(stream));
            }

            _maxRecordBytes = maxRecordBytes;
            _budget = budget ?? throw new ArgumentNullException(nameof(budget));
            _consumeRecordBudget = consumeRecordBudget;
            _leaveOpen = leaveOpen;
            _inputBuffer = ArrayPool<byte>.Shared.Rent(InputBufferSize);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal bool TryRead(out XlsbRecordSlice record) {
            int recordOffset = _offset;
            if (!TryReadByte(out int firstTypeByte)) {
                record = default;
                return false;
            }
            int type = firstTypeByte & 0x7F;
            if ((firstTypeByte & 0x80) != 0) {
                int secondTypeByte = ReadRequiredByte("record type");
                type |= (secondTypeByte & 0x7F) << 7;
                if (type < 128) {
                    throw new InvalidDataException("The BIFF12 record type uses a non-canonical two-byte encoding.");
                }
            }

            int size = ReadVariableLengthValue();
            if (size > _maxRecordBytes) {
                throw new InvalidDataException(
                    $"The BIFF12 record at offset {recordOffset} declares {size} payload bytes, exceeding the configured limit of {_maxRecordBytes} bytes.");
            }

            ReadPayload(size, recordOffset, out byte[] bytes, out int payloadOffset);
            if (_consumeRecordBudget) {
                _budget.Consume();
            }
            record = new XlsbRecordSlice(bytes, recordOffset, type, payloadOffset, size);
            return true;
        }

        public void Dispose() {
            if (_disposed) {
                return;
            }

            _disposed = true;
            if (!_leaveOpen) {
                _stream.Dispose();
            }
            byte[] inputBuffer = _inputBuffer;
            _inputBuffer = Array.Empty<byte>();
            ArrayPool<byte>.Shared.Return(inputBuffer);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private int ReadVariableLengthValue() {
            int value = 0;
            for (int index = 0; index < 4; index++) {
                int current = ReadRequiredByte("record size");
                value |= (current & 0x7F) << (index * 7);
                if ((current & 0x80) == 0) {
                    return value;
                }
            }

            throw new InvalidDataException("The BIFF12 record size header is invalid.");
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private int ReadRequiredByte(string fieldName) {
            if (!TryReadByte(out int value)) {
                throw new EndOfStreamException($"The BIFF12 stream ended inside the {fieldName} header.");
            }

            return value;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private bool TryReadByte(out int value) {
            if (_inputOffset == _inputLength) {
                _inputLength = _stream.Read(_inputBuffer, 0, _inputBuffer.Length);
                _inputOffset = 0;
                if (_inputLength == 0) {
                    value = -1;
                    return false;
                }
            }

            value = _inputBuffer[_inputOffset++];
            _offset++;
            return true;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void ReadPayload(
            int size,
            int recordOffset,
            out byte[] bytes,
            out int payloadOffset) {
            int buffered = _inputLength - _inputOffset;
            if (size <= buffered) {
                bytes = _inputBuffer;
                payloadOffset = _inputOffset;
                _inputOffset += size;
                _offset = checked(_offset + size);
                return;
            }

            EnsurePayloadBuffer(size);
            int copied = 0;
            if (buffered > 0) {
                Array.Copy(_inputBuffer, _inputOffset, _payloadBuffer, 0, buffered);
                copied = buffered;
                _inputOffset = _inputLength;
            }

            while (copied < size) {
                int count = _stream.Read(_payloadBuffer, copied, size - copied);
                if (count == 0) {
                    throw new EndOfStreamException(
                        $"The BIFF12 record at offset {recordOffset} ended after {copied} of {size} payload bytes.");
                }

                copied += count;
            }

            _offset = checked(_offset + size);
            bytes = _payloadBuffer;
            payloadOffset = 0;
        }

        private void EnsurePayloadBuffer(int size) {
            if (size <= _payloadBuffer.Length) {
                return;
            }

            int capacity = _payloadBuffer.Length;
            while (capacity < size) {
                capacity = checked(capacity * 2);
            }

            _payloadBuffer = new byte[capacity];
        }

    }

    internal readonly struct XlsbRecordSlice {
        internal XlsbRecordSlice(byte[] bytes, int recordOffset, int type, int payloadOffset, int size) {
            Bytes = bytes;
            RecordOffset = recordOffset;
            Type = type;
            PayloadOffset = payloadOffset;
            Size = size;
        }

        internal byte[] Bytes { get; }

        internal int RecordOffset { get; }

        internal int Type { get; }

        internal int PayloadOffset { get; }

        internal int Size { get; }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal XlsbSliceCursor CreateCursor() => new(Bytes, PayloadOffset, Size);
    }

    internal struct XlsbSliceCursor {
        private readonly byte[] _bytes;
        private readonly int _end;

        internal XlsbSliceCursor(byte[] bytes, int offset, int length) {
            _bytes = bytes ?? throw new ArgumentNullException(nameof(bytes));
            if (offset < 0 || length < 0 || offset > bytes.Length - length) {
                throw new ArgumentOutOfRangeException(nameof(offset));
            }

            Position = offset;
            _end = offset + length;
        }

        internal int Position { get; private set; }

        internal int Remaining => _end - Position;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal byte ReadByte() {
            EnsureAvailable(1);
            return _bytes[Position++];
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal ushort ReadUInt16() {
            EnsureAvailable(2);
            int offset = Position;
            Position += 2;
            return (ushort)(_bytes[offset] | (_bytes[offset + 1] << 8));
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal uint ReadUInt32() {
            EnsureAvailable(4);
            int offset = Position;
            Position += 4;
            return (uint)(_bytes[offset]
                | (_bytes[offset + 1] << 8)
                | (_bytes[offset + 2] << 16)
                | (_bytes[offset + 3] << 24));
        }

        internal int ReadInt32() => unchecked((int)ReadUInt32());

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        internal double ReadDouble() {
            EnsureAvailable(8);
            int offset = Position;
            Position += 8;
            ulong bits = _bytes[offset]
                | ((ulong)_bytes[offset + 1] << 8)
                | ((ulong)_bytes[offset + 2] << 16)
                | ((ulong)_bytes[offset + 3] << 24)
                | ((ulong)_bytes[offset + 4] << 32)
                | ((ulong)_bytes[offset + 5] << 40)
                | ((ulong)_bytes[offset + 6] << 48)
                | ((ulong)_bytes[offset + 7] << 56);
            return BitConverter.Int64BitsToDouble(unchecked((long)bits));
        }

        internal string ReadWideString(int maxCharacters) {
            uint count = ReadUInt32();
            if (count > maxCharacters) {
                throw new InvalidDataException(
                    $"The BIFF12 string declares {count} characters, exceeding the configured limit of {maxCharacters} characters.");
            }

            int byteCount = checked((int)count * 2);
            EnsureAvailable(byteCount);
            string value = Encoding.Unicode.GetString(_bytes, Position, byteCount);
            Position += byteCount;
            return value;
        }

        internal void Skip(int count) {
            EnsureAvailable(count);
            Position += count;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void EnsureAvailable(int count) {
            if (count < 0 || count > Remaining) {
                throw new EndOfStreamException(
                    $"The BIFF12 payload ended at byte {Position}; {count} additional bytes were required but only {Remaining} remain.");
            }
        }
    }
}
