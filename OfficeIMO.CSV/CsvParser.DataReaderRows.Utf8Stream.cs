#nullable enable

#if NET8_0_OR_GREATER
using System.Buffers;
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
using System.Text;
using System.Threading;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    /// <summary>
    /// Pull-based UTF-8 row source for ordinary unquoted records. Complex records are handed to
    /// the canonical text row source from the current record boundary so quote and comment
    /// behavior continues to have one implementation.
    /// </summary>
    internal sealed class CsvUtf8StreamDataReaderRowSource : ICsvDataReaderHeaderRowSource, ICsvDataReaderPositionSource
    {
        private const int BufferSize = 256 * 1024;
        private const int FallbackTextBufferSize = 128 * 1024;

        private readonly CsvLoadOptions _options;
        private readonly Encoding _encoding;
        private readonly byte _delimiter;
        private readonly byte _commentCharacter;
        private readonly Vector256<byte> _delimiterVector;
        private static readonly Vector256<byte> QuoteVector = Vector256.Create((byte)'"');
        private static readonly Vector256<byte> CarriageReturnVector = Vector256.Create((byte)'\r');
        private static readonly Vector256<byte> LineFeedVector = Vector256.Create((byte)'\n');
        private readonly bool _allowEmpty;
        private Stream _stream;
        private byte[] _buffer;
        private CsvUtf8DataReaderRowVisitor _visitor;
        private CsvStreamDataReaderRowSource? _fallback;
        private int _position;
        private int _length;
        private int _sourceColumnCount;
        private int _emittedRecordCount;
        private int _physicalLineSeparatorsConsumed;
        private int? _currentPhysicalLineNumber;
        private int? _currentPhysicalEndLineNumber;
        private bool _endOfStream;
        private bool _skipLineFeedAfterCarriageReturn;
        private bool _streamTransferred;
        private bool _disposed;

        private CsvUtf8StreamDataReaderRowSource(
            Stream stream,
            CsvLoadOptions options,
            Encoding encoding,
            byte[] buffer,
            int length,
            int position)
        {
            _stream = stream;
            _options = options;
            _encoding = encoding;
            _delimiter = checked((byte)GetDelimiterChar(options));
            _commentCharacter = checked((byte)options.CommentCharacter);
            _delimiterVector = Vector256.Create(_delimiter);
            _allowEmpty = options.AllowEmptyLines;
            _buffer = buffer;
            _length = length;
            _position = position;
            _visitor = new CsvUtf8DataReaderRowVisitor(buffer, encoding);
        }

        internal static bool TryCreate(
            Stream stream,
            CsvLoadOptions options,
            out CsvUtf8StreamDataReaderRowSource? rows)
        {
            byte[] buffer = ArrayPool<byte>.Shared.Rent(BufferSize);
            int length = 0;
            try
            {
                while (length < 4)
                {
                    options.CancellationToken.ThrowIfCancellationRequested();
                    int read = stream.Read(buffer, length, buffer.Length - length);
                    if (read == 0)
                    {
                        break;
                    }

                    length += read;
                }

                if (HasNonUtf8Preamble(buffer.AsSpan(0, length)))
                {
                    ArrayPool<byte>.Shared.Return(buffer);
                    rows = null;
                    return false;
                }

                int position = length >= 3 &&
                    buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF
                    ? 3
                    : 0;
                rows = new CsvUtf8StreamDataReaderRowSource(
                    stream,
                    options,
                    options.Encoding ?? Encoding.UTF8,
                    buffer,
                    length,
                    position);
                return true;
            }
            catch
            {
                ArrayPool<byte>.Shared.Return(buffer);
                stream.Dispose();
                throw;
            }
        }

        internal int FieldCount => _fallback?.FieldCount ?? _visitor.FieldCount;

        int ICsvDataReaderHeaderRowSource.FieldCount => FieldCount;

        int? ICsvDataReaderPositionSource.CurrentPhysicalLineNumber =>
            _fallback is not null
                ? ((ICsvDataReaderPositionSource)_fallback).CurrentPhysicalLineNumber
                : _currentPhysicalLineNumber;

        int? ICsvDataReaderPositionSource.CurrentPhysicalEndLineNumber =>
            _fallback is not null
                ? ((ICsvDataReaderPositionSource)_fallback).CurrentPhysicalEndLineNumber
                : _currentPhysicalEndLineNumber;

        internal void SetSourceColumnCount(int sourceColumnCount)
        {
            if (sourceColumnCount < 0)
            {
                throw new ArgumentOutOfRangeException(nameof(sourceColumnCount));
            }

            _sourceColumnCount = sourceColumnCount;
            _visitor.SetSourceColumnCount(sourceColumnCount);
            _fallback?.SetSourceColumnCount(sourceColumnCount);
        }

        void ICsvDataReaderHeaderRowSource.SetSourceColumnCount(int sourceColumnCount) =>
            SetSourceColumnCount(sourceColumnCount);

        public bool Read() => Read(_options.CancellationToken);

        public bool Read(CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ThrowIfCancellationRequested(_options);
            return ReadCore(cancellationToken);
        }

        private bool ReadCore(CancellationToken cancellationToken)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            if (_fallback is not null)
            {
                return _fallback.Read(cancellationToken);
            }

            _currentPhysicalLineNumber = null;
            _currentPhysicalEndLineNumber = null;

            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                ThrowIfCancellationRequested(_options);
                if (!PrepareRecordStart(cancellationToken))
                {
                    return false;
                }

                int recordStart = _position;
                int fieldStart = recordStart;
                int fieldIndex = 0;
                int recordLineNumber = _physicalLineSeparatorsConsumed + 1;
                int cancellationCheckPosition = _position + 16 * 1024;
                _visitor.Reset();

                if (_buffer[recordStart] == _commentCharacter)
                {
                    return StartFallback(recordStart, recordLineNumber, cancellationToken);
                }

                while (true)
                {
                    if (_position >= cancellationCheckPosition)
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        ThrowIfCancellationRequested(_options);
                        cancellationCheckPosition = _position + 16 * 1024;
                    }

                    if (_position == _length)
                    {
                        if (!_endOfStream && !FillForCurrentRecord(
                                ref recordStart,
                                ref fieldStart,
                                cancellationToken) && !_endOfStream)
                        {
                            return StartFallback(recordStart, recordLineNumber, cancellationToken);
                        }

                        if (_endOfStream)
                        {
                            if (_position == recordStart && fieldIndex == 0)
                            {
                                return false;
                            }

                            if (CompleteRecord(
                                recordStart,
                                _position,
                                fieldStart,
                                fieldIndex,
                                recordLineNumber,
                                hasLineSeparator: false))
                            {
                                return true;
                            }

                            break;
                        }
                    }

                    if (Avx2.IsSupported && _length - _position >= Vector256<byte>.Count)
                    {
                        int chunkStart = _position;
                        Vector256<byte> values = Vector256.LoadUnsafe(
                            ref MemoryMarshal.GetArrayDataReference(_buffer),
                            (nuint)chunkStart);
                        uint delimiterMask = (uint)Avx2.MoveMask(Avx2.CompareEqual(values, _delimiterVector));
                        uint quoteMask = (uint)Avx2.MoveMask(Avx2.CompareEqual(values, QuoteVector));
                        uint carriageReturnMask = (uint)Avx2.MoveMask(Avx2.CompareEqual(values, CarriageReturnVector));
                        uint lineFeedMask = (uint)Avx2.MoveMask(Avx2.CompareEqual(values, LineFeedVector));
                        uint terminalMask = quoteMask | carriageReturnMask | lineFeedMask;
                        if (terminalMask == 0)
                        {
                            VisitDelimiterMask(delimiterMask, chunkStart, ref fieldStart, ref fieldIndex);
                            _position += Vector256<byte>.Count;
                            continue;
                        }

                        int terminalOffset = BitOperations.TrailingZeroCount(terminalMask);
                        uint delimitersBeforeTerminal = delimiterMask & ((1u << terminalOffset) - 1u);
                        VisitDelimiterMask(delimitersBeforeTerminal, chunkStart, ref fieldStart, ref fieldIndex);
                        int terminalIndex = chunkStart + terminalOffset;
                        byte terminal = _buffer[terminalIndex];
                        if (terminal == (byte)'"')
                        {
                            return StartFallback(recordStart, recordLineNumber, cancellationToken);
                        }

                        _position = terminalIndex + 1;
                        if (terminal == (byte)'\r')
                        {
                            _skipLineFeedAfterCarriageReturn = true;
                        }

                        _physicalLineSeparatorsConsumed++;
                        if (CompleteRecord(
                            recordStart,
                            terminalIndex,
                            fieldStart,
                            fieldIndex,
                            recordLineNumber,
                            hasLineSeparator: true))
                        {
                            return true;
                        }

                        break;
                    }

                    byte special = _buffer[_position];
                    if (special != _delimiter && special != (byte)'"' &&
                        special != (byte)'\r' && special != (byte)'\n')
                    {
                        _position++;
                        continue;
                    }

                    int specialIndex = _position;
                    if (special == _delimiter)
                    {
                        _visitor.VisitFieldRange(fieldIndex++, fieldStart, specialIndex - fieldStart);
                        fieldStart = specialIndex + 1;
                        _position = fieldStart;
                        continue;
                    }

                    if (special == (byte)'"')
                    {
                        return StartFallback(recordStart, recordLineNumber, cancellationToken);
                    }

                    _position = specialIndex + 1;
                    if (special == (byte)'\r')
                    {
                        _skipLineFeedAfterCarriageReturn = true;
                    }

                    _physicalLineSeparatorsConsumed++;
                    if (CompleteRecord(
                        recordStart,
                        specialIndex,
                        fieldStart,
                        fieldIndex,
                        recordLineNumber,
                        hasLineSeparator: true))
                    {
                        return true;
                    }

                    break;
                }
            }
        }

        public ReadOnlySpan<char> GetSpan(int ordinal) => _fallback is not null
            ? _fallback.GetSpan(ordinal)
            : _visitor.GetString(ordinal).AsSpan();

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        public string GetString(int ordinal) => _fallback?.GetString(ordinal) ?? _visitor.GetString(ordinal);

        public bool IsMissing(int ordinal) => _fallback?.IsMissing(ordinal) ?? _visitor.IsMissing(ordinal);

        public bool IsNull(int ordinal, string? nullValue) =>
            _fallback?.IsNull(ordinal, nullValue) ??
            (nullValue is not null && !_visitor.IsMissing(ordinal) &&
                string.Equals(_visitor.GetString(ordinal), nullValue, StringComparison.Ordinal));

        public int CopyStringValues(object[] values, int count, string? nullValue)
        {
            if (_fallback is not null)
            {
                return _fallback.CopyStringValues(values, count, nullValue);
            }

            int valueCount = Math.Min(count, _visitor.SourceColumnCount);
            for (int index = 0; index < valueCount; index++)
            {
                values[index] = nullValue is not null && IsNull(index, nullValue)
                    ? DBNull.Value
                    : _visitor.GetString(index);
            }

            for (int index = valueCount; index < count; index++)
            {
                values[index] = DBNull.Value;
            }

            return count;
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _disposed = true;
            _fallback?.Dispose();
            if (!_streamTransferred)
            {
                _stream.Dispose();
            }

            byte[] buffer = _buffer;
            _buffer = Array.Empty<byte>();
            ArrayPool<byte>.Shared.Return(buffer);
        }

        private bool PrepareRecordStart(CancellationToken cancellationToken)
        {
            if (_skipLineFeedAfterCarriageReturn)
            {
                _skipLineFeedAfterCarriageReturn = false;
                if (EnsureBuffered(cancellationToken) && _buffer[_position] == (byte)'\n')
                {
                    _position++;
                }
            }

            return EnsureBuffered(cancellationToken);
        }

        private bool EnsureBuffered(CancellationToken cancellationToken)
        {
            if (_position < _length)
            {
                return true;
            }

            if (_endOfStream)
            {
                return false;
            }

            cancellationToken.ThrowIfCancellationRequested();
            _options.CancellationToken.ThrowIfCancellationRequested();
            _position = 0;
            _length = _stream.Read(_buffer, 0, _buffer.Length);
            cancellationToken.ThrowIfCancellationRequested();
            _options.CancellationToken.ThrowIfCancellationRequested();
            _endOfStream = _length == 0;
            return !_endOfStream;
        }

        private bool FillForCurrentRecord(
            ref int recordStart,
            ref int fieldStart,
            CancellationToken cancellationToken)
        {
            int retainedLength = _length - recordStart;
            if (retainedLength == _buffer.Length)
            {
                return false;
            }

            if (retainedLength > 0 && recordStart != 0)
            {
                Buffer.BlockCopy(_buffer, recordStart, _buffer, 0, retainedLength);
                _visitor.ShiftStarts(-recordStart);
            }

            fieldStart -= recordStart;
            recordStart = 0;
            _position = retainedLength;
            _length = retainedLength;
            cancellationToken.ThrowIfCancellationRequested();
            _options.CancellationToken.ThrowIfCancellationRequested();
            int read = _stream.Read(_buffer, retainedLength, _buffer.Length - retainedLength);
            cancellationToken.ThrowIfCancellationRequested();
            _options.CancellationToken.ThrowIfCancellationRequested();
            if (read == 0)
            {
                _endOfStream = true;
                return false;
            }

            _length += read;
            return true;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void VisitDelimiterMask(
            uint delimiterMask,
            int chunkStart,
            ref int fieldStart,
            ref int fieldIndex)
        {
            while (delimiterMask != 0)
            {
                int delimiterIndex = chunkStart + BitOperations.TrailingZeroCount(delimiterMask);
                _visitor.VisitFieldRange(fieldIndex++, fieldStart, delimiterIndex - fieldStart);
                fieldStart = delimiterIndex + 1;
                delimiterMask &= delimiterMask - 1;
            }
        }

        private bool CompleteRecord(
            int recordStart,
            int recordEnd,
            int fieldStart,
            int fieldIndex,
            int recordLineNumber,
            bool hasLineSeparator)
        {
            if (!_allowEmpty && recordEnd == recordStart && fieldIndex == 0)
            {
                _visitor.Reset();
                return false;
            }

            _visitor.VisitFieldRange(fieldIndex, fieldStart, recordEnd - fieldStart);
            _visitor.Complete(
                fieldIndex + 1,
                _options.ColumnCountMismatchPolicy,
                Ascii.IsValid(_buffer.AsSpan(recordStart, recordEnd - recordStart)));
            _emittedRecordCount++;
            ReportProgress(_options, _emittedRecordCount, recordLineNumber);
            _currentPhysicalLineNumber = recordLineNumber;
            _currentPhysicalEndLineNumber = hasLineSeparator
                ? Math.Max(recordLineNumber, _physicalLineSeparatorsConsumed)
                : recordLineNumber;
            return true;
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private bool StartFallback(
            int recordStart,
            int recordLineNumber,
            CancellationToken cancellationToken)
        {
            int prefixLength = _length - recordStart;
            byte[] prefix = new byte[prefixLength];
            Buffer.BlockCopy(_buffer, recordStart, prefix, 0, prefixLength);
            var prefixedStream = new PrefixReadStream(prefix, _stream);
            try
            {
                var reader = new StreamReader(
                    prefixedStream,
                    _encoding,
                    detectEncodingFromByteOrderMarks: false,
                    bufferSize: FallbackTextBufferSize,
                    leaveOpen: false);
                _fallback = new CsvStreamDataReaderRowSource(
                    reader,
                    _options,
                    _emittedRecordCount,
                    recordLineNumber - 1);
                _streamTransferred = true;
            }
            catch
            {
                prefixedStream.Dispose();
                throw;
            }
            if (_sourceColumnCount > 0)
            {
                _fallback.SetSourceColumnCount(_sourceColumnCount);
            }

            return _fallback.Read(cancellationToken);
        }

        private static bool HasNonUtf8Preamble(ReadOnlySpan<byte> bytes) =>
            bytes.Length >= 2 &&
                ((bytes[0] == 0xFF && bytes[1] == 0xFE) ||
                 (bytes[0] == 0xFE && bytes[1] == 0xFF)) ||
            bytes.Length >= 4 &&
                ((bytes[0] == 0x00 && bytes[1] == 0x00 && bytes[2] == 0xFE && bytes[3] == 0xFF) ||
                 (bytes[0] == 0xFF && bytes[1] == 0xFE && bytes[2] == 0x00 && bytes[3] == 0x00));

        private sealed class PrefixReadStream : Stream
        {
            private readonly byte[] _prefix;
            private readonly Stream _remainder;
            private int _position;

            internal PrefixReadStream(byte[] prefix, Stream remainder)
            {
                _prefix = prefix;
                _remainder = remainder;
            }

            public override bool CanRead => true;
            public override bool CanSeek => false;
            public override bool CanWrite => false;
            public override long Length => throw new NotSupportedException();
            public override long Position
            {
                get => throw new NotSupportedException();
                set => throw new NotSupportedException();
            }

            public override int Read(byte[] buffer, int offset, int count) =>
                Read(buffer.AsSpan(offset, count));

            public override int Read(Span<byte> buffer)
            {
                int prefixRemaining = _prefix.Length - _position;
                if (prefixRemaining > 0)
                {
                    int copied = Math.Min(prefixRemaining, buffer.Length);
                    _prefix.AsSpan(_position, copied).CopyTo(buffer);
                    _position += copied;
                    return copied;
                }

                return _remainder.Read(buffer);
            }

            public override void Flush() { }
            public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

            protected override void Dispose(bool disposing)
            {
                if (disposing)
                {
                    _remainder.Dispose();
                }

                base.Dispose(disposing);
            }
        }
    }

}
#endif
