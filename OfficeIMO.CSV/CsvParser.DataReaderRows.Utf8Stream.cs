#if NET8_0_OR_GREATER
using System.Buffers;
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
using System.Text;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    /// <summary>
    /// Pull-based UTF-8 row source. RFC 4180 and lenient quoted records are parsed from the
    /// bytes with the canonical text parser's semantics; a comment record hands the rest of the
    /// input to the canonical text row source from that record boundary.
    /// </summary>
    internal sealed partial class CsvUtf8StreamDataReaderRowSource : ICsvDataReaderHeaderRowSource, ICsvDataReaderPositionSource
    {
        private const int BufferSize = 256 * 1024;
        private const int FallbackTextBufferSize = 128 * 1024;

        private readonly CsvParsingSettings _options;
        private readonly Encoding _encoding;
        private CsvStringCache? _stringCache;
        private readonly byte _delimiter;
        private readonly byte _commentCharacter;
        private readonly SearchValues<byte> _specialBytes;
        private const int CancellationCheckInterval = 16 * 1024;
        private const int MaximumBufferSize = 1 << 30;
        private const int BlockSize = 64;
        private readonly bool _allowEmpty;
        private readonly bool _strictQuotes;
        private readonly bool _trim;
        private Stream _stream;
        private byte[] _buffer;
        private byte[]? _flexibleFieldBytes;
        private RecordScan _scan;
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
            CsvParsingSettings options,
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
            _specialBytes = SearchValues.Create(new[] { _delimiter, (byte)'"', (byte)'\r', (byte)'\n' });
            _allowEmpty = options.AllowEmptyLines;
            _strictQuotes = options.QuoteParsingMode == CsvQuoteParsingMode.Strict;
            _trim = options.TrimWhitespace;
            _buffer = buffer;
            _length = length;
            _position = position;
            _visitor = new CsvUtf8DataReaderRowVisitor(buffer, encoding);
            _stringCache = RentStringCacheWorkspace(options);
        }

        internal static bool TryCreate(
            Stream stream,
            CsvParsingSettings options,
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

                Encoding encoding = options.Encoding ?? new UTF8Encoding(false);
                bool hasUtf8Preamble = length >= 3 &&
                    buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF
                    && (options.DetectEncodingFromByteOrderMarks
                        || encoding.GetPreamble().AsSpan().SequenceEqual(buffer.AsSpan(0, 3)));
                int position = hasUtf8Preamble ? 3 : 0;
                rows = new CsvUtf8StreamDataReaderRowSource(
                    stream,
                    options,
                    encoding,
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
            ArgumentOutOfRangeException.ThrowIfNegative(sourceColumnCount);

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
            bool checkCancellation = false;
            while (true)
            {
                if (checkCancellation)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    ThrowIfCancellationRequested(_options);
                }

                checkCancellation = true;
                if (!PrepareRecordStart(cancellationToken))
                {
                    return false;
                }

                int recordStart = _position;
                int recordLineNumber = _physicalLineSeparatorsConsumed + 1;
                if (_buffer[recordStart] == _commentCharacter)
                {
                    return StartFallback(recordStart, recordLineNumber, cancellationToken);
                }

                RecordScanResult result;
                while (true)
                {
                    result = ScanRecord(recordStart, cancellationToken, out _scan);
                    if (result == RecordScanResult.Invalid && !_strictQuotes)
                    {
                        result = ScanFlexibleRecord(recordStart, out _scan);
                    }

                    if (result != RecordScanResult.NeedMoreData)
                    {
                        break;
                    }

                    if (!FillForCurrentRecord(ref recordStart, cancellationToken) && !_endOfStream)
                    {
                        return StartFallback(recordStart, recordLineNumber, cancellationToken);
                    }
                }

                if (result == RecordScanResult.Complete)
                {
                    if (CompleteRecord(recordStart, recordLineNumber))
                    {
                        return true;
                    }

                    continue;
                }

                if (result == RecordScanResult.Invalid)
                {
                    int errorLineNumber = recordLineNumber + _scan.LineBreaks;
                    HandleParseError(
                        _options,
                        new CsvParseException("Invalid quoted field.", errorLineNumber),
                        errorLineNumber);
                    SkipInvalidRecordLine(cancellationToken);
                    continue;
                }

                // An unterminated quote consumes the remaining input, as the text reader's line
                // continuation does. Its line is the last physical line actually read.
                int lastLineNumber = recordLineNumber + _scan.LineBreaks - (_scan.EndsWithLineBreak ? 1 : 0);
                _position = _length;
                _physicalLineSeparatorsConsumed += _scan.LineBreaks;
                HandleParseError(
                    _options,
                    new CsvParseException("Unterminated quoted field.", lastLineNumber),
                    lastLineNumber);
            }
        }

        private enum RecordScanResult
        {
            Complete,
            NeedMoreData,
            Invalid,
            Unterminated,
            /// <summary>The block scanner cannot classify the record; the exact scanner rescans it.</summary>
            Irregular
        }

        /// <summary>Boundaries of the record most recently scanned from the current buffer.</summary>
        private struct RecordScan
        {
            /// <summary>Index of the record terminator, or the end of input.</summary>
            internal int End;
            /// <summary>Index after the terminator; for an invalid record, the invalid byte.</summary>
            internal int Next;
            internal int FieldCount;
            /// <summary>Line separators inside quoted fields before <see cref="Next"/>.</summary>
            internal int LineBreaks;
            internal byte Terminator;
            internal bool EndsWithLineBreak;
            /// <summary>Whether the block scanner already established if the record is ASCII.</summary>
            internal AsciiState Ascii;
            /// <summary>The block scanner found no escaped quote pair in the record.</summary>
            internal bool NoEscapedQuotes;
        }

        private enum AsciiState : byte
        {
            Unknown,
            Ascii,
            NonAscii
        }

        /// <summary>
        /// Scans one record from <paramref name="recordStart"/>. Quoted fields are stored as their
        /// inner byte range and unescaped only when requested. The scan restarts from the record
        /// start after a refill, so a partially buffered record never leaves partial state.
        /// </summary>
        private RecordScanResult ScanRecord(int recordStart, CancellationToken cancellationToken, out RecordScan scan)
        {
            if (Vector128.IsHardwareAccelerated && BitConverter.IsLittleEndian)
            {
                RecordScanResult result = ScanRecordBlocks(recordStart, cancellationToken, out scan);
                if (result != RecordScanResult.Irregular)
                {
                    return result;
                }
            }

            return ScanRecordExact(recordStart, cancellationToken, out scan);
        }

        /// <summary>
        /// Scans 64-byte blocks of a well-formed record. Quote parity is a prefix XOR of the quote
        /// mask, so delimiters and line separators inside quotes drop out of the structural mask and
        /// only field boundaries are visited. A quoted field is accepted when it starts and ends
        /// with a quote and every inner quote is part of an escaped pair; anything else, and the
        /// final partial block, is <see cref="RecordScanResult.Irregular"/> for the exact scanner.
        /// </summary>
        private RecordScanResult ScanRecordBlocks(int recordStart, CancellationToken cancellationToken, out RecordScan scan)
        {
            if (_trim)
            {
                return ScanTrimmedRecordBlocks(recordStart, cancellationToken, out scan);
            }

            scan = default;
            _visitor.Reset();
            int length = _length;
            int position = recordStart;
            int fieldStart = recordStart;
            int fieldIndex = 0;
            int lineBreaks = 0;
            ulong insideQuotes = 0;
            ulong nonAscii = 0;
            bool closingQuotePending = false;
            bool openingQuoteAllowed = true;
            bool escapedQuotes = false;
            bool carriageReturnInside = false;
            int cancellationCheckPosition = position + CancellationCheckInterval;
            ref byte data = ref MemoryMarshal.GetArrayDataReference(_buffer);
            while (true)
            {
                if (length - position < BlockSize)
                {
                    return _endOfStream ? RecordScanResult.Irregular : RecordScanResult.NeedMoreData;
                }

                LoadBlockMasks(ref data, position, out ulong delimiters, out ulong quotes,
                    out ulong carriageReturns, out ulong lineFeeds, out ulong high);
                ulong lineSeparators = carriageReturns | lineFeeds;
                ulong structural, recordEnds, boundaries;
                if ((quotes | insideQuotes) == 0 && !closingQuotePending)
                {
                    structural = delimiters | lineSeparators;
                    recordEnds = lineSeparators;
                    boundaries = structural;
                    carriageReturnInside = false;
                }
                else
                {
                    // Quote parity marks bytes inside quotes. A closing quote must be followed by a
                    // quote (an escaped pair), a delimiter or a line separator; an opening quote must
                    // start a field or follow a closing quote. Anything else is left to the exact scan.
                    ulong inside = PrefixXor(quotes) ^ insideQuotes;
                    ulong opening = quotes & inside;
                    ulong closing = quotes & ~inside;
                    structural = (delimiters | lineSeparators) & ~inside;
                    recordEnds = lineSeparators & ~inside;
                    boundaries = structural | closing;
                    ulong scope = recordEnds == 0 ? ulong.MaxValue : recordEnds ^ (recordEnds - 1UL);
                    ulong followsClosing = quotes | delimiters | lineSeparators;
                    ulong invalid = ((closing << 1) & ~followsClosing) |
                        (opening & ~((boundaries << 1) | (openingQuoteAllowed ? 1UL : 0UL)));
                    if ((invalid & scope) != 0 || (closingQuotePending && (followsClosing & 1UL) == 0))
                    {
                        return RecordScanResult.Irregular;
                    }

                    // The second quote of an escaped pair directly follows a closing quote.
                    escapedQuotes |= ((((closing << 1) | (closingQuotePending ? 1UL : 0UL)) & quotes & scope) != 0);

                    ulong innerCarriageReturns = carriageReturns & inside & scope;
                    ulong innerLineFeeds = lineFeeds & inside & scope;
                    if ((innerCarriageReturns | innerLineFeeds) != 0)
                    {
                        // CR, LF and CRLF inside quotes each end one physical line.
                        lineBreaks += BitOperations.PopCount(innerCarriageReturns) + BitOperations.PopCount(
                            innerLineFeeds & ~((innerCarriageReturns << 1) | (carriageReturnInside ? 1UL : 0UL)));
                    }

                    carriageReturnInside = (innerCarriageReturns >> 63) != 0;
                    closingQuotePending = (closing >> 63) != 0;
                    insideQuotes = (ulong)((long)inside >> 63);
                }

                if (recordEnds != 0)
                {
                    int end = BitOperations.TrailingZeroCount(recordEnds);
                    ulong before = (1UL << end) - 1UL;
                    _visitor.VisitDelimiterMask(structural & before, position, ref fieldStart, ref fieldIndex);
                    RecordScanResult result = CompleteBlockRecord(position + end, fieldStart, fieldIndex, lineBreaks, (nonAscii | (high & before)) == 0, out scan);
                    scan.NoEscapedQuotes = !escapedQuotes;
                    return result;
                }

                // Quoted fields keep their quotes; the visitor removes them when the text is read.
                _visitor.VisitDelimiterMask(structural, position, ref fieldStart, ref fieldIndex);
                openingQuoteAllowed = (boundaries >> 63) != 0;
                nonAscii |= high;
                position += BlockSize;
                if (position >= cancellationCheckPosition)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    ThrowIfCancellationRequested(_options);
                    cancellationCheckPosition = position + CancellationCheckInterval;
                }
            }
        }

        /// <summary>
        /// Block scan with trimming. Each field is visited on its own so whitespace around quoted
        /// text can be removed before the quotes are checked.
        /// </summary>
        private RecordScanResult ScanTrimmedRecordBlocks(int recordStart, CancellationToken cancellationToken, out RecordScan scan)
        {
            scan = default;
            _visitor.Reset();
            byte[] buffer = _buffer;
            int length = _length;
            int position = recordStart;
            int fieldStart = recordStart;
            int fieldIndex = 0;
            int fieldQuotes = 0;
            int lineBreaks = 0;
            bool innerLineBreaks = false;
            ulong insideQuotes = 0;
            ulong nonAscii = 0;
            int cancellationCheckPosition = position + CancellationCheckInterval;
            ref byte data = ref MemoryMarshal.GetArrayDataReference(buffer);
            while (true)
            {
                if (length - position < BlockSize)
                {
                    return _endOfStream ? RecordScanResult.Irregular : RecordScanResult.NeedMoreData;
                }

                LoadBlockMasks(ref data, position, out ulong delimiters, out ulong quotes,
                    out ulong carriageReturns, out ulong lineFeeds, out ulong high);
                ulong lineSeparators = carriageReturns | lineFeeds;
                ulong inside = insideQuotes;
                if (quotes != 0)
                {
                    inside = PrefixXor(quotes) ^ insideQuotes;
                    insideQuotes = (ulong)((long)inside >> 63);
                }

                innerLineBreaks |= (lineSeparators & inside) != 0;
                ulong structural = (delimiters | lineSeparators) & ~inside;
                while (structural != 0)
                {
                    int bit = BitOperations.TrailingZeroCount(structural);
                    ulong before = (1UL << bit) - 1UL;
                    fieldQuotes += BitOperations.PopCount(quotes & before);
                    quotes &= ~before;
                    int end = position + bit;
                    if (fieldQuotes == 0)
                    {
                        VisitUnquotedField(fieldIndex, fieldStart, end);
                    }
                    else if (!TryVisitQuotedField(fieldIndex, fieldStart, end, fieldQuotes, innerLineBreaks, ref lineBreaks))
                    {
                        return RecordScanResult.Irregular;
                    }

                    fieldQuotes = 0;
                    if ((lineSeparators & (1UL << bit)) != 0)
                    {
                        return CompleteBlockRecord(end, fieldStart: -1, fieldIndex, lineBreaks, (nonAscii | (high & before)) == 0, out scan);
                    }

                    fieldIndex++;
                    fieldStart = end + 1;
                    structural &= structural - 1UL;
                }

                fieldQuotes += BitOperations.PopCount(quotes);
                nonAscii |= high;
                position += BlockSize;

                if (position >= cancellationCheckPosition)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    ThrowIfCancellationRequested(_options);
                    cancellationCheckPosition = position + CancellationCheckInterval;
                }
            }
        }

        /// <summary>Completes a block-scanned record at a line separator; a negative field start means its last field was visited.</summary>
        private RecordScanResult CompleteBlockRecord(int end, int fieldStart, int fieldIndex, int lineBreaks, bool isAscii, out RecordScan scan)
        {
            if (fieldStart >= 0)
            {
                VisitUnquotedField(fieldIndex, fieldStart, end);
            }

            scan = default;
            scan.End = end;
            scan.Next = end + 1;
            scan.FieldCount = fieldIndex + 1;
            scan.LineBreaks = lineBreaks;
            scan.Terminator = _buffer[end];
            scan.Ascii = isAscii ? AsciiState.Ascii : AsciiState.NonAscii;
            return RecordScanResult.Complete;
        }

        private bool TryVisitQuotedField(int fieldIndex, int fieldStart, int end, int quotes, bool innerLineBreaks, ref int lineBreaks)
        {
            byte[] buffer = _buffer;
            if (_trim)
            {
                // Strict quoting admits whitespace only after the closing quote.
                while (!_strictQuotes && fieldStart < end && IsAsciiWhiteSpace(buffer[fieldStart])) fieldStart++;
                while (end > fieldStart && IsAsciiWhiteSpace(buffer[end - 1])) end--;
            }

            if ((quotes & 1) != 0 || end - fieldStart < 2 || buffer[fieldStart] != (byte)'"' || buffer[end - 1] != (byte)'"')
            {
                return false;
            }

            int contentStart = fieldStart + 1;
            int contentLength = end - fieldStart - 2;
            bool escaped = quotes > 2;
            if (escaped)
            {
                ReadOnlySpan<byte> content = buffer.AsSpan(contentStart, contentLength);
                while (true)
                {
                    int quote = content.IndexOf((byte)'"');
                    if (quote < 0)
                    {
                        break;
                    }

                    if (quote + 1 == content.Length || content[quote + 1] != (byte)'"')
                    {
                        return false;
                    }

                    content = content.Slice(quote + 2);
                }
            }

            if (innerLineBreaks)
            {
                lineBreaks += CountLineBreaks(buffer, contentStart, end - 1, out _);
            }

            _visitor.VisitQuotedFieldRange(fieldIndex, contentStart, contentLength, escaped);
            return true;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void LoadBlockMasks(ref byte data, int position, out ulong delimiters, out ulong quotes,
            out ulong carriageReturns, out ulong lineFeeds, out ulong high)
        {
            if (Vector512.IsHardwareAccelerated)
            {
                Vector512<byte> values = Vector512.LoadUnsafe(ref data, (nuint)position);
                high = values.ExtractMostSignificantBits();
                delimiters = Vector512.Equals(values, Vector512.Create(_delimiter)).ExtractMostSignificantBits();
                quotes = Vector512.Equals(values, Vector512.Create((byte)'"')).ExtractMostSignificantBits();
                carriageReturns = Vector512.Equals(values, Vector512.Create((byte)'\r')).ExtractMostSignificantBits();
                lineFeeds = Vector512.Equals(values, Vector512.Create((byte)'\n')).ExtractMostSignificantBits();
            }
            else if (Vector256.IsHardwareAccelerated)
            {
                LoadMasks256(ref data, position, out uint lowDelimiters, out uint lowQuotes, out uint lowReturns, out uint lowFeeds, out uint lowHigh);
                LoadMasks256(ref data, position + 32, out uint highDelimiters, out uint highQuotes, out uint highReturns, out uint highFeeds, out uint highHigh);
                high = lowHigh | ((ulong)highHigh << 32);
                delimiters = lowDelimiters | ((ulong)highDelimiters << 32);
                quotes = lowQuotes | ((ulong)highQuotes << 32);
                carriageReturns = lowReturns | ((ulong)highReturns << 32);
                lineFeeds = lowFeeds | ((ulong)highFeeds << 32);
            }
            else
            {
                delimiters = quotes = carriageReturns = lineFeeds = high = 0;
                for (int offset = 0; offset < BlockSize; offset += 16)
                {
                    Vector128<byte> values = Vector128.LoadUnsafe(ref data, (nuint)(position + offset));
                    high |= (ulong)values.ExtractMostSignificantBits() << offset;
                    delimiters |= (ulong)Vector128.Equals(values, Vector128.Create(_delimiter)).ExtractMostSignificantBits() << offset;
                    quotes |= (ulong)Vector128.Equals(values, Vector128.Create((byte)'"')).ExtractMostSignificantBits() << offset;
                    carriageReturns |= (ulong)Vector128.Equals(values, Vector128.Create((byte)'\r')).ExtractMostSignificantBits() << offset;
                    lineFeeds |= (ulong)Vector128.Equals(values, Vector128.Create((byte)'\n')).ExtractMostSignificantBits() << offset;
                }
            }
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void LoadMasks256(ref byte data, int position, out uint delimiters, out uint quotes,
            out uint carriageReturns, out uint lineFeeds, out uint high)
        {
            Vector256<byte> values = Vector256.LoadUnsafe(ref data, (nuint)position);
            high = values.ExtractMostSignificantBits();
            delimiters = Vector256.Equals(values, Vector256.Create(_delimiter)).ExtractMostSignificantBits();
            quotes = Vector256.Equals(values, Vector256.Create((byte)'"')).ExtractMostSignificantBits();
            carriageReturns = Vector256.Equals(values, Vector256.Create((byte)'\r')).ExtractMostSignificantBits();
            lineFeeds = Vector256.Equals(values, Vector256.Create((byte)'\n')).ExtractMostSignificantBits();
        }


        /// <summary>Sets each bit to the parity of the quote bits at and below it.</summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static ulong PrefixXor(ulong mask)
        {
            if (Pclmulqdq.IsSupported)
            {
                return Pclmulqdq.CarrylessMultiply(
                    Vector128.CreateScalar(mask), Vector128.CreateScalar(ulong.MaxValue), 0).ToScalar();
            }

            mask ^= mask << 1;
            mask ^= mask << 2;
            mask ^= mask << 4;
            mask ^= mask << 8;
            mask ^= mask << 16;
            mask ^= mask << 32;
            return mask;
        }

        /// <summary>
        /// Exact byte-by-byte RFC 4180 record scan used for irregular records and input tails. It
        /// reports the precise position of invalid quoting and unterminated quoted fields.
        /// </summary>
        private RecordScanResult ScanRecordExact(int recordStart, CancellationToken cancellationToken, out RecordScan scan)
        {
            scan = default;
            _visitor.Reset();
            byte[] buffer = _buffer;
            int length = _length;
            int position = recordStart;
            int fieldStart = recordStart;
            int fieldIndex = 0;
            int lineBreaks = 0;
            int cancellationCheckPosition = position + CancellationCheckInterval;
            while (true)
            {
                if (position >= cancellationCheckPosition)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    ThrowIfCancellationRequested(_options);
                    cancellationCheckPosition = position + CancellationCheckInterval;
                }

                int found = buffer.AsSpan(position, length - position).IndexOfAny(_specialBytes);
                if (found < 0)
                {
                    if (!_endOfStream)
                    {
                        return RecordScanResult.NeedMoreData;
                    }

                    VisitUnquotedField(fieldIndex, fieldStart, length);
                    scan.End = length;
                    scan.Next = length;
                    scan.FieldCount = fieldIndex + 1;
                    scan.LineBreaks = lineBreaks;
                    return RecordScanResult.Complete;
                }

                int special = position + found;
                if (buffer[special] == _delimiter)
                {
                    VisitUnquotedField(fieldIndex++, fieldStart, special);
                    fieldStart = position = special + 1;
                    continue;
                }

                byte terminal = buffer[special];
                if (terminal != (byte)'"')
                {
                    VisitUnquotedField(fieldIndex, fieldStart, special);
                    scan.End = special;
                    scan.Next = special + 1;
                    scan.FieldCount = fieldIndex + 1;
                    scan.LineBreaks = lineBreaks;
                    scan.Terminator = terminal;
                    return RecordScanResult.Complete;
                }

                if (special != fieldStart)
                {
                    // A quote inside an unquoted field is not RFC 4180 syntax.
                    scan.Next = special;
                    scan.LineBreaks = lineBreaks;
                    return RecordScanResult.Invalid;
                }

                int contentStart = special + 1;
                int search = contentStart;
                bool escaped = false;
                int closingQuote;
                while (true)
                {
                    int quote = buffer.AsSpan(search, length - search).IndexOf((byte)'"');
                    if (quote < 0)
                    {
                        if (!_endOfStream)
                        {
                            return RecordScanResult.NeedMoreData;
                        }

                        scan.LineBreaks = lineBreaks + CountLineBreaks(buffer, contentStart, length, out bool endsWithLineBreak);
                        scan.EndsWithLineBreak = endsWithLineBreak;
                        return RecordScanResult.Unterminated;
                    }

                    closingQuote = search + quote;
                    if (closingQuote + 1 == length && !_endOfStream)
                    {
                        return RecordScanResult.NeedMoreData;
                    }

                    if (closingQuote + 1 < length && buffer[closingQuote + 1] == (byte)'"')
                    {
                        escaped = true;
                        search = closingQuote + 2;
                        continue;
                    }

                    break;
                }

                lineBreaks += CountLineBreaks(buffer, contentStart, closingQuote, out _);
                _visitor.VisitQuotedFieldRange(fieldIndex, contentStart, closingQuote - contentStart, escaped);
                position = closingQuote + 1;
                if (_trim && !SkipWhiteSpaceAfterQuote(ref position))
                {
                    return RecordScanResult.NeedMoreData;
                }

                if (position == length)
                {
                    scan.End = position;
                    scan.Next = position;
                    scan.FieldCount = fieldIndex + 1;
                    scan.LineBreaks = lineBreaks;
                    return RecordScanResult.Complete;
                }

                byte next = buffer[position];
                if (next == _delimiter)
                {
                    fieldIndex++;
                    fieldStart = ++position;
                    continue;
                }

                if (next == (byte)'\r' || next == (byte)'\n')
                {
                    scan.End = position;
                    scan.Next = position + 1;
                    scan.FieldCount = fieldIndex + 1;
                    scan.LineBreaks = lineBreaks;
                    scan.Terminator = next;
                    return RecordScanResult.Complete;
                }

                scan.Next = position;
                scan.LineBreaks = lineBreaks;
                return RecordScanResult.Invalid;
            }
        }

        /// <summary>
        /// Lenient parse of a record that is not RFC 4180 quoted, matching the canonical flexible
        /// parser: a quote opens a quoted section anywhere, and a quote right after a closing quote
        /// is literal. With trimming, whitespace before an opening quote and after a closing quote
        /// is dropped and fields without quotes are trimmed. Each field is decoded during the scan.
        /// </summary>
        private RecordScanResult ScanFlexibleRecord(int recordStart, out RecordScan scan)
        {
            scan = default;
            _visitor.Reset();
            _flexibleFieldBytes ??= new byte[256];
            byte[] buffer = _buffer;
            int length = _length;
            int position = recordStart;
            int fieldIndex = 0;
            int fieldLength = 0;
            int lineBreaks = 0;
            bool inQuotes = false;
            bool afterClosingQuote = false;
            bool fieldWasQuoted = false;
            bool whiteSpaceOnly = true;
            while (true)
            {
                if (position == length)
                {
                    if (!_endOfStream)
                    {
                        return RecordScanResult.NeedMoreData;
                    }

                    if (inQuotes)
                    {
                        scan.LineBreaks = lineBreaks;
                        scan.EndsWithLineBreak = position > recordStart &&
                            buffer[position - 1] is (byte)'\r' or (byte)'\n';
                        return RecordScanResult.Unterminated;
                    }

                    VisitFlexibleField(fieldIndex, fieldLength, fieldWasQuoted);
                    scan.End = position;
                    scan.Next = position;
                    scan.FieldCount = fieldIndex + 1;
                    scan.LineBreaks = lineBreaks;
                    return RecordScanResult.Complete;
                }

                byte value = buffer[position];
                if (value == (byte)'"')
                {
                    if (inQuotes)
                    {
                        if (position + 1 == length && !_endOfStream)
                        {
                            return RecordScanResult.NeedMoreData;
                        }

                        if (position + 1 < length && buffer[position + 1] == (byte)'"')
                        {
                            AppendFlexibleBytes(ref fieldLength, position, 1);
                            whiteSpaceOnly = false;
                            position += 2;
                            continue;
                        }

                        inQuotes = false;
                        afterClosingQuote = true;
                    }
                    else if (afterClosingQuote)
                    {
                        AppendFlexibleBytes(ref fieldLength, position, 1);
                        whiteSpaceOnly = false;
                        afterClosingQuote = false;
                    }
                    else
                    {
                        if (_trim && whiteSpaceOnly)
                        {
                            fieldLength = 0;
                        }

                        inQuotes = true;
                        fieldWasQuoted = true;
                    }

                    position++;
                    continue;
                }

                if (value == (byte)'\r' || value == (byte)'\n')
                {
                    if (!inQuotes)
                    {
                        VisitFlexibleField(fieldIndex, fieldLength, fieldWasQuoted);
                        scan.End = position;
                        scan.Next = position + 1;
                        scan.FieldCount = fieldIndex + 1;
                        scan.LineBreaks = lineBreaks;
                        scan.Terminator = value;
                        return RecordScanResult.Complete;
                    }

                    // A separator inside quotes is kept as read and does not change whether the
                    // field so far is whitespace.
                    int separatorLength = 1;
                    if (value == (byte)'\r')
                    {
                        if (position + 1 == length && !_endOfStream)
                        {
                            return RecordScanResult.NeedMoreData;
                        }

                        if (position + 1 < length && buffer[position + 1] == (byte)'\n')
                        {
                            separatorLength = 2;
                        }
                    }

                    lineBreaks++;
                    AppendFlexibleBytes(ref fieldLength, position, separatorLength);
                    position += separatorLength;
                    continue;
                }

                if (value == _delimiter && !inQuotes)
                {
                    VisitFlexibleField(fieldIndex++, fieldLength, fieldWasQuoted);
                    fieldLength = 0;
                    afterClosingQuote = false;
                    fieldWasQuoted = false;
                    whiteSpaceOnly = true;
                    position++;
                    continue;
                }

                int characterLength = 1;
                bool whiteSpace = false;
                if (_trim)
                {
                    if (!TryClassifyWhiteSpace(position, out whiteSpace, out characterLength))
                    {
                        return RecordScanResult.NeedMoreData;
                    }

                    if (!inQuotes && afterClosingQuote && whiteSpace)
                    {
                        position += characterLength;
                        continue;
                    }
                }

                if (!inQuotes)
                {
                    afterClosingQuote = false;
                }

                AppendFlexibleBytes(ref fieldLength, position, characterLength);
                whiteSpaceOnly &= whiteSpace;
                position += characterLength;
            }
        }

        private void AppendFlexibleBytes(ref int fieldLength, int start, int count)
        {
            if (fieldLength + count > _flexibleFieldBytes!.Length)
            {
                Array.Resize(ref _flexibleFieldBytes, checked(Math.Max(fieldLength + count, _flexibleFieldBytes.Length * 2)));
            }

            _buffer.AsSpan(start, count).CopyTo(_flexibleFieldBytes.AsSpan(fieldLength));
            fieldLength += count;
        }

        private void VisitFlexibleField(int fieldIndex, int fieldLength, bool fieldWasQuoted)
        {
            ReadOnlySpan<byte> bytes = _flexibleFieldBytes.AsSpan(0, fieldLength);
            string value = fieldLength == 0
                ? string.Empty
                : Ascii.IsValid(bytes) ? Encoding.Latin1.GetString(bytes) : _encoding.GetString(bytes);
            _visitor.VisitMaterializedField(fieldIndex, _trim && !fieldWasQuoted ? value.Trim() : value);
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static bool IsAsciiWhiteSpace(byte value) => value == (byte)' ' || (uint)(value - 9) <= 4;

        /// <summary>
        /// Classifies the character at <paramref name="position"/> as <see cref="char.IsWhiteSpace(char)"/>
        /// does, decoding a non-ASCII scalar. Returns false when a sequence continues past buffered input.
        /// </summary>
        private bool TryClassifyWhiteSpace(int position, out bool whiteSpace, out int length)
        {
            byte value = _buffer[position];
            if (value < 0x80)
            {
                whiteSpace = IsAsciiWhiteSpace(value);
                length = 1;
                return true;
            }

            OperationStatus status = Rune.DecodeFromUtf8(_buffer.AsSpan(position, _length - position), out Rune rune, out length);
            if (status == OperationStatus.NeedMoreData && !_endOfStream)
            {
                whiteSpace = false;
                return false;
            }

            // Invalid UTF-8 decodes to U+FFFD, which is not whitespace.
            whiteSpace = status == OperationStatus.Done && Rune.IsWhiteSpace(rune);
            return true;
        }

        /// <summary>Skips whitespace between a closing quote and the next delimiter or line separator.</summary>
        private bool SkipWhiteSpaceAfterQuote(ref int position)
        {
            while (position < _length)
            {
                byte value = _buffer[position];
                if (value == _delimiter || value == (byte)'\r' || value == (byte)'\n')
                {
                    return true;
                }

                if (!TryClassifyWhiteSpace(position, out bool whiteSpace, out int length))
                {
                    return false;
                }

                if (!whiteSpace)
                {
                    return true;
                }

                position += length;
            }

            return _endOfStream;
        }

        /// <summary>Stores an unquoted field, trimming it as the canonical parser does when requested.</summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void VisitUnquotedField(int fieldIndex, int start, int end)
        {
            if (_trim && end > start &&
                (IsAsciiWhiteSpace(_buffer[start]) || IsAsciiWhiteSpace(_buffer[end - 1]) ||
                 _buffer[start] >= 0x80 || _buffer[end - 1] >= 0x80))
            {
                VisitTrimmedField(fieldIndex, start, end);
                return;
            }

            _visitor.VisitFieldRange(fieldIndex, start, end - start);
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private void VisitTrimmedField(int fieldIndex, int start, int end)
        {
            byte[] buffer = _buffer;
            while (start < end && IsAsciiWhiteSpace(buffer[start])) start++;
            while (end > start && IsAsciiWhiteSpace(buffer[end - 1])) end--;
            ReadOnlySpan<byte> text = buffer.AsSpan(start, end - start);
            if (!text.IsEmpty &&
                ((text[0] >= 0x80 && Rune.DecodeFromUtf8(text, out Rune first, out _) == OperationStatus.Done && Rune.IsWhiteSpace(first)) ||
                 (text[^1] >= 0x80 && Rune.DecodeLastFromUtf8(text, out Rune last, out _) == OperationStatus.Done && Rune.IsWhiteSpace(last))))
            {
                // Non-ASCII whitespace at an edge: trim the decoded text exactly as string.Trim does.
                _visitor.VisitMaterializedField(fieldIndex, _encoding.GetString(text).Trim());
                return;
            }

            _visitor.VisitFieldRange(fieldIndex, start, end - start);
        }


        /// <summary>Counts CR, LF and CRLF separators in quoted content.</summary>
        private static int CountLineBreaks(byte[] buffer, int start, int end, out bool endsWithLineBreak)
        {
            endsWithLineBreak = end > start && buffer[end - 1] is (byte)'\r' or (byte)'\n';
            ReadOnlySpan<byte> content = buffer.AsSpan(start, end - start);
            int index = content.IndexOfAny((byte)'\r', (byte)'\n');
            if (index < 0)
            {
                return 0;
            }

            int count = 0;
            for (; index < content.Length; index++)
            {
                byte value = content[index];
                if (value == (byte)'\n')
                {
                    count++;
                }
                else if (value == (byte)'\r')
                {
                    count++;
                    if (index + 1 < content.Length && content[index + 1] == (byte)'\n')
                    {
                        index++;
                    }
                }
            }

            return count;
        }

        /// <summary>Skips an invalid record through the end of the physical line that contains the error.</summary>
        private void SkipInvalidRecordLine(CancellationToken cancellationToken)
        {
            _physicalLineSeparatorsConsumed += _scan.LineBreaks;
            int position = _scan.Next;
            while (true)
            {
                int found = _buffer.AsSpan(position, _length - position).IndexOfAny((byte)'\r', (byte)'\n');
                if (found >= 0)
                {
                    position += found;
                    if (_buffer[position] == (byte)'\r')
                    {
                        _skipLineFeedAfterCarriageReturn = true;
                    }

                    _position = position + 1;
                    _physicalLineSeparatorsConsumed++;
                    return;
                }

                if (_endOfStream)
                {
                    _position = _length;
                    return;
                }

                // Nothing before the end of the buffer is needed again.
                _position = _length;
                if (!EnsureBuffered(cancellationToken))
                {
                    return;
                }

                position = _position;
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
            (nullValue is not null && !_visitor.IsMissing(ordinal) && _visitor.TextEquals(ordinal, nullValue));

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
            CsvStringCache? cache = Interlocked.Exchange(ref _stringCache, null);
            try
            {
                _fallback?.Dispose();
                if (!_streamTransferred) _stream.Dispose();
            }
            finally
            {
                byte[] buffer = _buffer;
                _buffer = Array.Empty<byte>();
                try { ArrayPool<byte>.Shared.Return(buffer); }
                finally { ReturnStringCacheWorkspace(cache); }
            }
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

        /// <summary>
        /// Moves the current record to the start of the buffer, growing the buffer when the record
        /// already fills it, and appends more input. The caller rescans the record afterwards.
        /// </summary>
        private bool FillForCurrentRecord(ref int recordStart, CancellationToken cancellationToken)
        {
            int retainedLength = _length - recordStart;
            if (retainedLength == _buffer.Length)
            {
                if (_buffer.Length >= MaximumBufferSize)
                {
                    return false;
                }

                byte[] larger = ArrayPool<byte>.Shared.Rent(_buffer.Length * 2);
                Buffer.BlockCopy(_buffer, recordStart, larger, 0, retainedLength);
                ArrayPool<byte>.Shared.Return(_buffer);
                _buffer = larger;
                _visitor.SetBuffer(larger);
            }
            else if (retainedLength > 0 && recordStart != 0)
            {
                Buffer.BlockCopy(_buffer, recordStart, _buffer, 0, retainedLength);
            }

            recordStart = 0;
            _position = 0;
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

        private bool CompleteRecord(int recordStart, int recordLineNumber)
        {
            ref readonly RecordScan scan = ref _scan;
            _position = scan.Next;
            if (scan.Terminator == (byte)'\r')
            {
                _skipLineFeedAfterCarriageReturn = true;
            }

            _physicalLineSeparatorsConsumed += scan.LineBreaks + (scan.Terminator != 0 ? 1 : 0);
            if (!_allowEmpty && scan.FieldCount == 1 && _visitor.IsEmpty(0))
            {
                _visitor.Reset();
                return false;
            }

            int recordLength = scan.End - recordStart;
            bool isAscii = scan.Ascii == AsciiState.Unknown
                ? Ascii.IsValid(_buffer.AsSpan(recordStart, recordLength))
                : scan.Ascii == AsciiState.Ascii;
            if (!isAscii && _encoding.DecoderFallback is DecoderExceptionFallback)
                _encoding.GetCharCount(_buffer, recordStart, recordLength);
            _visitor.Complete(
                scan.FieldCount,
                _options.ColumnCountMismatchPolicy,
                isAscii,
                mayHaveEscapedQuotes: !scan.NoEscapedQuotes);
            if (_stringCache is { } stringCache) _visitor.InternFields(stringCache);
            _emittedRecordCount++;
            ReportProgress(_options, _emittedRecordCount, recordLineNumber);
            int endLineNumber = recordLineNumber + scan.LineBreaks;
            _currentPhysicalLineNumber = recordLineNumber;
            _currentPhysicalEndLineNumber = endLineNumber;
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
                // This is a record boundary, not the start of the input. A literal U+FEFF
                // here belongs to the field; preserve decoder policy without stripping it.
                var fallbackEncoding = (Encoding)new UTF8Encoding(false).Clone();
                fallbackEncoding.DecoderFallback = _encoding.DecoderFallback;
                fallbackEncoding.EncoderFallback = _encoding.EncoderFallback;
                var reader = new StreamReader(
                    prefixedStream,
                    fallbackEncoding,
                    detectEncodingFromByteOrderMarks: false,
                    bufferSize: FallbackTextBufferSize,
                    leaveOpen: false);
                _fallback = new CsvStreamDataReaderRowSource(
                    reader,
                    _options,
                    _emittedRecordCount,
                    recordLineNumber - 1,
                    _stringCache);
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
