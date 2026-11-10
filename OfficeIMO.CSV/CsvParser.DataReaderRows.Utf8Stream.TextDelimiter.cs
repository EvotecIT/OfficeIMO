using System.Buffers;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    internal sealed partial class CsvUtf8StreamDataReaderRowSource
    {
        /// <summary>The UTF-8 bytes of a multi-character ASCII delimiter, or null.</summary>
        private readonly byte[]? _delimiterText;

        /// <summary>The one-byte delimiter as a span for the shared quote rules.</summary>
        private readonly byte[] _delimiterByte;

        private enum TextRecordResult
        {
            NeedMoreData,
            Record,
            Comment,
            Invalid,
            Unterminated
        }

        private enum TextQuoteResult
        {
            Complete,
            InQuotes,
            Invalid
        }

        /// <summary>Quote and field state of <see cref="ParseTextDelimitedQuotes"/>.</summary>
        private struct TextQuoteState
        {
            internal bool InQuotes;
            internal bool AfterClosingQuote;
            internal bool FieldWasQuoted;
            internal bool WhiteSpaceOnly;
            internal int FieldIndex;
            internal int FieldLength;
        }

        /// <summary>
        /// Reads records separated by a multi-character delimiter with the canonical line-based rules:
        /// a record is one physical line unless that line contains a quote, a quoted record continues
        /// over whole lines while a quote is open, and a strict-mode error on the first line skips only
        /// that line.
        /// </summary>
        private bool ReadTextDelimitedRecord(CancellationToken cancellationToken)
        {
            bool checkCancellation = false;
            while (true)
            {
                if (checkCancellation)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    ThrowIfCancellationRequested(_options);
                }

                checkCancellation = true;
                if (!EnsureBuffered(cancellationToken))
                {
                    return false;
                }

                int recordStart = _position;
                int recordLineNumber = _physicalLineSeparatorsConsumed + 1;
                TextRecordResult result;
                int next;
                int separators;
                int errorLineOffset;
                while ((result = ScanTextDelimitedRecord(recordStart, out next, out separators, out errorLineOffset)) ==
                    TextRecordResult.NeedMoreData)
                {
                    if (!FillForCurrentRecord(ref recordStart, cancellationToken) && !_endOfStream)
                    {
                        throw new CsvException($"A CSV record is longer than {MaximumBufferSize} bytes.");
                    }
                }

                if (result == TextRecordResult.Record)
                {
                    if (CompleteRecord(recordStart, recordLineNumber))
                    {
                        return true;
                    }

                    continue;
                }

                _position = next;
                _physicalLineSeparatorsConsumed += separators;
                if (result == TextRecordResult.Invalid)
                {
                    int lineNumber = recordLineNumber + errorLineOffset;
                    SkipOrThrowParseError(_options, new CsvParseException("Invalid quoted field.", lineNumber), lineNumber);
                }
                else if (result == TextRecordResult.Unterminated)
                {
                    int lineNumber = recordLineNumber + errorLineOffset;
                    SkipOrThrowParseError(_options, new CsvParseException("Unterminated quoted field.", lineNumber), lineNumber);
                }
            }
        }

        /// <summary>
        /// Scans the record at <paramref name="recordStart"/>. A record is stored in <see cref="_scan"/>
        /// and the visitor; otherwise <paramref name="next"/> and <paramref name="separators"/> describe
        /// the lines it consumed and <paramref name="errorLineOffset"/> the line of its error.
        /// </summary>
        private TextRecordResult ScanTextDelimitedRecord(int recordStart, out int next, out int separators, out int errorLineOffset)
        {
            next = recordStart;
            separators = 0;
            errorLineOffset = 0;
            if (!TryFindLineEnd(recordStart, out int firstEnd, out int firstNext))
            {
                return TextRecordResult.NeedMoreData;
            }

            ReadOnlySpan<byte> firstLine = _buffer.AsSpan(recordStart, firstEnd - recordStart);
            bool hasQuote = !_quotesAreData && firstLine.Contains((byte)'"');
            int firstSeparators = firstNext > firstEnd ? 1 : 0;
            if (!firstLine.IsEmpty && firstLine[0] == _commentCharacter && ShouldSkipTextDelimitedComment(firstLine))
            {
                return SkipTextDelimitedComment(recordStart, firstEnd, firstNext, hasQuote, out next, out separators);
            }

            _visitor.Reset();
            if (!hasQuote)
            {
                VisitTextDelimitedLine(recordStart, firstEnd);
                SetTextDelimitedScan(firstEnd, firstNext, lineBreaks: 0);
                return TextRecordResult.Record;
            }

            _flexibleFieldBytes ??= new byte[256];
            var state = new TextQuoteState { WhiteSpaceOnly = true };
            TextQuoteResult parse = ParseTextDelimitedQuotes(recordStart, firstEnd, _delimiterText, _strictQuotes, visit: true, ref state);
            if (parse == TextQuoteResult.Invalid)
            {
                next = firstNext;
                separators = firstSeparators;
                return TextRecordResult.Invalid;
            }

            if (parse == TextQuoteResult.Complete)
            {
                VisitFlexibleField(state.FieldIndex, state.FieldLength, state.FieldWasQuoted);
                SetTextDelimitedScan(firstEnd, firstNext, lineBreaks: 0);
                return TextRecordResult.Record;
            }

            // The quote is still open: whole lines join the record until it closes.
            var lineState = new TextQuoteState { InQuotes = state.InQuotes, AfterClosingQuote = state.AfterClosingQuote };
            int position = firstNext;
            int recordEnd = firstEnd;
            int continuations = 0;
            separators = firstSeparators;
            while (lineState.InQuotes)
            {
                if (position == _length)
                {
                    if (!_endOfStream)
                    {
                        return TextRecordResult.NeedMoreData;
                    }

                    next = position;
                    errorLineOffset = continuations;
                    return TextRecordResult.Unterminated;
                }

                if (!TryFindLineEnd(position, out int lineEnd, out int lineNext))
                {
                    return TextRecordResult.NeedMoreData;
                }

                ParseTextDelimitedQuotes(position, lineEnd, _delimiterText, strict: false, visit: false, ref lineState);
                continuations++;
                separators += lineNext > lineEnd ? 1 : 0;
                recordEnd = lineEnd;
                position = lineNext;
            }

            _visitor.Reset();
            state = new TextQuoteState { WhiteSpaceOnly = true };
            parse = ParseTextDelimitedQuotes(recordStart, recordEnd, _delimiterText, _strictQuotes, visit: true, ref state);
            if (parse != TextQuoteResult.Complete)
            {
                next = position;
                errorLineOffset = continuations;
                return parse == TextQuoteResult.Invalid ? TextRecordResult.Invalid : TextRecordResult.Unterminated;
            }

            VisitFlexibleField(state.FieldIndex, state.FieldLength, state.FieldWasQuoted);
            SetTextDelimitedScan(recordEnd, position, continuations);
            return TextRecordResult.Record;
        }

        private void SetTextDelimitedScan(int end, int next, int lineBreaks)
        {
            _scan = new RecordScan
            {
                End = end,
                Next = next,
                FieldCount = _visitor.FieldCount,
                LineBreaks = lineBreaks,
                // The line ends were consumed whole, so a CR is never followed by a pending LF.
                Terminator = next > end ? (byte)'\n' : (byte)0,
                NoEscapedQuotes = true
            };
        }

        /// <summary>Finds the end of the physical line at <paramref name="position"/> and the start of the next.</summary>
        private bool TryFindLineEnd(int position, out int lineEnd, out int next)
        {
            int found = _buffer.AsSpan(position, _length - position).IndexOfAny((byte)'\r', (byte)'\n');
            if (found < 0)
            {
                lineEnd = _length;
                next = _length;
                return _endOfStream;
            }

            lineEnd = position + found;
            if (_buffer[lineEnd] == (byte)'\r')
            {
                if (lineEnd + 1 == _length && !_endOfStream)
                {
                    next = 0;
                    return false;
                }

                next = lineEnd + 1 < _length && _buffer[lineEnd + 1] == (byte)'\n' ? lineEnd + 2 : lineEnd + 1;
                return true;
            }

            next = lineEnd + 1;
            return true;
        }

        /// <summary>Visits the fields of a line without quotes, trimming them when requested.</summary>
        private void VisitTextDelimitedLine(int start, int end)
        {
            ReadOnlySpan<byte> delimiter = _delimiterText;
            int fieldIndex = 0;
            int fieldStart = start;
            while (true)
            {
                int found = _buffer.AsSpan(fieldStart, end - fieldStart).IndexOf(delimiter);
                if (found < 0)
                {
                    VisitUnquotedField(fieldIndex, fieldStart, end);
                    return;
                }

                VisitUnquotedField(fieldIndex++, fieldStart, fieldStart + found);
                fieldStart += found + delimiter.Length;
            }
        }

        /// <summary>
        /// Runs the canonical quoted-record rules for <paramref name="delimiter"/> over
        /// <paramref name="start"/>..<paramref name="end"/>.
        /// With <paramref name="visit"/> the completed fields are stored, except the last one.
        /// </summary>
        private TextQuoteResult ParseTextDelimitedQuotes(int start, int end, ReadOnlySpan<byte> delimiter, bool strict, bool visit, ref TextQuoteState state)
        {
            byte[] buffer = _buffer;
            int position = start;
            while (position < end)
            {
                byte value = buffer[position];
                if (state.InQuotes)
                {
                    if (value == (byte)'"')
                    {
                        if (position + 1 < end && buffer[position + 1] == (byte)'"')
                        {
                            if (visit) AppendFlexibleBytes(ref state.FieldLength, position, 1);
                            state.WhiteSpaceOnly = false;
                            position += 2;
                            continue;
                        }

                        state.InQuotes = false;
                        state.AfterClosingQuote = true;
                        position++;
                        continue;
                    }

                    int quotedLength = ClassifyTextCharacter(position, out bool quotedWhiteSpace);
                    if (visit) AppendFlexibleBytes(ref state.FieldLength, position, quotedLength);
                    state.WhiteSpaceOnly &= quotedWhiteSpace;
                    position += quotedLength;
                    continue;
                }

                if (value == delimiter[0] && buffer.AsSpan(position, end - position).StartsWith(delimiter))
                {
                    if (visit) VisitFlexibleField(state.FieldIndex, state.FieldLength, state.FieldWasQuoted);
                    state.FieldIndex++;
                    state.FieldLength = 0;
                    state.FieldWasQuoted = false;
                    state.WhiteSpaceOnly = true;
                    state.AfterClosingQuote = false;
                    position += delimiter.Length;
                    continue;
                }

                if (value == (byte)'"')
                {
                    if (state.AfterClosingQuote)
                    {
                        if (strict)
                        {
                            return TextQuoteResult.Invalid;
                        }

                        if (visit) AppendFlexibleBytes(ref state.FieldLength, position, 1);
                        state.WhiteSpaceOnly = false;
                        state.AfterClosingQuote = false;
                        position++;
                        continue;
                    }

                    if (_trim && state.WhiteSpaceOnly)
                    {
                        state.FieldLength = 0;
                    }

                    state.InQuotes = true;
                    state.FieldWasQuoted = true;
                    position++;
                    continue;
                }

                int length = ClassifyTextCharacter(position, out bool whiteSpace);
                if (state.AfterClosingQuote && whiteSpace && _trim)
                {
                    position += length;
                    continue;
                }

                if (state.AfterClosingQuote && strict)
                {
                    return TextQuoteResult.Invalid;
                }

                state.AfterClosingQuote = false;
                if (visit) AppendFlexibleBytes(ref state.FieldLength, position, length);
                state.WhiteSpaceOnly &= whiteSpace;
                position += length;
            }

            return state.InQuotes ? TextQuoteResult.InQuotes : TextQuoteResult.Complete;
        }

        /// <summary>
        /// Returns the byte length of the character at <paramref name="position"/>. Whitespace only
        /// changes the result when trimming, so it is classified only then.
        /// </summary>
        private int ClassifyTextCharacter(int position, out bool whiteSpace)
        {
            if (!_trim)
            {
                whiteSpace = false;
                return 1;
            }

            // The record is fully buffered and ends at a line separator or the end of input.
            TryClassifyWhiteSpace(position, out whiteSpace, out int length);
            return length;
        }

        private bool ShouldSkipTextDelimitedComment(ReadOnlySpan<byte> firstLine)
        {
            CsvParsingSettings options = _options;
            bool canReadW3CFieldsHeader = _emittedRecordCount <= GetParserInitialRecordsToSkip(options) &&
                options.HasHeaderRow && options.Header is null && options.RecognizeW3CFieldsHeader;
            if (canReadW3CFieldsHeader && IsW3CFieldsLine(DecodeLinePrefix(firstLine), options))
            {
                return false;
            }

            return options.SkipCommentRows ||
                (options.HasHeaderRow && options.Header is null && options.SkipCommentRowsBeforeHeader &&
                    _emittedRecordCount <= GetParserInitialRecordsToSkip(options));
        }

        // "#Fields:" is eight UTF-16 characters, which never take more than 32 UTF-8 bytes.
        private string DecodeLinePrefix(ReadOnlySpan<byte> line) => _encoding.GetString(line.Slice(0, Math.Min(line.Length, 32)));

        /// <summary>
        /// Skips a comment record. One with a quote that leaves a quote open and looks delimited
        /// continues over whole lines until the quote closes, as canonical comment skipping does.
        /// </summary>
        private TextRecordResult SkipTextDelimitedComment(int recordStart, int firstEnd, int firstNext, bool hasQuote, out int next, out int separators)
        {
            next = firstNext;
            separators = firstNext > firstEnd ? 1 : 0;
            if (!hasQuote || !LooksLikeDelimitedComment(_buffer.AsSpan(recordStart, firstEnd - recordStart), _delimiterText))
            {
                return TextRecordResult.Comment;
            }

            var state = new TextQuoteState();
            if (ParseTextDelimitedQuotes(recordStart, firstEnd, _delimiterText, strict: false, visit: false, ref state) != TextQuoteResult.InQuotes)
            {
                return TextRecordResult.Comment;
            }

            int position = firstNext;
            while (true)
            {
                if (position == _length)
                {
                    if (!_endOfStream)
                    {
                        return TextRecordResult.NeedMoreData;
                    }

                    next = position;
                    return TextRecordResult.Comment;
                }

                if (!TryFindLineEnd(position, out int lineEnd, out int lineNext))
                {
                    return TextRecordResult.NeedMoreData;
                }

                ParseTextDelimitedQuotes(position, lineEnd, _delimiterText, strict: false, visit: false, ref state);
                separators += lineNext > lineEnd ? 1 : 0;
                position = lineNext;
                if (!state.InQuotes)
                {
                    next = position;
                    return TextRecordResult.Comment;
                }
            }
        }


        private static readonly SearchValues<byte> DelimitedCommentBytes = SearchValues.Create(new[] { (byte)',', (byte)';', (byte)'|', (byte)'\t' });
    }
}
