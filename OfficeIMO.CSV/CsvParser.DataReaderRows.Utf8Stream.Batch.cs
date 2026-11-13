#if NET8_0_OR_GREATER
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;
using System.Text;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    internal sealed partial class CsvUtf8StreamDataReaderRowSource
    {
        /// <summary>Blocks indexed per batch; the index stays small enough to remain in the first-level cache.</summary>
        private const int BatchBlocks = 64;
        private const int BatchEntryCapacity = 1 + BatchBlocks * BlockSize + BlockSize;
        /// <summary>A trimmed batch stores two entries per field and two row slots per record.</summary>
        private const int TrimmedBatchEntryCapacity = 2 * BatchBlocks * BlockSize;
        private const int MaximumBatchBackoff = 64;

        /// <summary>
        /// Buffer offsets of every delimiter and line separator outside quotes in the current batch, in
        /// order, after entry 0, which is the offset before the batch. Field <c>i</c> of a record whose
        /// first field ends at entry <c>k</c> spans from after entry <c>k + i - 1</c> to entry <c>k + i</c>.
        /// </summary>
        private int[]? _batchEnds;
        /// <summary>
        /// Entry numbers of the batch's line separators, in order; each ends a record, except the line
        /// feed of a CRLF pair. Entries after the last of them belong to an incomplete record.
        /// </summary>
        private int[]? _batchRowEnds;
        /// <summary>Entry number of the next record's first field end.</summary>
        private int _batchCursor;
        private int _batchRow;
        private int _batchRowCount;
        private bool _batchAscii;
        private bool _batchEscapedQuotes;
        private int _batchBackoff;
        private int _batchBackoffSize;
        /// <summary>
        /// Records can be read from a batch index: no field length limits apply, which the record
        /// scanner handles per field, and the hardware has vector support.
        /// </summary>
        private readonly bool _batchable;

        /// <summary>
        /// Delivers the next record of the current batch. Returns false, leaving the record for the
        /// record scanner, when the record starts with the comment character.
        /// </summary>
        private bool TryCompleteBatchRecord(out bool emitted)
        {
            if (_trim)
            {
                emitted = CompleteTrimmedBatchRecord();
                return true;
            }

            emitted = false;
            ref int ends = ref MemoryMarshal.GetArrayDataReference(_batchEnds!);
            ref byte buffer = ref MemoryMarshal.GetArrayDataReference(_buffer);
            int first = _batchCursor;
            int recordStart = Unsafe.Add(ref ends, first - 1) + 1;
            if (Unsafe.Add(ref buffer, recordStart) == _commentCharacter && MaySkipCommentRecord())
            {
                _batchRow = _batchRowCount = 0;
                BackOffBatching();
                return false;
            }

            ref int rowEnds = ref MemoryMarshal.GetArrayDataReference(_batchRowEnds!);
            int last = Unsafe.Add(ref rowEnds, _batchRow++);
            int end = Unsafe.Add(ref ends, last);
            int next = last + 1;
            _position = end + 1;
            _skipLineFeedAfterCarriageReturn = false;
            if (Unsafe.Add(ref buffer, end) == (byte)'\r')
            {
                if (_batchRow < _batchRowCount && Unsafe.Add(ref buffer, end + 1) == (byte)'\n')
                {
                    // The line feed of a CRLF pair is the next entry and line separator; the following
                    // record starts after it.
                    next++;
                    _batchRow++;
                    _position++;
                }
                else
                {
                    _skipLineFeedAfterCarriageReturn = true;
                }
            }

            _batchCursor = next;
            emitted = CompleteIndexedRecord(recordStart, end, first, last - first + 1, shift: 0);
            return true;
        }

        /// <summary>
        /// Delivers the next record of a trimmed batch, whose fields are index entry pairs. Comment
        /// records never enter a trimmed batch.
        /// </summary>
        private bool CompleteTrimmedBatchRecord()
        {
            ref int rowSlots = ref MemoryMarshal.GetArrayDataReference(_batchRowEnds!);
            int row = _batchRow++;
            int first = row == 0 ? 0 : Unsafe.Add(ref rowSlots, 2 * row - 2);
            int last = Unsafe.Add(ref rowSlots, 2 * row);
            int next = Unsafe.Add(ref rowSlots, 2 * row + 1);
            int recordStart = _position;
            _position = next;
            _skipLineFeedAfterCarriageReturn = false;
            return CompleteIndexedRecord(recordStart, next - 1, first + 1, (last - first) >> 1, shift: 1);
        }

        /// <summary>Publishes a batch record, which spans one physical line, as <see cref="CompleteRecord"/> does.</summary>
        private bool CompleteIndexedRecord(int recordStart, int end, int first, int fieldCount, int shift)
        {
            int lineNumber = ++_physicalLineSeparatorsConsumed;
            _visitor.SetIndexedRecord(first, fieldCount, shift);
            if (fieldCount == 1 && IsSkippedEmptyRecord())
            {
                return false;
            }

            bool isAscii = _batchAscii || IsAsciiRecord(recordStart, end);
            _visitor.Complete(fieldCount, _options.ColumnCountMismatchPolicy, isAscii, mayHaveEscapedQuotes: _batchEscapedQuotes);
            if (_stringCache is { } stringCache) _visitor.InternFields(stringCache);
            _emittedRecordCount++;
            ReportProgress(_options, _emittedRecordCount, lineNumber);
            _currentPhysicalLineNumber = lineNumber;
            _currentPhysicalEndLineNumber = lineNumber;
            return true;
        }

        /// <summary>
        /// Whether a record that starts with the comment character may be skipped, as
        /// <see cref="ShouldSkipCommentRecord"/> decides; only the record scanner skips it. Otherwise it
        /// is an ordinary record and stays in the batch.
        /// </summary>
        [MethodImpl(MethodImplOptions.NoInlining)]
        private bool MaySkipCommentRecord()
        {
            CsvParsingSettings options = _options;
            return options.SkipCommentRows ||
                (options.SkipCommentRowsBeforeHeader && options.HasHeaderRow && options.Header is null &&
                 _emittedRecordCount <= GetParserInitialRecordsToSkip(options));
        }

        [MethodImpl(MethodImplOptions.NoInlining)]
        private bool IsSkippedEmptyRecord()
        {
            if (_allowEmpty || !_visitor.IsEmpty(0))
            {
                return false;
            }

            _visitor.Reset();
            return true;
        }

        /// <summary>Checks a record of a batch with non-ASCII bytes, validating it when the decoder rejects invalid UTF-8.</summary>
        [MethodImpl(MethodImplOptions.NoInlining)]
        private bool IsAsciiRecord(int recordStart, int end)
        {
            if (Ascii.IsValid(_buffer.AsSpan(recordStart, end - recordStart)))
            {
                return true;
            }

            if (_encoding.DecoderFallback is DecoderExceptionFallback)
            {
                _encoding.GetCharCount(_buffer, recordStart, end - recordStart);
            }

            return false;
        }

        /// <summary>
        /// Indexes the complete records in up to <see cref="BatchBlocks"/> blocks from the current
        /// record start. Indexing stops before a record with quoting the block rules cannot accept or
        /// a line separator inside quotes, and before the final partial block; the record scanner reads those.
        /// </summary>
        private bool TryIndexBatch(CancellationToken cancellationToken)
        {
            _batchRow = _batchRowCount = 0;
            if (!_batchable)
            {
                return false;
            }

            if (_batchBackoff > 0)
            {
                _batchBackoff--;
                return false;
            }

            int blocks = Math.Min((_length - _position) / BlockSize, BatchBlocks);
            if (blocks == 0)
            {
                return false;
            }

            cancellationToken.ThrowIfCancellationRequested();
            ThrowIfCancellationRequested(_options);
            int[] ends = _batchEnds ??= CreateBatchIndex();
            int[] rowEnds = _batchRowEnds ??= System.Buffers.ArrayPool<int>.Shared.Rent(_trim ? TrimmedBatchEntryCapacity : BatchEntryCapacity);
            ends[0] = _position - 1;
            int rows;
            if (_trim)
            {
                rows = IndexTrimmedBlocks(ends, rowEnds, blocks);
            }
            else
#if NET9_0_OR_GREATER
            if (Avx512Vbmi2.IsSupported && Avx512Vbmi.IsSupported && Bmi2.X64.IsSupported && Vector512.IsHardwareAccelerated)
            {
                rows = IndexBlocksCompressed(ends, rowEnds, blocks);
            }
            else if (Avx512F.IsSupported && Vector512.IsHardwareAccelerated)
            {
                rows = IndexBlocksDwordCompressed(ends, rowEnds, blocks);
            }
            else
#endif
            if (Avx2.IsSupported && Vector256.IsHardwareAccelerated)
            {
                rows = IndexBlocksByTable(ends, rowEnds, blocks);
            }
            else
            {
                rows = IndexBlocks(ends, rowEnds, blocks);
            }

            if (rows == 0)
            {
                BackOffBatching();
                return false;
            }

            _batchBackoffSize = 0;
            _batchCursor = 1;
            _batchRowCount = rows;
            return true;
        }

        private int[] CreateBatchIndex()
        {
            int[] ends = System.Buffers.ArrayPool<int>.Shared.Rent(_trim ? TrimmedBatchEntryCapacity : BatchEntryCapacity);
            _visitor.SetBatchIndex(ends);
            return ends;
        }

        /// <summary>
        /// Returns the pooled batch index. The index lives as long as the reader, so pooling it keeps a
        /// short read from promoting it to an older generation.
        /// </summary>
        private void ReturnBatchIndex()
        {
            int[]? ends = _batchEnds;
            int[]? rowEnds = _batchRowEnds;
            _batchEnds = _batchRowEnds = null;
            _batchRowCount = _batchCursor = 0;
            if (ends is not null) System.Buffers.ArrayPool<int>.Shared.Return(ends);
            if (rowEnds is not null) System.Buffers.ArrayPool<int>.Shared.Return(rowEnds);
        }

        /// <summary>
        /// After a batch that delivered no record (a record longer than the window, irregular at its
        /// start, or a comment record), skips batching for a growing number of records so records the
        /// batch cannot take do not each pay for an index that is thrown away.
        /// </summary>
        private void BackOffBatching()
        {
            _batchBackoffSize = _batchBackoffSize == 0 ? 1 : Math.Min(_batchBackoffSize * 2, MaximumBatchBackoff);
            _batchBackoff = _batchBackoffSize;
        }

        /// <summary>Quote state carried from one block to the next.</summary>
        private struct BlockQuoteState
        {
            internal ulong InsideQuotes;
            internal bool ClosingQuotePending;
            internal ulong OpeningQuoteAllowed;
            internal bool EscapedQuotes;
        }

        /// <summary>
        /// Classifies a block that contains quotes or continues a quoted field. Returns the delimiters
        /// and line separators outside quotes; when the block breaks the rules, returns the separators
        /// before the record that contains the first violation and sets <paramref name="stop"/>.
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static ulong ClassifyQuotedBlock(ulong delimiters, ulong quotes, ulong lineSeparators, ref BlockQuoteState state, out bool stop)
        {
            // Quote parity marks bytes inside quotes. A closing quote must be followed by a quote
            // (an escaped pair), a delimiter or a line separator; an opening quote must start a field
            // or follow a closing quote. A line separator inside quotes adds a physical line, which
            // the record scanner counts, so the batch ends before that record too.
            ulong inside = PrefixXor(quotes) ^ state.InsideQuotes;
            ulong opening = quotes & inside;
            ulong closing = quotes & ~inside;
            ulong separators = (delimiters | lineSeparators) & ~inside;
            ulong boundaries = separators | closing;
            ulong followsClosing = quotes | delimiters | lineSeparators;
            // Recorded before any violation: records kept ahead of it may contain an escaped pair.
            state.EscapedQuotes |= (((closing << 1) | (state.ClosingQuotePending ? 1UL : 0UL)) & quotes) != 0;
            ulong violations = ((closing << 1) & ~followsClosing) |
                (opening & ~((boundaries << 1) | state.OpeningQuoteAllowed)) |
                (lineSeparators & inside) |
                (state.ClosingQuotePending && (followsClosing & 1UL) == 0 ? 1UL : 0UL);
            if (violations != 0)
            {
                stop = true;
                ulong recordEnds = lineSeparators & ~inside & ((1UL << BitOperations.TrailingZeroCount(violations)) - 1UL);
                return recordEnds == 0 ? 0 : separators & (ulong.MaxValue >> BitOperations.LeadingZeroCount(recordEnds));
            }

            stop = false;
            state.ClosingQuotePending = (closing >> 63) != 0;
            state.InsideQuotes = (ulong)((long)inside >> 63);
            state.OpeningQuoteAllowed = boundaries >> 63;
            return separators;
        }

        /// <summary>
        /// Appends the entry numbers of a block's line separators, whose entries start at
        /// <paramref name="count"/>, to the row ends and returns the new row end count. An entry's
        /// number within the block is the count of separators before it.
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static int AppendRowEnds(ref int rowEnds, int rows, int count, ulong structural, ulong lineSeparators)
        {
            // A block usually holds zero, one or two line separators. Two slots are written unconditionally
            // so the common counts take no data-dependent branch; slots past the count are scratch.
            ulong ends = structural & lineSeparators;
            int found = BitOperations.PopCount(ends);
            ref int slot = ref Unsafe.Add(ref rowEnds, rows);
            for (int index = 0; index < 2; index++)
            {
                Unsafe.Add(ref slot, index) = count + BitOperations.PopCount(structural & ((1UL << BitOperations.TrailingZeroCount(ends)) - 1UL));
                ends &= ends - 1UL;
            }

            for (int index = 2; index < found; index++)
            {
                Unsafe.Add(ref slot, index) = count + BitOperations.PopCount(structural & ((1UL << BitOperations.TrailingZeroCount(ends)) - 1UL));
                ends &= ends - 1UL;
            }

            return rows + found;
        }

#if NET9_0_OR_GREATER
        /// <summary>
        /// <see cref="AppendRowEnds"/> with the separators' entry numbers gathered by PEXT, which every
        /// processor with AVX-512 VBMI2 executes in a single cycle.
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static int AppendRowEndsGathered(ref int rowEnds, int rows, int count, ulong structural, ulong lineSeparators)
        {
            ulong ranks = Bmi2.X64.ParallelBitExtract(structural & lineSeparators, structural);
            int found = BitOperations.PopCount(ranks);
            ref int slot = ref Unsafe.Add(ref rowEnds, rows);
            for (int index = 0; index < 2; index++)
            {
                Unsafe.Add(ref slot, index) = count + BitOperations.TrailingZeroCount(ranks);
                ranks &= ranks - 1UL;
            }

            for (int index = 2; index < found; index++)
            {
                Unsafe.Add(ref slot, index) = count + BitOperations.TrailingZeroCount(ranks);
                ranks &= ranks - 1UL;
            }

            return rows + found;
        }
#endif

        /// <summary>
        /// Index with bit iteration, for processors without AVX2 (including Arm64); eight positions are
        /// written per block without a data-dependent branch.
        /// </summary>
        private int IndexBlocks(int[] ends, int[] rowEnds, int blocks)
        {
            ref byte data = ref MemoryMarshal.GetArrayDataReference(_buffer);
            ref int destination = ref MemoryMarshal.GetArrayDataReference(ends);
            ref int rowEndSlots = ref MemoryMarshal.GetArrayDataReference(rowEnds);
            BlockQuoteState state = default;
            state.OpeningQuoteAllowed = 1;
            ulong nonAscii = 0;
            int count = 1;
            int rows = 0;
            int position = _position;
            for (int block = 0; block < blocks; block++, position += BlockSize)
            {
                LoadBlockMasks(ref data, position, out ulong delimiters, out ulong quotes,
                    out ulong carriageReturns, out ulong lineFeeds, out ulong high);
                ulong lineSeparators = carriageReturns | lineFeeds;
                ulong structural;
                bool stop = false;
                if ((quotes | state.InsideQuotes) == 0 && !state.ClosingQuotePending)
                {
                    structural = delimiters | lineSeparators;
                    state.OpeningQuoteAllowed = structural >> 63;
                }
                else
                {
                    structural = ClassifyQuotedBlock(delimiters, quotes, lineSeparators, ref state, out stop);
                }

                nonAscii |= high;
                int found = BitOperations.PopCount(structural);
                ref int slot = ref Unsafe.Add(ref destination, count);
                ulong remaining = structural;
                for (int index = 0; index < 8; index++)
                {
                    Unsafe.Add(ref slot, index) = position + BitOperations.TrailingZeroCount(remaining);
                    remaining &= remaining - 1UL;
                }

                for (int index = 8; index < found; index++)
                {
                    Unsafe.Add(ref slot, index) = position + BitOperations.TrailingZeroCount(remaining);
                    remaining &= remaining - 1UL;
                }

                rows = AppendRowEnds(ref rowEndSlots, rows, count, structural, lineSeparators);
                count += found;
                if (stop)
                {
                    break;
                }
            }

            _batchAscii = nonAscii == 0;
            _batchEscapedQuotes = state.EscapedQuotes;
            return rows;
        }

        /// <summary>
        /// Index with trimming. Field <c>i</c> of a record whose entries start at <c>k</c> spans from after
        /// entry <c>k + 2i</c> to entry <c>k + 2i + 1</c>: its text without the surrounding whitespace, a
        /// quoted field with its quotes, which the reader removes. Row slot <c>2r</c> is the entry count
        /// after record <c>r</c> and slot <c>2r + 1</c> the offset where the next record starts. Indexing
        /// stops before a record whose quoting or edge whitespace the record scanner must handle, a record
        /// with a line separator inside quotes, a record that may be a skipped comment, and a carriage
        /// return whose following byte is not buffered.
        /// </summary>
        private int IndexTrimmedBlocks(int[] ends, int[] rowEnds, int blocks)
        {
            ref byte data = ref MemoryMarshal.GetArrayDataReference(_buffer);
            ref int entries = ref MemoryMarshal.GetArrayDataReference(ends);
            ref int rowSlots = ref MemoryMarshal.GetArrayDataReference(rowEnds);
            int length = _length;
            int position = _position;
            int fieldStart = position;
            // The current field's first kept byte, once a block has shown it, and the offset after the
            // last kept byte before the current block. Separators count as kept bytes, so the first bounds
            // the text of a field that is only whitespace and the second always lies within the field.
            int textStart = 0;
            bool textStartPending = true;
            int keptEnd = position;
            int entry = 0;
            int rows = 0;
            int fieldQuotes = 0;
            ulong insideQuotes = 0;
            ulong nonAscii = 0;
            ulong skippedLineFeed = 0;
            bool escapedQuotes = false;
            if (StartsSkippedComment(position))
            {
                return 0;
            }

            for (int block = 0; block < blocks; block++, position += BlockSize)
            {
                LoadBlockMasks(ref data, position, out ulong delimiters, out ulong quotes,
                    out ulong carriageReturns, out ulong lineFeeds, out ulong high);
                ulong lineSeparators = carriageReturns | lineFeeds;
                ulong separators = delimiters | lineSeparators;
                ulong inside = insideQuotes;
                if (quotes != 0)
                {
                    inside = PrefixXor(quotes) ^ insideQuotes;
                    insideQuotes = (ulong)((long)inside >> 63);
                }

                // Trimming removes whitespace outside quotes; everything else, separators included, is kept.
                ulong kept = ~(LoadWhiteSpaceMask(ref data, position) & ~inside & ~separators);
                ulong structural = separators & ~inside & ~skippedLineFeed;
                skippedLineFeed = 0;
                ulong innerLineSeparators = lineSeparators & inside;
                if (innerLineSeparators != 0)
                {
                    // Only records that end before the first line separator inside quotes are complete.
                    structural &= (1UL << BitOperations.TrailingZeroCount(innerLineSeparators)) - 1UL;
                    blocks = block + 1;
                }

                nonAscii |= high;
                if (textStartPending)
                {
                    // The field started in an earlier block, or at bit 1 after a CRLF pair split across blocks.
                    ulong keptAfter = kept & (ulong.MaxValue << Math.Max(fieldStart - position, 0));
                    textStartPending = keptAfter == 0;
                    textStart = position + BitOperations.TrailingZeroCount(keptAfter);
                }

                while (structural != 0)
                {
                    int bit = BitOperations.TrailingZeroCount(structural);
                    ulong before = (1UL << bit) - 1UL;
                    fieldQuotes += BitOperations.PopCount(quotes & before);
                    quotes &= ~before;
                    int separator = position + bit;
                    ulong keptBefore = kept & before;
                    int textEnd = Math.Max(
                        keptBefore != 0 ? position + BlockSize - BitOperations.LeadingZeroCount(keptBefore) : keptEnd,
                        textStart);
                    if (fieldQuotes != 0
                        ? !IsTrimmedQuotedField(fieldStart, textStart, textEnd, fieldQuotes, ref escapedQuotes)
                        : textStart < textEnd &&
                            (Unsafe.Add(ref data, textStart) >= 0x80 || Unsafe.Add(ref data, textEnd - 1) >= 0x80) &&
                            HasNonAsciiWhiteSpaceEdge(textStart, textEnd))
                    {
                        return CompleteTrimmedIndex(rows, nonAscii, escapedQuotes);
                    }

                    ref int pair = ref Unsafe.Add(ref entries, entry);
                    pair = textStart - 1;
                    Unsafe.Add(ref pair, 1) = textEnd;
                    entry += 2;
                    fieldQuotes = 0;
                    fieldStart = separator + 1;
                    if ((lineSeparators & (1UL << bit)) != 0)
                    {
                        if (Unsafe.Add(ref data, separator) == (byte)'\r')
                        {
                            if (fieldStart == length)
                            {
                                return CompleteTrimmedIndex(rows, nonAscii, escapedQuotes);
                            }

                            if (Unsafe.Add(ref data, fieldStart) == (byte)'\n')
                            {
                                // The line feed of a CRLF pair is no separator of its own.
                                fieldStart++;
                                if (bit == 63)
                                {
                                    skippedLineFeed = 1;
                                }
                                else
                                {
                                    structural &= ~(1UL << (bit + 1));
                                }
                            }
                        }

                        ref int slot = ref Unsafe.Add(ref rowSlots, 2 * rows);
                        slot = entry;
                        Unsafe.Add(ref slot, 1) = fieldStart;
                        rows++;
                        if (StartsSkippedComment(fieldStart))
                        {
                            return CompleteTrimmedIndex(rows, nonAscii, escapedQuotes);
                        }
                    }

                    int next = fieldStart - position;
                    ulong keptAfter = next < BlockSize ? kept & (ulong.MaxValue << next) : 0;
                    textStartPending = keptAfter == 0;
                    textStart = position + BitOperations.TrailingZeroCount(keptAfter);
                    structural &= structural - 1UL;
                }

                fieldQuotes += BitOperations.PopCount(quotes);
                if (kept != 0)
                {
                    keptEnd = position + BlockSize - BitOperations.LeadingZeroCount(kept);
                }
            }

            return CompleteTrimmedIndex(rows, nonAscii, escapedQuotes);
        }

        /// <summary>
        /// Whether trimmed field text is one quoted section, as <see cref="TryTrimQuotedField"/> accepts it:
        /// it starts and ends with a quote and every inner quote is part of an escaped pair. Strict quoting
        /// admits whitespace only after the closing quote.
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private bool IsTrimmedQuotedField(int fieldStart, int textStart, int textEnd, int quotes, ref bool escapedQuotes)
        {
            ref byte data = ref MemoryMarshal.GetArrayDataReference(_buffer);
            if ((quotes & 1) != 0 || textEnd - textStart < 2 ||
                Unsafe.Add(ref data, textStart) != (byte)'"' || Unsafe.Add(ref data, textEnd - 1) != (byte)'"' ||
                (_strictQuotes && textStart != fieldStart))
            {
                return false;
            }

            if (quotes > 2)
            {
                if (!HasOnlyEscapedQuotePairs(_buffer.AsSpan(textStart + 1, textEnd - textStart - 2)))
                {
                    return false;
                }

                escapedQuotes = true;
            }

            return true;
        }

        /// <summary>Marks the ASCII whitespace bytes of a block, as <see cref="IsAsciiWhiteSpace"/> classifies them.</summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static ulong LoadWhiteSpaceMask(ref byte data, int position)
        {
            if (Vector512.IsHardwareAccelerated)
            {
                Vector512<byte> values = Vector512.LoadUnsafe(ref data, (nuint)position);
                return (Vector512.Equals(values, Vector512.Create((byte)' ')) |
                    Vector512.LessThanOrEqual(values - Vector512.Create((byte)9), Vector512.Create((byte)4))).ExtractMostSignificantBits();
            }

            ulong mask = 0;
            if (Vector256.IsHardwareAccelerated)
            {
                for (int offset = 0; offset < BlockSize; offset += 32)
                {
                    Vector256<byte> values = Vector256.LoadUnsafe(ref data, (nuint)(position + offset));
                    mask |= (ulong)(Vector256.Equals(values, Vector256.Create((byte)' ')) |
                        Vector256.LessThanOrEqual(values - Vector256.Create((byte)9), Vector256.Create((byte)4))).ExtractMostSignificantBits() << offset;
                }

                return mask;
            }

            for (int offset = 0; offset < BlockSize; offset += 16)
            {
                Vector128<byte> values = Vector128.LoadUnsafe(ref data, (nuint)(position + offset));
                mask |= (ulong)(Vector128.Equals(values, Vector128.Create((byte)' ')) |
                    Vector128.LessThanOrEqual(values - Vector128.Create((byte)9), Vector128.Create((byte)4))).ExtractMostSignificantBits() << offset;
            }

            return mask;
        }

        private int CompleteTrimmedIndex(int rows, ulong nonAscii, bool escapedQuotes)
        {
            _batchAscii = nonAscii == 0;
            _batchEscapedQuotes = escapedQuotes;
            return rows;
        }

        /// <summary>Whether a buffered record starts with the comment character and may be skipped, which only the record scanner does.</summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private bool StartsSkippedComment(int recordStart) =>
            recordStart < _length && _buffer[recordStart] == _commentCharacter && MaySkipCommentRecord();

#if NET9_0_OR_GREATER
        /// <summary>Index with byte compression: the offsets of a block's separators are packed in one instruction.</summary>
        private int IndexBlocksCompressed(int[] ends, int[] rowEnds, int blocks)
        {
            ref byte data = ref MemoryMarshal.GetArrayDataReference(_buffer);
            ref int destination = ref MemoryMarshal.GetArrayDataReference(ends);
            ref int rowEndSlots = ref MemoryMarshal.GetArrayDataReference(rowEnds);
            Vector512<byte> delimiterValue = Vector512.Create(_delimiter);
            Vector512<byte> quoteValue = Vector512.Create((byte)'"');
            Vector512<byte> carriageReturnValue = Vector512.Create((byte)'\r');
            Vector512<byte> lineFeedValue = Vector512.Create((byte)'\n');
            Vector512<byte> offsets = Vector512<byte>.Indices;
            Vector512<byte> nonAscii = Vector512<byte>.Zero;
            BlockQuoteState state = default;
            state.OpeningQuoteAllowed = 1;
            int count = 1;
            int rows = 0;
            int position = _position;
            for (int block = 0; block < blocks; block++, position += BlockSize)
            {
                Vector512<byte> values = Vector512.LoadUnsafe(ref data, (nuint)position);
                nonAscii |= values;
                Vector512<byte> delimiterBytes = Vector512.Equals(values, delimiterValue);
                Vector512<byte> lineSeparatorBytes = Vector512.Equals(values, carriageReturnValue) | Vector512.Equals(values, lineFeedValue);
                ulong quotes = Vector512.Equals(values, quoteValue).ExtractMostSignificantBits() & _quoteMask;
                ulong lineSeparators = lineSeparatorBytes.ExtractMostSignificantBits();
                Vector512<byte> selected;
                ulong structural;
                bool stop = false;
                if ((quotes | state.InsideQuotes) == 0 && !state.ClosingQuotePending)
                {
                    selected = delimiterBytes | lineSeparatorBytes;
                    structural = selected.ExtractMostSignificantBits();
                    state.OpeningQuoteAllowed = structural >> 63;
                }
                else
                {
                    structural = ClassifyQuotedBlock(delimiterBytes.ExtractMostSignificantBits(), quotes, lineSeparators, ref state, out stop);
                    selected = ExpandMask(structural);
                }

                Vector512<byte> packed = Avx512Vbmi2.Compress(Vector512<byte>.Zero, selected, offsets);
                Vector512<int> blockStart = Vector512.Create(position);
                int found = BitOperations.PopCount(structural);
                ref int slot = ref Unsafe.Add(ref destination, count);
                (Avx512F.ConvertToVector512Int32(packed.GetLower().GetLower()) + blockStart).StoreUnsafe(ref slot);
                if (found > 16)
                {
                    (Avx512F.ConvertToVector512Int32(packed.GetLower().GetUpper()) + blockStart).StoreUnsafe(ref slot, 16);
                    if (found > 32)
                    {
                        (Avx512F.ConvertToVector512Int32(packed.GetUpper().GetLower()) + blockStart).StoreUnsafe(ref slot, 32);
                        (Avx512F.ConvertToVector512Int32(packed.GetUpper().GetUpper()) + blockStart).StoreUnsafe(ref slot, 48);
                    }
                }

                rows = AppendRowEndsGathered(ref rowEndSlots, rows, count, structural, lineSeparators);
                count += found;
                if (stop)
                {
                    break;
                }
            }

            _batchAscii = nonAscii.ExtractMostSignificantBits() == 0;
            _batchEscapedQuotes = state.EscapedQuotes;
            return rows;
        }

        /// <summary>Turns bit <c>i</c> of <paramref name="mask"/> into an all-ones byte <c>i</c>.</summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static Vector512<byte> ExpandMask(ulong mask)
        {
            // Byte i takes source byte i / 8 and tests bit i % 8 of it.
            Vector512<byte> spread = Avx512Vbmi.PermuteVar64x8(Vector512.Create(mask).AsByte(), Vector512<byte>.Indices >>> 3);
            Vector512<byte> bits = Vector512.Create(0x8040201008040201UL).AsByte();
            return Vector512.Equals(spread & bits, bits);
        }
#endif
    }
}
#endif
