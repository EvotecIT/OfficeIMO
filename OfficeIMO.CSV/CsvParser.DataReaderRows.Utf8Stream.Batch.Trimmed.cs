#if NET9_0_OR_GREATER
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
    internal sealed partial class CsvUtf8StreamDataReaderRowSource
    {
        /// <summary>The block masks the compressed trimmed index works from.</summary>
        private struct TrimmedBlockMasks
        {
            /// <summary>Delimiters and record ends outside quotes; the line feed of a CRLF pair is neither.</summary>
            internal ulong Structural;
            internal ulong RecordEnds;
            internal ulong Quotes;
            internal ulong Inside;
            /// <summary>ASCII whitespace outside quotes that is not a separator.</summary>
            internal ulong WhiteSpace;
            internal ulong PairedLineFeeds;
            internal ulong InnerLineSeparators;
            internal ulong High;
        }

        private static bool CanCompressTrimmedBlocks =>
            Avx512Vbmi2.IsSupported && Avx512Vbmi.IsSupported && Bmi2.X64.IsSupported && Vector512.IsHardwareAccelerated;

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private void LoadTrimmedBlockMasks(ref byte data, int position, ref ulong insideQuotes, ref ulong carriageReturnBefore, out TrimmedBlockMasks masks)
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

            ulong pairedLineFeeds = lineFeeds & ((carriageReturns << 1) | carriageReturnBefore) & ~inside;
            carriageReturnBefore = (carriageReturns & ~inside) >> 63;
            masks.Structural = separators & ~inside & ~pairedLineFeeds;
            masks.RecordEnds = lineSeparators & ~inside & ~pairedLineFeeds;
            masks.Quotes = quotes;
            masks.Inside = inside;
            masks.WhiteSpace = LoadWhiteSpaceMask(ref data, position) & ~inside & ~separators;
            masks.PairedLineFeeds = pairedLineFeeds;
            masks.InnerLineSeparators = lineSeparators & inside;
            masks.High = high;
        }

        /// <summary>
        /// <see cref="IndexTrimmedBlocks"/> with whole-block operations. The whitespace a field loses is
        /// found with masks: leading whitespace by a carry from each field start, trailing whitespace by
        /// filling back from the byte before each separator. The kept bytes are packed into the trimmed
        /// text in one instruction, and the separators' places in that text are the index entries. A
        /// quote must open a field's text, close it, or be part of an escaped pair. The last block of the
        /// window is read only to settle whitespace and quotes at the end of the block before it.
        /// </summary>
        private int IndexTrimmedBlocksCompressed(int[] ends, int[] rowEnds, int blocks)
        {
            ref byte data = ref MemoryMarshal.GetArrayDataReference(_buffer);
            ref int entries = ref MemoryMarshal.GetArrayDataReference(ends);
            ref int rowSlots = ref MemoryMarshal.GetArrayDataReference(rowEnds);
            ref byte text = ref MemoryMarshal.GetArrayDataReference(_trimmedText ??= ArrayPool<byte>.Shared.Rent(TrimmedTextCapacity));
            Vector512<byte> offsets = Vector512<byte>.Indices;
            bool strictQuotes = _strictQuotes;
            int position = _position;
            if (StartsSkippedComment(position))
            {
                return 0;
            }

            ulong insideQuotes = 0;
            ulong carriageReturnBefore = 0;
            LoadTrimmedBlockMasks(ref data, position, ref insideQuotes, ref carriageReturnBefore, out TrimmedBlockMasks current);
            // The batch start begins a field, as a separator in bit 63 or whitespace reaching it does for the next block.
            ulong startCarry = 1;
            ulong closingBefore = 0;
            ulong leadingBefore = 0;
            ulong nonAscii = 0;
            bool escapedQuotes = false;
            int written = 0;
            int entry = 1;
            int rows = 0;
            entries = -1;
            for (int block = 0; block < blocks - 1; block++, position += BlockSize)
            {
                LoadTrimmedBlockMasks(ref data, position + BlockSize, ref insideQuotes, ref carriageReturnBefore, out TrimmedBlockMasks following);
                ulong structural = current.Structural;
                ulong white = current.WhiteSpace;

                // Leading whitespace, and the line feed of a CRLF pair before a record: the carry from each field
                // start runs through it to the field's first kept byte.
                ulong skippable = white | current.PairedLineFeeds;
                ulong starts = (structural << 1) | startCarry;
                ulong carried = (starts & skippable) + skippable;
                ulong textStarts = (starts | carried) & ~skippable;
                ulong leading = skippable & ~carried;
                startCarry = (structural >> 63) | (carried < skippable ? 1UL : 0UL);

                // Trailing whitespace: runs that end at a separator. Whether the next kept byte after bit 63 is a
                // separator comes from the following block.
                int followingFirstKept = BitOperations.TrailingZeroCount(~following.WhiteSpace);
                ulong separatorFollows = followingFirstKept < BlockSize ? (following.Structural >> followingFirstKept) & 1UL : 0UL;
                ulong trailing = ((structural >> 1) | (separatorFollows << 63)) & white;
                ulong propagate = white;
                trailing |= (trailing >> 1) & propagate;
                propagate &= propagate >> 1;
                trailing |= (trailing >> 2) & propagate;
                propagate &= propagate >> 2;
                trailing |= (trailing >> 4) & propagate;
                propagate &= propagate >> 4;
                trailing |= (trailing >> 8) & propagate;
                propagate &= propagate >> 8;
                trailing |= (trailing >> 16) & propagate;
                propagate &= propagate >> 16;
                trailing |= (trailing >> 32) & propagate;
                ulong kept = ~(leading | trailing);
                ulong textEnds = kept & ~structural & (((trailing | structural) >> 1) | (separatorFollows << 63));

                ulong quotes = current.Quotes;
                ulong opening = quotes & current.Inside;
                ulong closing = quotes & ~current.Inside;
                ulong escapedSecond = opening & ((closing << 1) | closingBefore);
                ulong escapedFirst = closing & ((quotes >> 1) | ((following.Quotes & 1UL) << 63));
                ulong violations = (opening & ~textStarts & ~escapedSecond) | (closing & ~textEnds & ~escapedFirst);
                if (strictQuotes)
                {
                    violations |= opening & ~escapedSecond & (((leading & white) << 1) | leadingBefore);
                }

                if (followingFirstKept == BlockSize)
                {
                    // A whole block of whitespace leaves the end of this one undecided; the record scanner reads it.
                    violations |= 1UL << 63;
                }

                ulong edges = (textStarts | textEnds) & current.High;
                while (edges != 0)
                {
                    int bit = BitOperations.TrailingZeroCount(edges);
                    if (IsNonAsciiWhiteSpaceEdge(position + bit, ((textStarts >> bit) & 1UL) != 0, ((textEnds >> bit) & 1UL) != 0))
                    {
                        violations |= 1UL << bit;
                    }

                    edges &= edges - 1UL;
                }

                escapedQuotes |= escapedSecond != 0;
                closingBefore = closing >> 63;
                leadingBefore = (leading & white) >> 63;
                nonAscii |= current.High;
                ulong recordEnds = current.RecordEnds;
                ulong stops = violations | current.InnerLineSeparators;
                if (stops != 0)
                {
                    // Keep the records that end before the first byte the batch cannot take.
                    ulong endsBefore = recordEnds & ((1UL << BitOperations.TrailingZeroCount(stops)) - 1UL);
                    structural &= endsBefore == 0 ? 0 : ulong.MaxValue >> BitOperations.LeadingZeroCount(endsBefore);
                    recordEnds &= structural;
                }

                Avx512Vbmi2.Compress(Vector512<byte>.Zero, ExpandMask(kept), Vector512.LoadUnsafe(ref data, (nuint)position))
                    .StoreUnsafe(ref text, (nuint)written);
                Vector512<byte> packed = Avx512Vbmi2.Compress(Vector512<byte>.Zero,
                    ExpandMask(Bmi2.X64.ParallelBitExtract(structural, kept)), offsets);
                Vector512<int> textStart = Vector512.Create(written);
                int found = BitOperations.PopCount(structural);
                ref int slot = ref Unsafe.Add(ref entries, entry);
                (Avx512F.ConvertToVector512Int32(packed.GetLower().GetLower()) + textStart).StoreUnsafe(ref slot);
                if (found > 16)
                {
                    (Avx512F.ConvertToVector512Int32(packed.GetLower().GetUpper()) + textStart).StoreUnsafe(ref slot, 16);
                    if (found > 32)
                    {
                        (Avx512F.ConvertToVector512Int32(packed.GetUpper().GetLower()) + textStart).StoreUnsafe(ref slot, 32);
                        (Avx512F.ConvertToVector512Int32(packed.GetUpper().GetUpper()) + textStart).StoreUnsafe(ref slot, 48);
                    }
                }

                while (recordEnds != 0)
                {
                    int bit = BitOperations.TrailingZeroCount(recordEnds);
                    ulong pairedAfter = bit < 63 ? (current.PairedLineFeeds >> (bit + 1)) & 1UL : following.PairedLineFeeds & 1UL;
                    int nextRecord = position + bit + 1 + (int)pairedAfter;
                    ref int row = ref Unsafe.Add(ref rowSlots, 2 * rows);
                    row = entry + BitOperations.PopCount(structural & ((1UL << bit) - 1UL));
                    Unsafe.Add(ref row, 1) = nextRecord;
                    rows++;
                    if (StartsSkippedComment(nextRecord))
                    {
                        return CompleteTrimmedIndex(rows, nonAscii, escapedQuotes);
                    }

                    recordEnds &= recordEnds - 1UL;
                }

                if (stops != 0)
                {
                    break;
                }

                written += BitOperations.PopCount(kept);
                entry += found;
                current = following;
            }

            return CompleteTrimmedIndex(rows, nonAscii, escapedQuotes);
        }

        /// <summary>
        /// Whether the non-ASCII character that starts trimmed text at <paramref name="index"/>, or ends it
        /// with the byte at <paramref name="index"/>, is whitespace that only decoded text can trim.
        /// </summary>
        [MethodImpl(MethodImplOptions.NoInlining)]
        private bool IsNonAsciiWhiteSpaceEdge(int index, bool textStart, bool textEnd)
        {
            ReadOnlySpan<byte> bytes = _buffer.AsSpan(0, _length);
            if (textStart && Rune.DecodeFromUtf8(bytes.Slice(index), out Rune first, out _) == OperationStatus.Done && Rune.IsWhiteSpace(first))
            {
                return true;
            }

            int from = Math.Max(0, index - 3);
            return textEnd && Rune.DecodeLastFromUtf8(bytes.Slice(from, index + 1 - from), out Rune last, out _) == OperationStatus.Done &&
                Rune.IsWhiteSpace(last);
        }
    }
}
#endif
