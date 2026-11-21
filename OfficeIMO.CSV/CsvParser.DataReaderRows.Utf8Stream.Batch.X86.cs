#if NET8_0_OR_GREATER
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.Intrinsics;
using System.Runtime.Intrinsics.X86;

namespace OfficeIMO.CSV;

internal static partial class CsvParser
{
    internal sealed partial class CsvUtf8StreamDataReaderRowSource
    {
        /// <summary>
        /// Entry <c>m</c> lists the positions of the set bits of the byte <c>m</c>, lowest first, one per
        /// byte; bytes past the bit count are zero.
        /// </summary>
        private static readonly ulong[] s_setBitPositions = CreateSetBitPositions();

        private static ulong[] CreateSetBitPositions()
        {
            ulong[] table = new ulong[256];
            for (int mask = 1; mask < table.Length; mask++)
            {
                ulong entry = 0;
                int rank = 0;
                for (int bit = 0; bit < 8; bit++)
                {
                    if ((mask >> bit & 1) != 0)
                    {
                        entry |= (ulong)bit << (8 * rank++);
                    }
                }

                table[mask] = entry;
            }

            return table;
        }

        /// <summary>
        /// <see cref="IndexBlocks"/> for AVX2. Offsets are written eight mask bits at a time: a table gives
        /// the positions of a byte's set bits, which are widened, offset and stored as eight entries, and the
        /// next byte's entries start after its set bits. Each block takes a fixed sequence of instructions
        /// instead of a step per separator.
        /// </summary>
        private int IndexBlocksByTable(int[] ends, int[] rowEnds, int blocks)
        {
            ref byte data = ref MemoryMarshal.GetArrayDataReference(_buffer);
            ref int destination = ref MemoryMarshal.GetArrayDataReference(ends);
            ref int rowEndSlots = ref MemoryMarshal.GetArrayDataReference(rowEnds);
            ref ulong table = ref MemoryMarshal.GetArrayDataReference(s_setBitPositions);
            Vector256<byte> delimiterValue = Vector256.Create(_delimiter);
            Vector256<byte> quoteValue = Vector256.Create((byte)'"');
            Vector256<byte> carriageReturnValue = Vector256.Create((byte)'\r');
            Vector256<byte> lineFeedValue = Vector256.Create((byte)'\n');
            Vector256<byte> nonAscii = Vector256<byte>.Zero;
            ulong quoteMask = _quoteMask;
            BlockQuoteState state = default;
            state.OpeningQuoteAllowed = 1;
            int count = 1;
            int rows = 0;
            int position = _position;
            for (int block = 0; block < blocks; block++, position += BlockSize)
            {
                Vector256<byte> low = Vector256.LoadUnsafe(ref data, (nuint)position);
                Vector256<byte> high = Vector256.LoadUnsafe(ref data, (nuint)(position + 32));
                nonAscii |= low | high;
                Vector256<byte> lowLines = Vector256.Equals(low, carriageReturnValue) | Vector256.Equals(low, lineFeedValue);
                Vector256<byte> highLines = Vector256.Equals(high, carriageReturnValue) | Vector256.Equals(high, lineFeedValue);
                Vector256<byte> lowDelimiters = Vector256.Equals(low, delimiterValue);
                Vector256<byte> highDelimiters = Vector256.Equals(high, delimiterValue);
                ulong lineSeparators = Combine(lowLines, highLines);
                ulong quotes = Combine(Vector256.Equals(low, quoteValue), Vector256.Equals(high, quoteValue)) & quoteMask;
                ulong structural;
                bool stop = false;
                if ((quotes | state.InsideQuotes) == 0 && !state.ClosingQuotePending)
                {
                    structural = Combine(lowDelimiters | lowLines, highDelimiters | highLines);
                    state.OpeningQuoteAllowed = structural >> 63;
                }
                else
                {
                    structural = ClassifyQuotedBlock(Combine(lowDelimiters, highDelimiters), quotes, lineSeparators, ref state, out stop);
                }

                ref int slot = ref Unsafe.Add(ref destination, count);
                Vector256<int> blockStart = Vector256.Create(position);
                int written = WriteByteOffsets(ref slot, ref table, blockStart, (uint)structural & 0xFF, 0, 0);
                written = WriteByteOffsets(ref slot, ref table, blockStart, (uint)(structural >> 8) & 0xFF, 8, written);
                written = WriteByteOffsets(ref slot, ref table, blockStart, (uint)(structural >> 16) & 0xFF, 16, written);
                written = WriteByteOffsets(ref slot, ref table, blockStart, (uint)(structural >> 24) & 0xFF, 24, written);
                written = WriteByteOffsets(ref slot, ref table, blockStart, (uint)(structural >> 32) & 0xFF, 32, written);
                written = WriteByteOffsets(ref slot, ref table, blockStart, (uint)(structural >> 40) & 0xFF, 40, written);
                written = WriteByteOffsets(ref slot, ref table, blockStart, (uint)(structural >> 48) & 0xFF, 48, written);
                WriteByteOffsets(ref slot, ref table, blockStart, (uint)(structural >> 56), 56, written);
                rows = AppendRowEnds(ref rowEndSlots, rows, count, structural, lineSeparators);
                count += BitOperations.PopCount(structural);
                if (stop)
                {
                    break;
                }
            }

            _batchAscii = nonAscii.ExtractMostSignificantBits() == 0;
            _batchEscapedQuotes = state.EscapedQuotes;
            return rows;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static ulong Combine(Vector256<byte> low, Vector256<byte> high) =>
            low.ExtractMostSignificantBits() | ((ulong)high.ExtractMostSignificantBits() << 32);

        /// <summary>
        /// Stores the offsets of the set bits of <paramref name="bits"/>, the mask byte that starts at
        /// <paramref name="byteStart"/> in the block, from entry <paramref name="written"/>. Eight entries
        /// are always stored; the next byte overwrites those past the set bits. Returns the entries used.
        /// </summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static int WriteByteOffsets(ref int slot, ref ulong table, Vector256<int> blockStart, uint bits, int byteStart, int written)
        {
            Vector256<int> offsets = Avx2.ConvertToVector256Int32(Vector128.CreateScalarUnsafe(Unsafe.Add(ref table, bits)).AsByte());
            (offsets + (blockStart + Vector256.Create(byteStart))).StoreUnsafe(ref Unsafe.Add(ref slot, written));
            return written + BitOperations.PopCount(bits);
        }

#if NET9_0_OR_GREATER
        /// <summary>
        /// <see cref="IndexBlocks"/> for AVX-512 without VBMI2: each sixteen mask bits select their offsets
        /// from a vector of sixteen block offsets with one 32-bit compress, stored as sixteen entries that
        /// the next group overwrites past its set bits.
        /// </summary>
        private int IndexBlocksDwordCompressed(int[] ends, int[] rowEnds, int blocks)
        {
            ref byte data = ref MemoryMarshal.GetArrayDataReference(_buffer);
            ref int destination = ref MemoryMarshal.GetArrayDataReference(ends);
            ref int rowEndSlots = ref MemoryMarshal.GetArrayDataReference(rowEnds);
            Vector512<byte> delimiterValue = Vector512.Create(_delimiter);
            Vector512<byte> quoteValue = Vector512.Create((byte)'"');
            Vector512<byte> carriageReturnValue = Vector512.Create((byte)'\r');
            Vector512<byte> lineFeedValue = Vector512.Create((byte)'\n');
            Vector512<int> bitValues = Vector512.Create(1, 2, 4, 8, 16, 32, 64, 128, 256, 512, 1024, 2048, 4096, 8192, 16384, 32768);
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
                ulong structural;
                bool stop = false;
                if ((quotes | state.InsideQuotes) == 0 && !state.ClosingQuotePending)
                {
                    structural = (delimiterBytes | lineSeparatorBytes).ExtractMostSignificantBits();
                    state.OpeningQuoteAllowed = structural >> 63;
                }
                else
                {
                    structural = ClassifyQuotedBlock(delimiterBytes.ExtractMostSignificantBits(), quotes, lineSeparators, ref state, out stop);
                }

                ref int slot = ref Unsafe.Add(ref destination, count);
                Vector512<int> offsets = Vector512<int>.Indices + Vector512.Create(position);
                int written = 0;
                for (int group = 0; group < 4; group++)
                {
                    int bits = (int)(structural >> (16 * group)) & 0xFFFF;
                    Vector512<int> selected = Vector512.Equals(Vector512.Create(bits) & bitValues, bitValues);
                    Avx512F.Compress(Vector512<int>.Zero, selected, offsets).StoreUnsafe(ref Unsafe.Add(ref slot, written));
                    written += BitOperations.PopCount((uint)bits);
                    offsets += Vector512.Create(16);
                }

                rows = AppendRowEnds(ref rowEndSlots, rows, count, structural, lineSeparators);
                count += written;
                if (stop)
                {
                    break;
                }
            }

            _batchAscii = nonAscii.ExtractMostSignificantBits() == 0;
            _batchEscapedQuotes = state.EscapedQuotes;
            return rows;
        }
#endif
    }
}
#endif
