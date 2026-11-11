#if NET8_0_OR_GREATER
using System.Buffers.Binary;
using System.Numerics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

namespace OfficeIMO.Excel;

/// <summary>Raw deflate (RFC 1951) for <see cref="System.IO.Compression.CompressionLevel.Fastest"/> package parts.</summary>
/// <remarks>Input is collected in regions behind a 32 KiB history: the first fills the 160 KiB buffer, later ones add 128 KiB. Each region becomes one block with the
/// fixed Huffman codes (section 3.2.6), or stored blocks when those would be smaller. Matches come from a single
/// hash slot per four-byte prefix. Codes are written least-significant bit first, so the encoder relies on a
/// little-endian host.</remarks>
internal sealed class ExcelDeflateEncoder : IDisposable {
    private const int HistoryLength = 32 * 1024;
    private const int RegionLength = 128 * 1024;
    private const int BufferLength = HistoryLength + RegionLength;
    private const int HashBits = 15;
    private const int MinimumMatch = 4;
    private const int MaximumMatch = 258;
    private const int StoredBlockLimit = 65_535;
    // A region can hold the whole buffer before any history exists. Fixed codes spend at most 9 bits per byte and
    // stored output adds 5 bytes per block; the headroom also covers the 8-byte stores past the last whole byte.
    private const int OutputLength = BufferLength / 8 * 9 + 256;

    private static Workspace? _cachedWorkspace;
    private readonly Stream _destination;
    private byte[] _buffer;
    private byte[] _output;
    private uint[] _heads;
    private int _filled;
    private int _pending;
    private int _bufferUsed;
    private uint _bufferOrigin;
    private int _outputCount;
    private ulong _bits;
    private int _bitCount;
    private bool _finished;

    internal ExcelDeflateEncoder(Stream destination) {
        _destination = destination;
        // The buffers exceed the large-object threshold, so one cleared set is kept for the next part
        // instead of going through the shared pool, which would drop them with a short-lived worker thread.
        Workspace workspace = Interlocked.Exchange(ref _cachedWorkspace, null) ?? new Workspace();
        _buffer = workspace.Buffer;
        _output = workspace.Output;
        _heads = workspace.Heads;
    }

    internal void Write(ReadOnlySpan<byte> input) {
        if (_finished || _buffer.Length == 0) throw new ObjectDisposedException(nameof(ExcelDeflateEncoder));
        while (!input.IsEmpty) {
            int count = Math.Min(input.Length, BufferLength - _filled);
            input.Slice(0, count).CopyTo(_buffer.AsSpan(_filled));
            _filled += count;
            if (_filled > _bufferUsed) _bufferUsed = _filled;
            input = input.Slice(count);
            if (_filled == BufferLength) {
                EncodeRegion(final: false);
                // Keep the last 32 KiB as history for the next region.
                Buffer.BlockCopy(_buffer, BufferLength - HistoryLength, _buffer, 0, HistoryLength);
                _bufferOrigin += BufferLength - HistoryLength;
                _filled = HistoryLength;
                _pending = HistoryLength;
            }
        }
    }

    /// <summary>Encodes the remaining input as the final block and pads the last byte.</summary>
    internal void Finish() {
        if (_finished || _buffer.Length == 0) throw new ObjectDisposedException(nameof(ExcelDeflateEncoder));
        _finished = true;
        EncodeRegion(final: true);
        while (_bitCount > 0) {
            _output[_outputCount++] = (byte)_bits;
            _bits >>= 8;
            _bitCount = Math.Max(0, _bitCount - 8);
        }
        Drain();
    }

    public void Dispose() {
        if (_buffer.Length == 0) return;
        // Both buffers held caller data; clear what was used before they return to the shared pool.
        Array.Clear(_buffer, 0, _bufferUsed);
        Array.Clear(_output, 0, OutputLength);
        Array.Clear(_heads, 0, _heads.Length);
        Interlocked.CompareExchange(ref _cachedWorkspace, new Workspace(_buffer, _output, _heads), null);
        _buffer = Array.Empty<byte>();
        _output = Array.Empty<byte>();
        _heads = Array.Empty<uint>();
    }

    private void Drain() {
        if (_outputCount == 0) return;
        _destination.Write(_output, 0, _outputCount);
        _outputCount = 0;
    }

    private void EncodeRegion(bool final) {
        int start = _pending, end = _filled;
        int savedCount = _outputCount;
        ulong savedBits = _bits;
        int savedBitCount = _bitCount;
        PutBits(final ? 0b011UL : 0b010UL, 3); // BFINAL, then BTYPE 01 (fixed codes) least-significant bit first
        EncodeSymbols(start, end);
        PutBits(0, 7); // end of block: symbol 256 has the 7-bit fixed code 0000000
        long fixedBits = (long)(_outputCount - savedCount) * 8 + _bitCount - savedBitCount;
        if (fixedBits > StoredBits(end - start, savedBitCount)) {
            _outputCount = savedCount;
            _bits = savedBits;
            _bitCount = savedBitCount;
            WriteStored(start, end, final);
        }
        _pending = end;
        Drain();
    }

    private static long StoredBits(int length, int bitCount) {
        int blocks = Math.Max(1, (length + StoredBlockLimit - 1) / StoredBlockLimit);
        // The first header is padded from the current bit position; later headers start on a byte boundary.
        int firstPadding = (8 - ((bitCount + 3) & 7)) & 7;
        return 8L * length + blocks * (3L + 32) + firstPadding + (blocks - 1) * 5L;
    }

    private void WriteStored(int start, int end, bool final) {
        int position = start;
        do {
            int length = Math.Min(StoredBlockLimit, end - position);
            bool last = position + length == end;
            PutBits(final && last ? 1UL : 0UL, 3); // BTYPE 00
            // Pad to a byte boundary and move every whole byte into the output.
            _bitCount = (_bitCount + 7) & ~7;
            while (_bitCount > 0) {
                _output[_outputCount++] = (byte)_bits;
                _bits >>= 8;
                _bitCount -= 8;
            }
            _bits = 0;
            BinaryPrimitives.WriteUInt16LittleEndian(_output.AsSpan(_outputCount), (ushort)length);
            BinaryPrimitives.WriteUInt16LittleEndian(_output.AsSpan(_outputCount + 2), (ushort)~length);
            _outputCount += 4;
            Buffer.BlockCopy(_buffer, position, _output, _outputCount, length);
            _outputCount += length;
            position += length;
        } while (position < end);
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private void PutBits(ulong value, int count) {
        _bits |= value << _bitCount;
        _bitCount += count;
        if (_bitCount >= 32) {
            BinaryPrimitives.WriteUInt32LittleEndian(_output.AsSpan(_outputCount), (uint)_bits);
            _outputCount += 4;
            _bits >>= 32;
            _bitCount -= 32;
        }
    }

    // A long-running loop per region; skip tiering so the first regions of a process also run optimized code.
    [MethodImpl(MethodImplOptions.AggressiveOptimization)]
    private void EncodeSymbols(int start, int end) {
        ref byte data = ref MemoryMarshal.GetArrayDataReference(_buffer);
        ref uint heads = ref MemoryMarshal.GetArrayDataReference(_heads);
        ref uint literals = ref MemoryMarshal.GetArrayDataReference(FixedCodes.Literals);
        ref uint lengths = ref MemoryMarshal.GetArrayDataReference(FixedCodes.Lengths);
        ref uint nearDistances = ref MemoryMarshal.GetArrayDataReference(FixedCodes.NearDistances);
        ref byte farSymbols = ref MemoryMarshal.GetArrayDataReference(FixedCodes.FarDistanceSymbols);
        ref uint symbolCodes = ref MemoryMarshal.GetArrayDataReference(FixedCodes.DistanceSymbolCodes);
        ref int symbolBases = ref MemoryMarshal.GetArrayDataReference(FixedCodes.DistanceBases);
        byte[] output = _output;
        if (output.Length < OutputLength || end > _buffer.Length || (long)(end - start) * 9 / 8 + _outputCount + 64 > output.Length)
            throw new InvalidOperationException("The deflate buffers are invalid.");
        ref byte outputStart = ref MemoryMarshal.GetArrayDataReference(output);
        ulong bits = _bits;
        int bitCount = _bitCount;
        int written = _outputCount;
        uint origin = _bufferOrigin;
        int index = start;
        int lastMatchStart = end - MinimumMatch;
        while (index <= lastMatchStart) {
            uint prefix = Unsafe.ReadUnaligned<uint>(ref Unsafe.Add(ref data, index));
            ref uint head = ref Unsafe.Add(ref heads, (int)((prefix * 0x9E37_79B1U) >> (32 - HashBits)));
            uint here = origin + (uint)index;
            uint distance = here - head;
            head = here;
            ulong code;
            int codeBits;
            if (distance - 1 < HistoryLength && distance <= (uint)index
                && Unsafe.ReadUnaligned<uint>(ref Unsafe.Add(ref data, index - (int)distance)) == prefix) {
                int length = ExtendMatch(ref data, index - (int)distance, index, Math.Min(MaximumMatch, end - index));
                uint lengthCode = Unsafe.Add(ref lengths, length);
                int offset = (int)distance - 1;
                uint distanceCode;
                if (offset < 256) {
                    distanceCode = Unsafe.Add(ref nearDistances, offset);
                } else {
                    int symbol = Unsafe.Add(ref farSymbols, offset >> 7);
                    distanceCode = Unsafe.Add(ref symbolCodes, symbol)
                        + ((uint)((int)distance - Unsafe.Add(ref symbolBases, symbol)) << 5);
                }
                int lengthBits = (int)(lengthCode >> 24);
                code = (lengthCode & 0xFF_FFFF) | ((ulong)(distanceCode & 0xFF_FFFF) << lengthBits);
                codeBits = lengthBits + (int)(distanceCode >> 24);
                index += length;
            } else {
                uint literal = Unsafe.Add(ref literals, Unsafe.Add(ref data, index));
                code = literal & 0xFF_FFFF;
                codeBits = (int)(literal >> 24);
                index++;
            }
            // Store eight bytes and keep fewer than eight bits pending: at most 7 + 31 bits are ever live.
            bits |= code << bitCount;
            bitCount += codeBits;
            Unsafe.WriteUnaligned(ref Unsafe.Add(ref outputStart, written), bits);
            written += bitCount >> 3;
            bits >>= bitCount & ~7;
            bitCount &= 7;
        }
        for (; index < end; index++) {
            uint literal = Unsafe.Add(ref literals, Unsafe.Add(ref data, index));
            bits |= (ulong)(literal & 0xFF_FFFF) << bitCount;
            bitCount += (int)(literal >> 24);
            Unsafe.WriteUnaligned(ref Unsafe.Add(ref outputStart, written), bits);
            written += bitCount >> 3;
            bits >>= bitCount & ~7;
            bitCount &= 7;
        }
        _bits = bits;
        _bitCount = bitCount;
        _outputCount = written;
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    private static int ExtendMatch(ref byte data, int earlier, int current, int limit) {
        int length = MinimumMatch;
        while (length + sizeof(ulong) <= limit) {
            ulong difference = Unsafe.ReadUnaligned<ulong>(ref Unsafe.Add(ref data, earlier + length))
                ^ Unsafe.ReadUnaligned<ulong>(ref Unsafe.Add(ref data, current + length));
            if (difference != 0) return length + BitOperations.TrailingZeroCount(difference) / 8;
            length += sizeof(ulong);
        }
        while (length < limit && Unsafe.Add(ref data, earlier + length) == Unsafe.Add(ref data, current + length)) length++;
        return length;
    }

    /// <summary>Encoder buffers; a cached set is always fully cleared.</summary>
    private sealed class Workspace(byte[] buffer, byte[] output, uint[] heads) {
        internal Workspace() : this(new byte[BufferLength], new byte[OutputLength], new uint[1 << HashBits]) { }
        internal byte[] Buffer { get; } = buffer;
        internal byte[] Output { get; } = output;
        internal uint[] Heads { get; } = heads;
    }

    /// <summary>The fixed literal/length and distance codes of RFC 1951 section 3.2.6.</summary>
    /// <remarks>Entries hold the Huffman code with its bits reversed for least-significant-bit-first output,
    /// followed by any extra bits, in the low 24 bits; the high 8 bits hold the total bit count.</remarks>
    private static class FixedCodes {
        internal static readonly uint[] Literals = CreateLiterals();
        /// <summary>Indexed by match length, 3 through 258.</summary>
        internal static readonly uint[] Lengths = CreateLengths();
        /// <summary>Base distance of each of the 30 distance symbols.</summary>
        internal static readonly int[] DistanceBases = new int[30];
        private static readonly byte[] DistanceExtraBits = new byte[30];
        /// <summary>Reversed 5-bit symbol code with the symbol's total bit count; add (distance - base) &lt;&lt; 5.</summary>
        internal static readonly uint[] DistanceSymbolCodes = new uint[30];
        /// <summary>Complete code for distances 1-256, indexed by distance - 1.</summary>
        internal static readonly uint[] NearDistances = new uint[256];
        /// <summary>Symbol for distances above 256, indexed by (distance - 1) >> 7; those symbols have 7+ extra bits.</summary>
        internal static readonly byte[] FarDistanceSymbols = CreateDistanceTables();

        private static byte[] CreateDistanceTables() {
            int distanceBase = 1;
            for (int symbol = 0; symbol < 30; symbol++) {
                DistanceBases[symbol] = distanceBase;
                DistanceExtraBits[symbol] = (byte)(symbol < 4 ? 0 : symbol / 2 - 1);
                DistanceSymbolCodes[symbol] = Reverse((uint)symbol, 5) | ((uint)(5 + DistanceExtraBits[symbol]) << 24);
                distanceBase += 1 << DistanceExtraBits[symbol];
            }
            var far = new byte[256];
            for (int symbol = 0; symbol < 30; symbol++) {
                int first = DistanceBases[symbol], last = first + (1 << DistanceExtraBits[symbol]) - 1;
                for (int distance = first; distance <= last; distance++) {
                    if (distance <= 256) NearDistances[distance - 1] = DistanceSymbolCodes[symbol] + ((uint)(distance - first) << 5);
                    else far[(distance - 1) >> 7] = (byte)symbol;
                }
            }
            return far;
        }

        private static uint[] CreateLiterals() {
            var table = new uint[256];
            for (int value = 0; value < 256; value++) table[value] = LiteralLengthCode(value, 0, 0);
            return table;
        }

        private static uint[] CreateLengths() {
            var table = new uint[MaximumMatch + 1];
            int length = 3;
            // Symbols 257-284 cover four lengths per extra-bit size from 0 to 5 bits; 285 is exactly 258.
            for (int symbol = 257; symbol <= 284; symbol++) {
                int extraBits = symbol < 265 ? 0 : (symbol - 261) / 4;
                for (int extra = 0; extra < 1 << extraBits && length < MaximumMatch; extra++, length++)
                    table[length] = LiteralLengthCode(symbol, extra, extraBits);
            }
            table[MaximumMatch] = LiteralLengthCode(285, 0, 0);
            return table;
        }

        private static uint LiteralLengthCode(int symbol, int extra, int extraBits) {
            (uint code, int bits) = symbol switch {
                < 144 => ((uint)(0x30 + symbol), 8),
                < 256 => ((uint)(0x190 + symbol - 144), 9),
                < 280 => ((uint)(symbol - 256), 7),
                _ => ((uint)(0xC0 + symbol - 280), 8)
            };
            return (Reverse(code, bits) | ((uint)extra << bits)) | ((uint)(bits + extraBits) << 24);
        }

        private static uint Reverse(uint value, int bits) {
            uint result = 0;
            for (int bit = 0; bit < bits; bit++) result |= ((value >> bit) & 1) << (bits - 1 - bit);
            return result;
        }
    }
}
#endif
