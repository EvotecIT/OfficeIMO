using System.Buffers.Binary;
using System.Collections.Concurrent;
#if NET8_0_OR_GREATER
using System.Numerics;
#endif
using System.Runtime.CompilerServices;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using OfficeIMO.Drawing;

namespace OfficeIMO.Excel;

/// <summary>Raw deflate (RFC 1951) for <see cref="System.IO.Compression.CompressionLevel.Fastest"/> package parts.</summary>
/// <remarks>Input is cut into regions at fixed offsets: the first holds up to 160 KiB, each later one 128 KiB after a
/// copy of the preceding 32 KiB as history. A region is encoded on its own, with a match table that starts from its
/// history, as one block with the fixed Huffman codes (section 3.2.6), or stored blocks when those would be smaller.
/// A region other than the last ends with an empty stored block, so the next one starts on a byte boundary.
/// Regions therefore compress independently: with background encoding several are compressed at once on the thread
/// pool and written in order, and the output is the same whichever thread encoded them. Each region also hashes its
/// new bytes, and the encoder joins those values into the CRC-32 of the whole input. Matches come from a single
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
    // stored output adds 5 bytes per block; the headroom also covers the 8-byte stores past the last whole byte and
    // the empty stored block that aligns a region's end.
    private const int OutputLength = BufferLength / 8 * 9 + 256;
    // Background encoding keeps this many full regions in flight while the caller fills the next one.
    private static readonly int BackgroundRegions = Environment.ProcessorCount > 1 ? Math.Min(Environment.ProcessorCount - 1, 4) : 0;
    // Background regions of every encoder go through one queue served by at most this many pool workers, so a busy
    // caller wakes a worker only when one went idle rather than once per region.
    private static readonly ConcurrentQueue<Region> _work = new();
    private static readonly WaitCallback Work = _ => RunWorker();
    private static int _workers;

    private static Region? _cachedRegion;
    // Parts prepared on worker threads, and regions encoded in the background, each hold a region. The extra sets
    // are held weakly, so a burst of parallel work reuses them without the large-object allocations, and a later
    // full collection can still reclaim them. The slots reference the arrays themselves: a weakly held wrapper would
    // die in the next young collection.
    private static readonly WeakRegion?[] _sharedRegions = new WeakRegion?[Math.Min(Environment.ProcessorCount * 2, 64)];
    private readonly Stream _destination;
    private readonly int _maximumQueued;
    private readonly Queue<Region> _queued = new();
    // Regions this part already used; they are cleared once, when the part is done with them.
    private readonly Stack<Region> _spare = new();
    private Region? _current;
    private int _filled;
    private bool _finished;
    private uint _crc;
    private long _size;

    /// <summary>The CRC-32 of every byte written; valid after <see cref="Finish"/>.</summary>
    internal uint Crc => _crc;

    /// <summary>The number of bytes written.</summary>
    internal long Size => _size;

    /// <param name="destination">Receives the compressed bytes, always on the thread that writes or finishes.</param>
    /// <param name="background">Encodes full regions on the thread pool while the caller keeps writing.</param>
    internal ExcelDeflateEncoder(Stream destination, bool background = false) {
        _destination = destination;
        _maximumQueued = background ? BackgroundRegions : 0;
        _current = RentRegion();
    }

    internal void Write(ReadOnlySpan<byte> input) {
        if (_finished || _current == null) throw new ObjectDisposedException(nameof(ExcelDeflateEncoder));
        _size += input.Length;
        while (!input.IsEmpty) {
            Region region = _current;
            int count = Math.Min(input.Length, BufferLength - _filled);
            input.Slice(0, count).CopyTo(region.Buffer.AsSpan(_filled));
            _filled += count;
            input = input.Slice(count);
            if (_filled == BufferLength) {
                // The next region starts from the last 32 KiB of this one.
                Region next = _spare.Count > 0 ? _spare.Pop() : RentRegion();
                Buffer.BlockCopy(region.Buffer, BufferLength - HistoryLength, next.Buffer, 0, HistoryLength);
                next.Start = HistoryLength;
                _current = next;
                _filled = HistoryLength;
                region.End = BufferLength;
                Submit(region);
            }
        }
    }

    /// <summary>Encodes the remaining input as the final block, then writes every region in order.</summary>
    internal void Finish() {
        if (_finished || _current == null) throw new ObjectDisposedException(nameof(ExcelDeflateEncoder));
        _finished = true;
        Region last = _current;
        _current = null;
        last.End = _filled;
        last.Final = true;
        try {
            // The last region is encoded here while earlier ones may still be running in the background.
            last.Encode();
            while (_queued.Count > 0) WriteOldest();
            Emit(last);
        } finally {
            ReturnRegion(last);
            while (_spare.Count > 0) ReturnRegion(_spare.Pop());
        }
    }

    public void Dispose() {
        // Regions still in the background own their arrays until they finish; collect them before returning any.
        while (_queued.Count > 0) {
            Region region = _queued.Dequeue();
            region.Cancel();
            ReturnRegion(region);
        }
        if (_current != null) {
            _current.End = _filled;
            ReturnRegion(_current);
            _current = null;
        }
        while (_spare.Count > 0) ReturnRegion(_spare.Pop());
    }

    private void Submit(Region region) {
        if (_maximumQueued == 0) {
            try {
                region.Encode();
                Emit(region);
            } finally {
                Recycle(region);
            }
            return;
        }
        _queued.Enqueue(region);
        region.MarkQueued();
        Schedule(region);
        // Write what has finished, in order, and keep at most the configured number of regions in flight.
        while (_queued.Count > 0 && (_queued.Count > _maximumQueued || _queued.Peek().IsComplete)) WriteOldest();
    }

    private void WriteOldest() {
        Region region = _queued.Peek();
        region.Join();
        _queued.Dequeue();
        try {
            region.ThrowIfFailed();
            Emit(region);
        } finally {
            Recycle(region);
        }
    }

    private void Emit(Region region) {
        region.WriteTo(_destination);
        _crc = ExcelCrc32Combine.Combine(_crc, region.Crc, region.End - region.Start);
    }

    /// <summary>Keeps a written region for this part's next one; its bytes are overwritten before they are read.</summary>
    private void Recycle(Region region) {
        region.Reset();
        _spare.Push(region);
    }

    private static void Schedule(Region region) {
        _work.Enqueue(region);
        int workers = Volatile.Read(ref _workers);
        while (workers < BackgroundRegions) {
            if (Interlocked.CompareExchange(ref _workers, workers + 1, workers) == workers) {
                try {
                    ThreadPool.UnsafeQueueUserWorkItem(Work, null);
                } catch {
                    // Without the pool the caller encodes its regions when it writes them.
                    Interlocked.Decrement(ref _workers);
                }
                return;
            }
            workers = Volatile.Read(ref _workers);
        }
    }

    private static void RunWorker() {
        while (true) {
            while (_work.TryDequeue(out Region? region)) region.RunIfQueued();
            Interlocked.Decrement(ref _workers);
            // A region queued while this worker was leaving may have found every slot taken; take the slot back for it.
            if (_work.IsEmpty) return;
            int workers = Volatile.Read(ref _workers);
            if (workers >= BackgroundRegions || Interlocked.CompareExchange(ref _workers, workers + 1, workers) != workers) return;
        }
    }

    private static Region RentRegion() {
        Region? region = Interlocked.Exchange(ref _cachedRegion, null);
        if (region != null) return region;
        for (int i = 0; i < _sharedRegions.Length; i++) {
            if (Volatile.Read(ref _sharedRegions[i]) != null
                && Interlocked.Exchange(ref _sharedRegions[i], null)?.TryGetRegion() is { } shared) {
                return shared;
            }
        }
        return new Region();
    }

    private static void ReturnRegion(Region region) {
        // The buffers held caller data; clear what was used before they can serve another part.
        region.Clear();
        if (Interlocked.CompareExchange(ref _cachedRegion, region, null) == null) return;
        for (int i = 0; i < _sharedRegions.Length; i++) {
            if (Volatile.Read(ref _sharedRegions[i]) == null
                && Interlocked.CompareExchange(ref _sharedRegions[i], new WeakRegion(region), null) == null) {
                return;
            }
        }
    }

    /// <summary>One region's input, match table and output, and its background state.</summary>
    private sealed class Region(byte[] buffer, byte[] output, uint[] heads) {
        // A caller that must wait moves Running to RunningWithWaiter; only then does the worker signal the event,
        // and the caller consumes that signal before it can reuse the region.
        private const int Idle = 0, Queued = 1, Running = 2, RunningWithWaiter = 3, Completed = 4;
        private readonly byte[] _buffer = buffer;
        private readonly byte[] _output = output;
        private readonly uint[] _heads = heads;
        private int _state;
        private ManualResetEventSlim? _completed;
        private ExceptionDispatchInfo? _failure;
        private int _outputCount;
        private uint _crc;
        private int _used;
        private int _outputUsed;
        private ulong _bits;
        private int _bitCount;

        internal Region() : this(new byte[BufferLength], new byte[OutputLength], new uint[1 << HashBits]) { }

        internal byte[] Buffer => _buffer;
        internal byte[] Output => _output;
        internal uint[] Heads => _heads;
        /// <summary>The first byte to encode; the bytes before it are history.</summary>
        internal int Start { get; set; }
        internal int End { get; set; }
        internal bool Final { get; set; }
        internal bool IsComplete => Volatile.Read(ref _state) == Completed;
        /// <summary>The CRC-32 of the region's new bytes, after it was encoded.</summary>
        internal uint Crc => _crc;

        internal void MarkQueued() => Volatile.Write(ref _state, Queued);

        /// <summary>Waits for a background encoding, or encodes here when no worker has started it.</summary>
        internal void Join() {
            if (Interlocked.CompareExchange(ref _state, Running, Queued) == Queued) {
                RunGuarded();
                return;
            }
            if (Volatile.Read(ref _state) == Completed) return;
            _completed ??= new ManualResetEventSlim(false);
            if (Interlocked.CompareExchange(ref _state, RunningWithWaiter, Running) == Running) _completed.Wait();
        }

        /// <summary>Drops a region nobody started, or waits for the worker that did.</summary>
        internal void Cancel() {
            if (Interlocked.CompareExchange(ref _state, Completed, Queued) != Queued) Join();
        }

        internal void ThrowIfFailed() => _failure?.Throw();

        internal void RunIfQueued() {
            if (Interlocked.CompareExchange(ref _state, Running, Queued) == Queued) RunGuarded();
        }

        private void RunGuarded() {
            try {
                Encode();
            } catch (Exception exception) {
                _failure = ExceptionDispatchInfo.Capture(exception);
            }
            if (Interlocked.Exchange(ref _state, Completed) == RunningWithWaiter) _completed!.Set();
        }

        internal void WriteTo(Stream destination) {
            if (_outputCount != 0) destination.Write(_output, 0, _outputCount);
        }

        internal void Clear() {
            Reset();
            Array.Clear(_buffer, 0, _used);
            Array.Clear(_output, 0, _outputUsed);
            _used = 0;
            _outputUsed = 0;
        }

        /// <summary>Readies the region for more input of the same part.</summary>
        internal void Reset() {
            _used = Math.Max(_used, End);
            _outputUsed = Math.Max(_outputUsed, _outputCount);
            _completed?.Reset();
            _state = Idle;
            _failure = null;
            _outputCount = 0;
            _bits = 0;
            _bitCount = 0;
            Start = 0;
            End = 0;
            Final = false;
        }

        internal void Encode() {
            int start = Start, end = End;
            _crc = OfficePngCrc32.Complete(OfficePngCrc32.Append(OfficePngCrc32.Begin(), _buffer, start, end - start));
            _outputCount = 0;
            _bits = 0;
            _bitCount = 0;
            PrimeHistory(start);
            PutBits(Final ? 0b011UL : 0b010UL, 3); // BFINAL, then BTYPE 01 (fixed codes) least-significant bit first
            EncodeSymbols(start, end);
            PutBits(0, 7); // end of block: symbol 256 has the 7-bit fixed code 0000000
            // A non-final region also spends an empty stored block (3 bits, padding, 4 bytes) to end on a byte boundary.
            long fixedBits = (long)_outputCount * 8 + _bitCount + (Final ? 0 : 3 + 7 + 32);
            if (fixedBits > StoredBits(end - start)) {
                _outputCount = 0;
                _bits = 0;
                _bitCount = 0;
                WriteStored(start, end, Final);
                return;
            }
            if (!Final) {
                PutBits(0, 3); // BFINAL 0, BTYPE 00
                FlushBits();
                BinaryPrimitives.WriteUInt32LittleEndian(_output.AsSpan(_outputCount), 0xFFFF_0000U); // LEN 0, NLEN 0xFFFF
                _outputCount += 4;
            } else {
                FlushBits();
            }
        }

        /// <summary>Starts the match table from the history, so matches can reach into the previous region.</summary>
        private void PrimeHistory(int start) {
            Array.Clear(_heads, 0, _heads.Length);
            ref byte data = ref ArrayStart(_buffer);
            ref uint heads = ref ArrayStart(_heads);
            for (int index = 0; index + MinimumMatch <= start; index++) {
                uint prefix = Unsafe.ReadUnaligned<uint>(ref Unsafe.Add(ref data, index));
                Unsafe.Add(ref heads, (int)((prefix * 0x9E37_79B1U) >> (32 - HashBits))) = (uint)index;
            }
        }

        private static long StoredBits(int length) {
            int blocks = Math.Max(1, (length + StoredBlockLimit - 1) / StoredBlockLimit);
            // Each block spends a 3-bit header padded to a byte, then LEN and NLEN.
            return 8L * length + blocks * (8L + 32);
        }

        private void WriteStored(int start, int end, bool final) {
            int position = start;
            do {
                int length = Math.Min(StoredBlockLimit, end - position);
                bool last = position + length == end;
                PutBits(final && last ? 1UL : 0UL, 3); // BTYPE 00
                FlushBits();
                BinaryPrimitives.WriteUInt16LittleEndian(_output.AsSpan(_outputCount), (ushort)length);
                BinaryPrimitives.WriteUInt16LittleEndian(_output.AsSpan(_outputCount + 2), (ushort)~length);
                _outputCount += 4;
                System.Buffer.BlockCopy(_buffer, position, _output, _outputCount, length);
                _outputCount += length;
                position += length;
            } while (position < end);
        }

        /// <summary>Pads to a byte boundary and moves every pending bit into the output.</summary>
        private void FlushBits() {
            _bitCount = (_bitCount + 7) & ~7;
            while (_bitCount > 0) {
                _output[_outputCount++] = (byte)_bits;
                _bits >>= 8;
                _bitCount -= 8;
            }
            _bits = 0;
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
#if NET8_0_OR_GREATER
        [MethodImpl(MethodImplOptions.AggressiveOptimization)]
#endif
        private void EncodeSymbols(int start, int end) {
            ref byte data = ref ArrayStart(_buffer);
            ref uint heads = ref ArrayStart(_heads);
            ref uint literals = ref ArrayStart(FixedCodes.Literals);
            ref uint lengths = ref ArrayStart(FixedCodes.Lengths);
            ref uint nearDistances = ref ArrayStart(FixedCodes.NearDistances);
            ref byte farSymbols = ref ArrayStart(FixedCodes.FarDistanceSymbols);
            ref uint symbolCodes = ref ArrayStart(FixedCodes.DistanceSymbolCodes);
            ref int symbolBases = ref ArrayStart(FixedCodes.DistanceBases);
            byte[] output = _output;
            if (output.Length < OutputLength || end > _buffer.Length || (long)(end - start) * 9 / 8 + _outputCount + 64 > output.Length)
                throw new InvalidOperationException("The deflate buffers are invalid.");
            ref byte outputStart = ref ArrayStart(output);
            ulong bits = _bits;
            int bitCount = _bitCount;
            int written = _outputCount;
            int index = start;
            int lastMatchStart = end - MinimumMatch;
            while (index <= lastMatchStart) {
                uint prefix = Unsafe.ReadUnaligned<uint>(ref Unsafe.Add(ref data, index));
                ref uint head = ref Unsafe.Add(ref heads, (int)((prefix * 0x9E37_79B1U) >> (32 - HashBits)));
                uint here = (uint)index;
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
                if (difference != 0) return length + TrailingZeroBytes(difference);
                length += sizeof(ulong);
            }
            while (length < limit && Unsafe.Add(ref data, earlier + length) == Unsafe.Add(ref data, current + length)) length++;
            return length;
        }

        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static ref T ArrayStart<T>(T[] array) {
#if NET8_0_OR_GREATER
            return ref MemoryMarshal.GetArrayDataReference(array);
#else
            // Every encoder table and buffer is non-empty.
            return ref array[0];
#endif
        }

        /// <summary>The number of equal low-order bytes before the first difference; <paramref name="difference"/> is not zero.</summary>
        [MethodImpl(MethodImplOptions.AggressiveInlining)]
        private static int TrailingZeroBytes(ulong difference) {
#if NET8_0_OR_GREATER
            return BitOperations.TrailingZeroCount(difference) / 8;
#else
            int bytes = 0;
            if ((uint)difference == 0) { bytes = 4; difference >>= 32; }
            if ((ushort)difference == 0) { bytes += 2; difference >>= 16; }
            if ((byte)difference == 0) bytes++;
            return bytes;
#endif
        }
    }

    private sealed class WeakRegion(Region region) {
        private readonly WeakReference<byte[]> _buffer = new(region.Buffer);
        private readonly WeakReference<byte[]> _output = new(region.Output);
        private readonly WeakReference<uint[]> _heads = new(region.Heads);

        /// <returns>The arrays as a region, or null when a collection has reclaimed any of them.</returns>
        internal Region? TryGetRegion() =>
            _buffer.TryGetTarget(out byte[]? buffer) && _output.TryGetTarget(out byte[]? output) && _heads.TryGetTarget(out uint[]? heads)
                ? new Region(buffer, output, heads)
                : null;
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
