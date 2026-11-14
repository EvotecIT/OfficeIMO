namespace OfficeIMO.Archives;

/// <summary>Raw Deflate (RFC 1951) encoder whose output depends only on the input bytes.</summary>
/// <remarks>The runtime DeflateStream links whichever zlib the platform ships, so its bytes differ between .NET
/// versions and operating systems. This encoder uses fixed parameters and pure managed arithmetic: hash chains with
/// lazy matching, and for every block the smallest of dynamic Huffman, fixed Huffman and stored encodings.
/// The result is identical for every target framework, platform and write chunking.</remarks>
internal sealed class ArchiveDeflateEncoder {
    private const int WindowSize = 32768;
    private const int WindowMask = WindowSize - 1;
    private const int HashBits = 15;
    private const int HashSize = 1 << HashBits;
    private const int HashMask = HashSize - 1;
    private const int MinMatch = 3;
    private const int MaxMatch = 258;
    private const int MinLookahead = MaxMatch + MinMatch + 1;
    private const int MaxDistance = WindowSize - MinLookahead;
    private const int MaxChain = 128;
    private const int GoodLength = 8;
    private const int NiceLength = 128;
    private const int MaxLazy = 16;
    private const int TooFar = 4096;
    private const int SymbolCapacity = 16384;
    private const int RawCapacity = 256 * 1024;

    private static readonly int[] LengthBases = {
        3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258
    };
    private static readonly int[] LengthExtraBits = {
        0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0
    };
    private static readonly int[] DistanceBases = {
        1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073,
        4097, 6145, 8193, 12289, 16385, 24577
    };
    private static readonly int[] DistanceExtraBits = {
        0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13
    };
    private static readonly int[] CodeLengthOrder = { 16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15 };
    private static readonly byte[] LengthCodes = CreateLengthCodes();
    private static readonly byte[] DistanceCodes = CreateDistanceCodes();
    private static readonly byte[] FixedLiteralLengths = CreateFixedLiteralLengths();
    private static readonly byte[] FixedDistanceLengths = Enumerable.Repeat((byte)5, 30).ToArray();

    private readonly Stream _destination;
    private readonly byte[] _window = new byte[2 * WindowSize + MaxMatch];
    private readonly int[] _head = new int[HashSize];
    private readonly int[] _previous = new int[WindowSize];
    private readonly ushort[] _symbolLengths = new ushort[SymbolCapacity];
    private readonly ushort[] _symbolDistances = new ushort[SymbolCapacity];
    private readonly byte[] _raw = new byte[RawCapacity + MaxMatch];
    private readonly int[] _literalFrequencies = new int[286];
    private readonly int[] _distanceFrequencies = new int[30];
    private readonly byte[] _output = new byte[64 * 1024];
    private int _outputCount;
    private ulong _bits;
    private int _bitCount;
    private int _symbols;
    private int _rawCount;
    private int _start;
    private int _lookahead;
    private int _matchLength = MinMatch - 1;
    private int _matchStart;
    private int _previousLength;
    private int _previousMatch;
    private bool _matchAvailable;
    private bool _finished;

    internal ArchiveDeflateEncoder(Stream destination) {
        _destination = destination;
        // Position 0 doubles as "no match"; the first window byte is still emitted, it just never starts a match.
        _start = 1;
    }

    internal void Write(byte[] buffer, int offset, int count) {
        if (_finished) throw new InvalidOperationException("The Deflate stream is already finished.");
        while (count > 0) {
            if (_start + _lookahead >= 2 * WindowSize) Slide();
            int space = 2 * WindowSize - (_start + _lookahead);
            int chunk = Math.Min(space, count);
            Buffer.BlockCopy(buffer, offset, _window, _start + _lookahead, chunk);
            _lookahead += chunk; offset += chunk; count -= chunk;
            Process(final: false);
        }
    }

    /// <summary>Encodes everything remaining as the final block and pads the last byte.</summary>
    internal void Finish() {
        if (_finished) return;
        _finished = true;
        Process(final: true);
        if (_matchAvailable) { AddLiteral(_window[_start - 1]); _matchAvailable = false; }
        FlushBlock(final: true);
        if (_bitCount > 0) { _output[_outputCount++] = (byte)_bits; _bits = 0; _bitCount = 0; }
        DrainOutput();
    }

    private void Process(bool final) {
        while (_lookahead >= MinLookahead || final && _lookahead > 0) {
            if (_start >= WindowSize + MaxDistance) Slide();
            int hashHead = 0;
            if (_lookahead >= MinMatch) hashHead = Insert(_start);
            _previousLength = _matchLength; _previousMatch = _matchStart; _matchLength = MinMatch - 1;
            if (hashHead != 0 && _previousLength < MaxLazy && _start - hashHead <= MaxDistance) {
                _matchLength = LongestMatch(hashHead);
                if (_matchLength == MinMatch && _start - _matchStart > TooFar) _matchLength = MinMatch - 1;
            }
            if (_previousLength >= MinMatch && _matchLength <= _previousLength) {
                int maxInsert = _start + _lookahead - MinMatch;
                AddMatch(_start - 1 - _previousMatch, _previousLength, _start - 1);
                _lookahead -= _previousLength - 1;
                int remaining = _previousLength - 2;
                do { if (++_start <= maxInsert) Insert(_start); } while (--remaining != 0);
                _matchAvailable = false; _matchLength = MinMatch - 1; _start++;
            } else if (_matchAvailable) {
                AddLiteral(_window[_start - 1]);
                _start++; _lookahead--;
            } else {
                _matchAvailable = true; _start++; _lookahead--;
            }
        }
    }

    private int Insert(int position) {
        int hash = ((_window[position] << 10) ^ (_window[position + 1] << 5) ^ _window[position + 2]) & HashMask;
        int head = _head[hash];
        _previous[position & WindowMask] = head;
        _head[hash] = position;
        return head;
    }

    private int LongestMatch(int current) {
        int chain = MaxChain;
        int best = _previousLength;
        int nice = Math.Min(NiceLength, _lookahead);
        int limit = _start > MaxDistance ? _start - MaxDistance : 0;
        int maximum = Math.Min(MaxMatch, _lookahead);
        if (_previousLength >= GoodLength) chain >>= 2;
        byte[] window = _window;
        int scan = _start;
        do {
            int match = current;
            if (best >= maximum) break;
            if (window[match + best] != window[scan + best] || window[match + best - 1] != window[scan + best - 1]
                || window[match] != window[scan] || window[match + 1] != window[scan + 1]) continue;
            int length = 2;
            while (length < maximum && window[match + length] == window[scan + length]) length++;
            if (length > best) {
                _matchStart = current; best = length;
                if (length >= nice) break;
            }
        } while ((current = _previous[current & WindowMask]) > limit && --chain != 0);
        return Math.Min(best, _lookahead);
    }

    private void Slide() {
        Buffer.BlockCopy(_window, WindowSize, _window, 0, WindowSize + MaxMatch);
        _start -= WindowSize; _matchStart -= WindowSize; _previousMatch -= WindowSize;
        for (int index = 0; index < HashSize; index++) _head[index] = _head[index] > WindowSize ? _head[index] - WindowSize : 0;
        for (int index = 0; index < WindowSize; index++) _previous[index] = _previous[index] > WindowSize ? _previous[index] - WindowSize : 0;
    }

    private void AddLiteral(byte value) {
        _symbolLengths[_symbols] = value; _symbolDistances[_symbols] = 0; _symbols++;
        _literalFrequencies[value]++;
        _raw[_rawCount++] = value;
        if (_symbols == SymbolCapacity || _rawCount >= RawCapacity) FlushBlock(final: false);
    }

    private void AddMatch(int distance, int length, int position) {
        _symbolLengths[_symbols] = (ushort)length; _symbolDistances[_symbols] = (ushort)distance; _symbols++;
        _literalFrequencies[257 + LengthCodes[length]]++;
        _distanceFrequencies[DistanceCodes[distance - 1]]++;
        Buffer.BlockCopy(_window, position, _raw, _rawCount, length);
        _rawCount += length;
        if (_symbols == SymbolCapacity || _rawCount >= RawCapacity) FlushBlock(final: false);
    }

    private void FlushBlock(bool final) {
        _literalFrequencies[256]++;
        var literalLengths = new byte[286];
        var distanceLengths = new byte[30];
        BuildLengths(_literalFrequencies, literalLengths, 15);
        BuildLengths(_distanceFrequencies, distanceLengths, 15);
        int literalCount = 286; while (literalCount > 257 && literalLengths[literalCount - 1] == 0) literalCount--;
        int distanceCount = 30; while (distanceCount > 1 && distanceLengths[distanceCount - 1] == 0) distanceCount--;
        var runs = new List<int>();
        AppendCodeLengthRuns(literalLengths, literalCount, runs);
        AppendCodeLengthRuns(distanceLengths, distanceCount, runs);
        var codeLengthFrequencies = new int[19];
        foreach (int run in runs) codeLengthFrequencies[run & 0xff]++;
        var codeLengthLengths = new byte[19];
        BuildLengths(codeLengthFrequencies, codeLengthLengths, 7);
        int codeLengthCount = 19; while (codeLengthCount > 4 && codeLengthLengths[CodeLengthOrder[codeLengthCount - 1]] == 0) codeLengthCount--;

        long dynamicBits = 3 + 5 + 5 + 4 + 3L * codeLengthCount + DataBits(literalLengths, distanceLengths);
        foreach (int run in runs) {
            int symbol = run & 0xff;
            dynamicBits += codeLengthLengths[symbol] + (symbol == 16 ? 2 : symbol == 17 ? 3 : symbol == 18 ? 7 : 0);
        }
        long fixedBits = 3 + DataBits(FixedLiteralLengths, FixedDistanceLengths);
        long storedBits = StoredBits();
        if (storedBits <= dynamicBits && storedBits <= fixedBits) WriteStored(final);
        else if (fixedBits <= dynamicBits) {
            WriteBits(final ? 1 : 0, 1); WriteBits(1, 2);
            WriteData(Codes(FixedLiteralLengths), FixedLiteralLengths, Codes(FixedDistanceLengths), FixedDistanceLengths);
        } else {
            WriteBits(final ? 1 : 0, 1); WriteBits(2, 2);
            WriteBits(literalCount - 257, 5); WriteBits(distanceCount - 1, 5); WriteBits(codeLengthCount - 4, 4);
            for (int index = 0; index < codeLengthCount; index++) WriteBits(codeLengthLengths[CodeLengthOrder[index]], 3);
            ushort[] codeLengthCodes = Codes(codeLengthLengths);
            foreach (int run in runs) {
                int symbol = run & 0xff;
                WriteBits(codeLengthCodes[symbol], codeLengthLengths[symbol]);
                if (symbol == 16) WriteBits(run >> 8, 2);
                else if (symbol == 17) WriteBits(run >> 8, 3);
                else if (symbol == 18) WriteBits(run >> 8, 7);
            }
            WriteData(Codes(literalLengths), literalLengths, Codes(distanceLengths), distanceLengths);
        }
        _symbols = 0; _rawCount = 0;
        Array.Clear(_literalFrequencies, 0, _literalFrequencies.Length);
        Array.Clear(_distanceFrequencies, 0, _distanceFrequencies.Length);
        DrainOutput();
    }

    private long DataBits(byte[] literalLengths, byte[] distanceLengths) {
        long bits = 0;
        for (int symbol = 0; symbol < 286; symbol++) {
            int frequency = _literalFrequencies[symbol];
            if (frequency == 0) continue;
            bits += (long)frequency * (literalLengths[symbol] + (symbol > 256 ? LengthExtraBits[symbol - 257] : 0));
        }
        for (int symbol = 0; symbol < 30; symbol++)
            bits += (long)_distanceFrequencies[symbol] * (distanceLengths[symbol] + DistanceExtraBits[symbol]);
        return bits;
    }

    private long StoredBits() {
        long bits = 0;
        int remaining = _rawCount;
        int bitPosition = _bitCount & 7;
        do {
            int chunk = Math.Min(remaining, 65535);
            bits += 3 + (8 - (bitPosition + 3) % 8) % 8 + 32 + 8L * chunk;
            bitPosition = 0; remaining -= chunk;
        } while (remaining > 0);
        return bits;
    }

    private void WriteStored(bool final) {
        int offset = 0;
        do {
            int chunk = Math.Min(_rawCount - offset, 65535);
            bool last = offset + chunk == _rawCount;
            WriteBits(final && last ? 1 : 0, 1); WriteBits(0, 2);
            if (_bitCount > 0) WriteBits(0, 8 - _bitCount % 8 == 8 ? 0 : 8 - _bitCount % 8);
            WriteBits(chunk, 16); WriteBits(~chunk & 0xffff, 16);
            FlushBitBytes();
            for (int index = 0; index < chunk; index++) {
                if (_outputCount == _output.Length) DrainOutput();
                _output[_outputCount++] = _raw[offset + index];
            }
            offset += chunk;
        } while (offset < _rawCount);
    }

    private void WriteData(ushort[] literalCodes, byte[] literalLengths, ushort[] distanceCodes, byte[] distanceLengths) {
        for (int index = 0; index < _symbols; index++) {
            int distance = _symbolDistances[index];
            int value = _symbolLengths[index];
            if (distance == 0) { WriteBits(literalCodes[value], literalLengths[value]); continue; }
            int lengthCode = LengthCodes[value];
            WriteBits(literalCodes[257 + lengthCode], literalLengths[257 + lengthCode]);
            if (LengthExtraBits[lengthCode] > 0) WriteBits(value - LengthBases[lengthCode], LengthExtraBits[lengthCode]);
            int distanceCode = DistanceCodes[distance - 1];
            WriteBits(distanceCodes[distanceCode], distanceLengths[distanceCode]);
            if (DistanceExtraBits[distanceCode] > 0) WriteBits(distance - DistanceBases[distanceCode], DistanceExtraBits[distanceCode]);
        }
        WriteBits(literalCodes[256], literalLengths[256]);
    }

    private static void AppendCodeLengthRuns(byte[] lengths, int count, List<int> runs) {
        int index = 0;
        while (index < count) {
            int value = lengths[index];
            int run = 1;
            while (index + run < count && lengths[index + run] == value) run++;
            if (value == 0 && run >= 3) {
                int remaining = run;
                while (remaining >= 3) {
                    if (remaining >= 11) { int take = Math.Min(remaining, 138); runs.Add(18 | (take - 11) << 8); remaining -= take; }
                    else { runs.Add(17 | (remaining - 3) << 8); remaining = 0; }
                }
                for (int left = 0; left < remaining; left++) runs.Add(0);
            } else if (value != 0 && run >= 4) {
                runs.Add(value);
                int remaining = run - 1;
                while (remaining >= 3) { int take = Math.Min(remaining, 6); runs.Add(16 | (take - 3) << 8); remaining -= take; }
                for (int left = 0; left < remaining; left++) runs.Add(value);
            } else {
                for (int left = 0; left < run; left++) runs.Add(value);
            }
            index += run;
        }
    }

    /// <summary>Builds deterministic length-limited Huffman code lengths. At least two symbols receive codes,
    /// so every tree is complete as PKZIP-era decoders require.</summary>
    private static void BuildLengths(int[] sourceFrequencies, byte[] lengths, int maximumBits) {
        int count = sourceFrequencies.Length;
        var frequencies = (int[])sourceFrequencies.Clone();
        int used = frequencies.Count(value => value > 0);
        for (int symbol = 0; used < 2 && symbol < count; symbol++) if (frequencies[symbol] == 0) { frequencies[symbol] = 1; used++; }
        while (true) {
            Array.Clear(lengths, 0, lengths.Length);
            int[] symbols = Enumerable.Range(0, count).Where(symbol => frequencies[symbol] > 0)
                .OrderBy(symbol => frequencies[symbol]).ThenBy(symbol => symbol).ToArray();
            int leaves = symbols.Length;
            var weights = new long[2 * leaves];
            var parents = new int[2 * leaves];
            for (int index = 0; index < leaves; index++) weights[index] = frequencies[symbols[index]];
            int nextLeaf = 0, nextNode = leaves, created = leaves;
            while (created < 2 * leaves - 1) {
                int first = Take(weights, leaves, ref nextLeaf, ref nextNode, created);
                int second = Take(weights, leaves, ref nextLeaf, ref nextNode, created);
                weights[created] = weights[first] + weights[second];
                parents[first] = created; parents[second] = created;
                created++;
            }
            int root = 2 * leaves - 2;
            var depths = new int[2 * leaves];
            int deepest = 0;
            for (int node = root - 1; node >= 0; node--) {
                depths[node] = depths[parents[node]] + 1;
                if (node < leaves) deepest = Math.Max(deepest, depths[node]);
            }
            if (deepest <= maximumBits) {
                for (int index = 0; index < leaves; index++) lengths[symbols[index]] = (byte)depths[index];
                return;
            }
            for (int symbol = 0; symbol < count; symbol++) if (frequencies[symbol] > 0) frequencies[symbol] = Math.Max(1, frequencies[symbol] >> 1);
        }
    }

    private static int Take(long[] weights, int leaves, ref int nextLeaf, ref int nextNode, int created) {
        if (nextLeaf < leaves && (nextNode >= created || weights[nextLeaf] <= weights[nextNode])) return nextLeaf++;
        return nextNode++;
    }

    private static ushort[] Codes(byte[] lengths) {
        var counts = new int[16];
        foreach (byte length in lengths) counts[length]++;
        counts[0] = 0;
        var next = new int[16];
        int code = 0;
        for (int bits = 1; bits < 16; bits++) { code = (code + counts[bits - 1]) << 1; next[bits] = code; }
        var codes = new ushort[lengths.Length];
        for (int symbol = 0; symbol < lengths.Length; symbol++) {
            int length = lengths[symbol];
            if (length == 0) continue;
            int value = next[length]++;
            int reversed = 0;
            for (int bit = 0; bit < length; bit++) { reversed = (reversed << 1) | (value & 1); value >>= 1; }
            codes[symbol] = (ushort)reversed;
        }
        return codes;
    }

    private void WriteBits(int value, int count) {
        if (count == 0) return;
        _bits |= (ulong)(uint)value << _bitCount;
        _bitCount += count;
        if (_bitCount >= 32) FlushBitBytes();
    }

    private void FlushBitBytes() {
        while (_bitCount >= 8) {
            if (_outputCount == _output.Length) DrainOutput();
            _output[_outputCount++] = (byte)_bits;
            _bits >>= 8; _bitCount -= 8;
        }
    }

    private void DrainOutput() {
        FlushBitBytes();
        if (_outputCount == 0) return;
        _destination.Write(_output, 0, _outputCount);
        _outputCount = 0;
    }

    private static byte[] CreateLengthCodes() {
        var codes = new byte[MaxMatch + 1];
        for (int code = 0; code < LengthBases.Length; code++) {
            int end = code + 1 < LengthBases.Length ? LengthBases[code + 1] : MaxMatch + 1;
            for (int length = LengthBases[code]; length < end; length++) codes[length] = (byte)code;
        }
        codes[MaxMatch] = 28;
        return codes;
    }

    private static byte[] CreateDistanceCodes() {
        var codes = new byte[WindowSize];
        for (int code = 0; code < DistanceBases.Length; code++) {
            int end = code + 1 < DistanceBases.Length ? DistanceBases[code + 1] : WindowSize + 1;
            for (int distance = DistanceBases[code]; distance < end; distance++) codes[distance - 1] = (byte)code;
        }
        return codes;
    }

    private static byte[] CreateFixedLiteralLengths() {
        var lengths = new byte[288];
        for (int symbol = 0; symbol < 288; symbol++) lengths[symbol] = (byte)(symbol < 144 ? 8 : symbol < 256 ? 9 : symbol < 280 ? 7 : 8);
        // All 288 lengths take part in the canonical code assignment, although symbols 286 and 287 never occur.
        return lengths;
    }
}
