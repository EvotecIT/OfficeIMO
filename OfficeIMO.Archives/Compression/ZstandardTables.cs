namespace OfficeIMO.Archives.Compression;

/// <summary>An FSE decoding table (RFC 8878, section 4.1.1).</summary>
internal sealed class ZstandardFseTable {
    private ZstandardFseTable(int accuracyLog, byte[] symbols, byte[] bits, int[] baselines) {
        AccuracyLog = accuracyLog;
        Symbols = symbols;
        Bits = bits;
        Baselines = baselines;
    }

    internal int AccuracyLog { get; }
    internal byte[] Symbols { get; }
    internal byte[] Bits { get; }
    internal int[] Baselines { get; }

    internal static ZstandardFseTable Rle(byte symbol) =>
        new(0, new[] { symbol }, new byte[1], new int[1]);

    /// <summary>Reads a table description and builds the table. Returns the bytes the description used.</summary>
    internal static int Read(byte[] data, int start, int end, int maxSymbol, int maxAccuracyLog, long inputOffset,
        string structure, out ZstandardFseTable table) {
        var bits = new ZstandardForwardBits(data, start, end);
        int accuracyLog = (int)bits.Read(4) + 5;
        if (accuracyLog > maxAccuracyLog)
            throw ZstandardDecoder.Corrupt(inputOffset, $"an accuracy log of at most {maxAccuracyLog} for the {structure}");
        var counts = new short[maxSymbol + 1];
        int remaining = (1 << accuracyLog) + 1;
        int threshold = 1 << accuracyLog;
        int bitCount = accuracyLog + 1;
        int symbol = 0;
        while (remaining > 1) {
            if (symbol > maxSymbol)
                throw ZstandardDecoder.Corrupt(inputOffset, $"at most {maxSymbol + 1} symbols in the {structure}");
            int max = (2 * threshold - 1) - remaining;
            int peek = (int)bits.Peek(bitCount);
            int value;
            if ((peek & (threshold - 1)) < max) {
                value = peek & (threshold - 1);
                bits.Skip(bitCount - 1);
            } else {
                value = peek & (2 * threshold - 1);
                if (value >= threshold) value -= max;
                bits.Skip(bitCount);
            }
            int probability = value - 1;
            remaining -= probability < 0 ? -probability : probability;
            counts[symbol++] = (short)probability;
            if (probability == 0) {
                while (true) {
                    int repeat = (int)bits.Read(2);
                    for (int index = 0; index < repeat; index++) {
                        if (symbol > maxSymbol)
                            throw ZstandardDecoder.Corrupt(inputOffset, $"at most {maxSymbol + 1} symbols in the {structure}");
                        counts[symbol++] = 0;
                    }
                    if (repeat != 3) break;
                }
            }
            while (remaining < threshold) {
                bitCount--;
                threshold >>= 1;
            }
        }
        int consumed = bits.BytesConsumed;
        if (remaining != 1 || consumed > end - start)
            throw ZstandardDecoder.Corrupt(inputOffset, "a complete " + structure);
        table = Build(counts, symbol, accuracyLog, inputOffset, structure);
        return consumed;
    }

    internal static ZstandardFseTable Build(short[] counts, int symbolCount, int accuracyLog, long inputOffset, string structure) {
        int size = 1 << accuracyLog;
        var symbols = new byte[size];
        var next = new int[symbolCount];
        int high = size - 1;
        for (int symbol = 0; symbol < symbolCount; symbol++) {
            if (counts[symbol] == -1) {
                symbols[high--] = (byte)symbol;
                next[symbol] = 1;
            } else {
                next[symbol] = counts[symbol];
            }
        }
        int position = 0;
        int step = (size >> 1) + (size >> 3) + 3;
        int mask = size - 1;
        for (int symbol = 0; symbol < symbolCount; symbol++) {
            for (int index = 0; index < counts[symbol]; index++) {
                symbols[position] = (byte)symbol;
                do {
                    position = (position + step) & mask;
                } while (position > high);
            }
        }
        if (position != 0) throw ZstandardDecoder.Corrupt(inputOffset, "probabilities that fill the " + structure);
        var bits = new byte[size];
        var baselines = new int[size];
        for (int state = 0; state < size; state++) {
            int symbol = symbols[state];
            int nextState = next[symbol]++;
            int bitCount = accuracyLog - ZstandardDecoder.HighBit((uint)nextState);
            bits[state] = (byte)bitCount;
            baselines[state] = (nextState << bitCount) - size;
        }
        return new ZstandardFseTable(accuracyLog, symbols, bits, baselines);
    }
}

/// <summary>A Huffman decoding table for literals (RFC 8878, section 4.2).</summary>
internal sealed class ZstandardHuffmanTable {
    private const int MaxBitsLimit = 11;

    private ZstandardHuffmanTable(int maxBits, byte[] symbols, byte[] lengths) {
        MaxBits = maxBits;
        Symbols = symbols;
        Lengths = lengths;
    }

    internal int MaxBits { get; }
    internal byte[] Symbols { get; }
    internal byte[] Lengths { get; }

    /// <summary>Reads a Huffman tree description and returns the bytes it used.</summary>
    internal static int Read(byte[] data, int start, int end, long inputOffset, ZstandardBackwardBits reader,
        out ZstandardHuffmanTable table) {
        if (start >= end) throw ZstandardDecoder.Corrupt(inputOffset, "a Huffman tree description");
        int header = data[start];
        var weights = new byte[256];
        int count;
        int consumed;
        if (header >= 128) {
            count = header - 127;
            int bytes = (count + 1) / 2;
            if (bytes > end - start - 1) throw ZstandardDecoder.Corrupt(inputOffset, "complete Huffman weights");
            for (int index = 0; index < count; index++) {
                int packed = data[start + 1 + index / 2];
                weights[index] = (byte)((index & 1) == 0 ? packed >> 4 : packed & 15);
            }
            consumed = 1 + bytes;
        } else {
            if (header == 0 || header > end - start - 1)
                throw ZstandardDecoder.Corrupt(inputOffset, "complete FSE-compressed Huffman weights");
            int weightsEnd = start + 1 + header;
            int description = ZstandardFseTable.Read(data, start + 1, weightsEnd, 255, 6, inputOffset + 1,
                "Huffman weight table", out ZstandardFseTable fse);
            reader.Reset(data, start + 1 + description, header - description, inputOffset + 1 + description, "Huffman weight stream");
            count = DecodeWeights(fse, reader, weights, inputOffset);
            consumed = 1 + header;
        }
        table = Build(weights, count, inputOffset);
        return consumed;
    }

    private static int DecodeWeights(ZstandardFseTable table, ZstandardBackwardBits reader, byte[] weights, long inputOffset) {
        // Two interleaved states share one stream; decoding stops once an update reads past its start.
        int first = (int)reader.Read(table.AccuracyLog);
        int second = (int)reader.Read(table.AccuracyLog);
        // Reading past the start is how the stream ends, but only after both states have been initialized.
        if (reader.Overflowed) throw ZstandardDecoder.Corrupt(inputOffset, "initial states for both Huffman weight decoders");
        int count = 0;
        while (true) {
            Emit(table.Symbols[first]);
            first = table.Baselines[first] + (int)reader.Read(table.Bits[first]);
            if (reader.Overflowed) {
                Emit(table.Symbols[second]);
                break;
            }
            Emit(table.Symbols[second]);
            second = table.Baselines[second] + (int)reader.Read(table.Bits[second]);
            if (reader.Overflowed) {
                Emit(table.Symbols[first]);
                break;
            }
        }
        return count;

        void Emit(byte weight) {
            if (count == 255) throw ZstandardDecoder.Corrupt(inputOffset, "at most 255 Huffman weights");
            weights[count++] = weight;
        }
    }

    private static ZstandardHuffmanTable Build(byte[] weights, int count, long inputOffset) {
        long total = 0;
        for (int index = 0; index < count; index++) {
            int weight = weights[index];
            if (weight > MaxBitsLimit) throw ZstandardDecoder.Corrupt(inputOffset, "Huffman weights of at most 11");
            if (weight > 0) total += 1L << (weight - 1);
        }
        if (total == 0) throw ZstandardDecoder.Corrupt(inputOffset, "at least one non-zero Huffman weight");
        int maxBits = ZstandardDecoder.HighBit((uint)total) + 1;
        long leftover = (1L << maxBits) - total;
        if (maxBits > MaxBitsLimit || (leftover & (leftover - 1)) != 0)
            throw ZstandardDecoder.Corrupt(inputOffset, "Huffman weights that complete a prefix code");
        weights[count++] = (byte)(ZstandardDecoder.HighBit((uint)leftover) + 1);
        // The shortest codes come in pairs, so a complete tree has an even number of weight-1 symbols, at least two.
        int ones = 0;
        for (int index = 0; index < count; index++) if (weights[index] == 1) ones++;
        if (ones < 2 || (ones & 1) != 0) throw ZstandardDecoder.Corrupt(inputOffset, "an even number of at least two weight-1 Huffman symbols");
        int size = 1 << maxBits;
        var symbols = new byte[size];
        var lengths = new byte[size];
        int position = 0;
        for (int weight = 1; weight <= maxBits; weight++) {
            int span = 1 << (weight - 1);
            byte length = (byte)(maxBits + 1 - weight);
            for (int symbol = 0; symbol < count; symbol++) {
                if (weights[symbol] != weight) continue;
                for (int index = 0; index < span; index++) {
                    symbols[position] = (byte)symbol;
                    lengths[position++] = length;
                }
            }
        }
        if (position != size) throw ZstandardDecoder.Corrupt(inputOffset, "Huffman weights that complete a prefix code");
        return new ZstandardHuffmanTable(maxBits, symbols, lengths);
    }
}
