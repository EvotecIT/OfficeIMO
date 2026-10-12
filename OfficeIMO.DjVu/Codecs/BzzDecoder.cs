namespace OfficeIMO.DjVu;

// DjVu v3, appendix 4: adaptive MTF ranks followed by inverse Burrows-Wheeler.
internal static class BzzDecoder {
    internal static byte[] Decode(byte[] source, int offset, int length, DjVuReadBudget budget) {
        if (length == 0) throw new InvalidDataException("Empty BZZ stream.");
        var arithmetic = new ZpDecoder(source, offset, length, budget.Cancellation);
        var contexts = new byte[262];
        using var output = new MemoryStream();
        while (true) {
            budget.Cancellation.ThrowIfCancellationRequested();
            int size = arithmetic.Raw(24);
            if (size == 0) return output.ToArray();
            if (size > budget.Options.MaxBzzBlockBytes) throw new DjVuResourceLimitException(nameof(DjVuReadOptions.MaxBzzBlockBytes));
            budget.Expanded(size - 1);
            // Both output growth/snapshot and the inverse transform are bounded before allocation.
            budget.WorkingBytes((output.Length + size) * 3 + size * 6L + 4096);
            byte[] block = DecodeBlock(arithmetic, contexts, size, budget.Cancellation);
            output.Write(block, 0, size - 1);
        }
    }

    private static byte[] DecodeBlock(ZpDecoder arithmetic, byte[] contexts, int size, CancellationToken cancellation) {
        int speed = arithmetic.RawBit() == 0 ? 0 : arithmetic.RawBit() == 0 ? 1 : 2;
        var symbols = new byte[256];
        for (int i = 0; i < symbols.Length; i++) symbols[i] = (byte)i;
        var frequencies = new uint[4];
        uint increment = 4;
        var transformed = new byte[size];
        int previousRank = 3;
        int marker = -1;
        for (int i = 0; i < size; i++) {
            if ((i & 4095) == 0) cancellation.ThrowIfCancellationRequested();
            int context = Math.Min(previousRank, 2);
            int rank;
            if (arithmetic.Bit(ref contexts[context]) != 0) rank = 0;
            else if (arithmetic.Bit(ref contexts[context + 3]) != 0) rank = 1;
            else {
                rank = 256;
                for (int bits = 1; bits <= 7; bits++) {
                    int first = (1 << bits) + 4;
                    if (arithmetic.Bit(ref contexts[first]) == 0) continue;
                    rank = (1 << bits) + DecodeRank(arithmetic, contexts, first + 1, bits);
                    break;
                }
            }
            previousRank = rank;
            if (rank == 256) {
                if (marker >= 0) throw new InvalidDataException("BZZ block contains multiple end markers.");
                marker = i;
                continue;
            }
            byte symbol = symbols[rank];
            transformed[i] = symbol;
            increment += increment >> speed;
            if (increment > 0x10000000) {
                increment >>= 24;
                for (int k = 0; k < 4; k++) frequencies[k] >>= 24;
            }
            uint frequency = increment + (rank < 4 ? frequencies[rank] : 0);
            int destination = rank;
            while (destination > 3) {
                symbols[destination] = symbols[destination - 1];
                destination--;
            }
            while (destination > 0 && frequency >= frequencies[destination - 1]) {
                symbols[destination] = symbols[destination - 1];
                frequencies[destination] = frequencies[destination - 1];
                destination--;
            }
            symbols[destination] = symbol;
            frequencies[destination] = frequency;
        }
        if (marker <= 0) throw new InvalidDataException($"BZZ block of {size} bytes has no valid end marker ({marker}).");
        return Invert(transformed, marker, cancellation);
    }

    private static int DecodeRank(ZpDecoder arithmetic, byte[] contexts, int offset, int bits) {
        int node = 1;
        int end = 1 << bits;
        while (node < end) node = (node << 1) | arithmetic.Bit(ref contexts[offset + node - 1]);
        return node - end;
    }

    private static byte[] Invert(byte[] transformed, int marker, CancellationToken cancellation) {
        var counts = new int[256];
        var ranks = new int[transformed.Length];
        for (int i = 0; i < transformed.Length; i++) {
            if ((i & 4095) == 0) cancellation.ThrowIfCancellationRequested();
            if (i != marker) ranks[i] = counts[transformed[i]]++;
        }
        int next = 1;
        for (int i = 0; i < counts.Length; i++) {
            int count = counts[i];
            counts[i] = next;
            next += count;
        }
        var decoded = new byte[transformed.Length - 1];
        int position = 0;
        for (int i = decoded.Length - 1; i >= 0; i--) {
            if ((i & 4095) == 0) cancellation.ThrowIfCancellationRequested();
            if (position == marker || (uint)position >= (uint)transformed.Length) throw new InvalidDataException("Invalid BZZ inverse transform.");
            byte symbol = transformed[position];
            decoded[i] = symbol;
            position = counts[symbol] + ranks[position];
        }
        if (position != marker) throw new InvalidDataException("BZZ inverse transform does not reach its end marker.");
        return decoded;
    }
}
