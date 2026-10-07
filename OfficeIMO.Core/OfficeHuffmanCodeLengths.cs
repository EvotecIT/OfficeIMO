using System;

namespace OfficeIMO.Drawing;

/// <summary>Deterministic frequency-based Huffman lengths shared by native raster encoders.</summary>
internal static class OfficeHuffmanCodeLengths {
    internal static byte[] Create(int[] frequencies, int maximumDepth) {
        int[] depths = BuildDepths(frequencies);
        var counts = new int[Math.Max(33, frequencies.Length + 1)];
        int symbols = 0;
        foreach (int depth in depths) if (depth > 0) { counts[depth]++; symbols++; }
        var result = new byte[frequencies.Length];
        if (symbols == 0) { result[0] = 1; return result; }
        if (symbols > (1 << maximumDepth)) throw new ArgumentException("The Huffman alphabet exceeds its depth limit.", nameof(frequencies));
        LimitDepthCounts(counts, maximumDepth);
        var ordered = new int[symbols];
        int position = 0;
        for (int symbol = 0; symbol < frequencies.Length; symbol++) if (frequencies[symbol] > 0) ordered[position++] = symbol;
        Array.Sort(ordered, (left, right) => {
            int comparison = frequencies[right].CompareTo(frequencies[left]);
            return comparison != 0 ? comparison : left.CompareTo(right);
        });
        position = 0;
        for (int depth = 1; depth <= maximumDepth; depth++) {
            for (int count = 0; count < counts[depth]; count++) result[ordered[position++]] = (byte)depth;
        }
        return result;
    }

    internal static int[] BuildDepths(int[] frequencies) {
        int maximumNodes = frequencies.Length * 2;
        var weights = new long[maximumNodes];
        var parents = new int[maximumNodes];
        var leaves = new int[frequencies.Length];
        for (int index = 0; index < leaves.Length; index++) leaves[index] = -1;
        int nodes = 0;
        for (int symbol = 0; symbol < frequencies.Length; symbol++) {
            if (frequencies[symbol] <= 0) continue;
            weights[nodes] = frequencies[symbol];
            parents[nodes] = -1;
            leaves[symbol] = nodes++;
        }
        int total = nodes;
        while (true) {
            int first = -1, second = -1;
            for (int index = 0; index < total; index++) {
                if (parents[index] != -1) continue;
                if (first < 0 || weights[index] < weights[first]) { second = first; first = index; }
                else if (second < 0 || weights[index] < weights[second]) second = index;
            }
            if (second < 0) break;
            weights[total] = weights[first] + weights[second];
            parents[first] = parents[second] = total;
            parents[total++] = -1;
        }
        var depths = new int[frequencies.Length];
        for (int symbol = 0; symbol < leaves.Length; symbol++) {
            int node = leaves[symbol];
            if (node < 0) continue;
            int depth = 0;
            while (parents[node] != -1) { depth++; node = parents[node]; }
            depths[symbol] = depth == 0 ? 1 : depth;
        }
        return depths;
    }

    /// <summary>
    /// Rebalances complete binary-tree depth counts while preserving leaves and Kraft equality.
    /// The donor must be at least two levels shallower; choosing depth minus one makes no progress.
    /// </summary>
    internal static void LimitDepthCounts(int[] counts, int maximumDepth) {
        for (int depth = counts.Length - 1; depth > maximumDepth; depth--) {
            while (counts[depth] > 0) {
                int donor = depth - 2;
                while (donor > 0 && counts[donor] == 0) donor--;
                if (donor == 0 || counts[depth] < 2) throw new InvalidOperationException("Cannot rebalance the Huffman depth counts.");
                counts[depth] -= 2;
                counts[depth - 1]++;
                counts[donor]--;
                counts[donor + 1] += 2;
            }
        }
    }
}
