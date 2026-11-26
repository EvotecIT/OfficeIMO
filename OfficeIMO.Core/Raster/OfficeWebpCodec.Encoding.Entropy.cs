namespace OfficeIMO.Drawing;

public static partial class OfficeWebpCodec {
    private const int Vp8lLengthPrefixCount = 24;
    private const int Vp8lDistancePrefixCount = 40;
    private const int Vp8lMaximumCacheBits = 10;

    private static readonly float[] Vp8lNLog2NTable = CreateVp8lNLog2NTable(4096);

    private static double Log2(double value) => Math.Log(value) * 1.4426950408889634;

    private static float[] CreateVp8lNLog2NTable(int size) {
        var table = new float[size];
        for (int index = 1; index < size; index++) table[index] = (float)(index * Log2(index));
        return table;
    }

    private static double NLog2N(long value) => value < Vp8lNLog2NTable.Length ? Vp8lNLog2NTable[value] : value * Log2(value);

    private static double Vp8lShannonBits(int[] counts, int offset, int length) {
        long total = 0;
        double sum = 0;
        for (int index = offset; index < offset + length; index++) {
            int count = counts[index];
            if (count == 0) continue;
            total += count;
            sum += NLog2N(count);
        }
        return NLog2N(total) - sum;
    }

    /// <summary>Symbol counts of one prefix-code group: green with lengths and cache indices, red, blue, alpha, distance.</summary>
    private sealed class Vp8lHistogram {
        internal Vp8lHistogram(int cacheBits) {
            Green = new int[256 + Vp8lLengthPrefixCount + (cacheBits > 0 ? 1 << cacheBits : 0)];
        }

        internal int[] Green { get; }
        internal int[] Red { get; } = new int[256];
        internal int[] Blue { get; } = new int[256];
        internal int[] Alpha { get; } = new int[256];
        internal int[] Distance { get; } = new int[Vp8lDistancePrefixCount];
        internal double Cost { get; set; }

        internal int[] Alphabet(int index) => index switch {
            0 => Green,
            1 => Red,
            2 => Blue,
            3 => Alpha,
            _ => Distance
        };

        internal void AddLiteral(uint color) {
            Green[(color >> 8) & 255]++;
            Red[(color >> 16) & 255]++;
            Blue[color & 255]++;
            Alpha[color >> 24]++;
        }

        internal void AddCacheIndex(int index) => Green[256 + Vp8lLengthPrefixCount + index]++;

        internal void AddCopy(int length, int distanceCode) {
            GetVp8lPrefix(length, Vp8lLengthPrefixCount, out int prefix, out _, out _);
            GetVp8lPrefix(distanceCode, Vp8lDistancePrefixCount, out int distancePrefix, out _, out _);
            Green[256 + prefix]++;
            Distance[distancePrefix]++;
        }

        internal void Add(Vp8lHistogram other) {
            for (int alphabet = 0; alphabet < 5; alphabet++) {
                int[] target = Alphabet(alphabet), source = other.Alphabet(alphabet);
                for (int index = 0; index < target.Length; index++) target[index] += source[index];
            }
        }

        internal bool IsEmpty() {
            foreach (int count in Green) if (count != 0) return false;
            return true;
        }

        internal double EstimateBits(Vp8lHistogram? other = null) {
            double bits = 0;
            for (int alphabet = 0; alphabet < 5; alphabet++) {
                bits += EstimateVp8lAlphabetBits(Alphabet(alphabet), other?.Alphabet(alphabet));
            }
            return bits;
        }
    }

    // Entropy bits of the symbols plus an approximation of the prefix-code header that describes them.
    private static double EstimateVp8lAlphabetBits(int[] counts, int[]? other) {
        long total = 0;
        double sum = 0;
        int used = 0, zeroRuns = 0;
        bool inZeros = false;
        for (int index = 0; index < counts.Length; index++) {
            int count = counts[index] + (other == null ? 0 : other[index]);
            if (count == 0) {
                if (!inZeros) zeroRuns++;
                inZeros = true;
                continue;
            }
            inZeros = false;
            used++;
            total += count;
            sum += NLog2N(count);
        }
        if (used <= 1) return used == 0 ? 4 : 12;
        if (used == 2) return total + 20;
        double entropy = Math.Max(NLog2N(total) - sum, total);
        return entropy + 24 + used * 3.0 + zeroRuns * 6.0;
    }

    /// <summary>
    /// Walks the tokens with the colour cache of the given size, reporting each token's start pixel.
    /// Without tokens, every pixel is a literal.
    /// </summary>
    private static void ReplayVp8lTokens(Vp8lTokens? tokens, uint[] pixels, int cacheBits, CancellationToken cancellationToken,
        Action<int, ulong, int> visit) {
        uint[] cache = cacheBits == 0 ? Array.Empty<uint>() : new uint[1 << cacheBits];
        int shift = 32 - cacheBits;
        int position = 0;
        int tokenCount = tokens?.Count ?? pixels.Length;
        for (int index = 0; index < tokenCount; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            ulong token = tokens == null ? pixels[index] : tokens[index];
            if (Vp8lTokens.IsCopy(token)) {
                int length = Vp8lTokens.Length(token);
                visit(position, token, -1);
                if (cacheBits != 0) {
                    for (int offset = 0; offset < length; offset++) {
                        uint color = pixels[position + offset];
                        cache[(int)((0x1E35A7BDU * color) >> shift)] = color;
                    }
                }
                position += length;
            } else {
                uint color = Vp8lTokens.Color(token);
                int cacheIndex = -1;
                if (cacheBits != 0) {
                    int slot = (int)((0x1E35A7BDU * color) >> shift);
                    if (cache[slot] == color) cacheIndex = slot;
                    else cache[slot] = color;
                }
                visit(position, token, cacheIndex);
                position++;
            }
        }
    }

    private static void AddVp8lToken(Vp8lHistogram histogram, ulong token, int cacheIndex) {
        if (Vp8lTokens.IsCopy(token)) histogram.AddCopy(Vp8lTokens.Length(token), Vp8lTokens.DistanceCode(token));
        else if (cacheIndex >= 0) histogram.AddCacheIndex(cacheIndex);
        else histogram.AddLiteral(Vp8lTokens.Color(token));
    }

    /// <summary>
    /// Picks the colour cache size whose single-group estimate is smallest, simulating all sizes in one pass,
    /// and reports that estimate with the copies' extra bits. Without tokens, every pixel is a literal.
    /// </summary>
    private static int SelectVp8lCacheBits(Vp8lTokens? tokens, uint[] pixels, int maximumBits, CancellationToken cancellationToken,
        out double estimatedBits, bool onlyLargest = false) {
        maximumBits = Math.Max(0, maximumBits);
        // Comparing candidates needs only a rough figure: no cache and the largest one.
        int firstBits = onlyLargest ? maximumBits : 1;
        long extraBits = 0;
        var histograms = new Vp8lHistogram[maximumBits + 1];
        var caches = new uint[maximumBits + 1][];
        for (int bits = 0; bits <= maximumBits; bits++) {
            bool simulated = bits == 0 || bits >= firstBits;
            histograms[bits] = new Vp8lHistogram(simulated ? bits : 0);
            caches[bits] = bits == 0 || !simulated ? Array.Empty<uint>() : new uint[1 << bits];
        }
        int position = 0;
        int tokenCount = tokens?.Count ?? pixels.Length;
        for (int index = 0; index < tokenCount; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            ulong token = tokens == null ? pixels[index] : tokens[index];
            if (Vp8lTokens.IsCopy(token)) {
                int length = Vp8lTokens.Length(token);
                int distanceCode = Vp8lTokens.DistanceCode(token);
                histograms[0].AddCopy(length, distanceCode);
                for (int bits = firstBits; bits <= maximumBits; bits++) histograms[bits].AddCopy(length, distanceCode);
                GetVp8lPrefix(length, Vp8lLengthPrefixCount, out _, out int lengthExtra, out _);
                GetVp8lPrefix(distanceCode, Vp8lDistancePrefixCount, out _, out int distanceExtra, out _);
                extraBits += lengthExtra + distanceExtra;
                for (int offset = 0; offset < length; offset++) {
                    uint color = pixels[position + offset];
                    uint hash = 0x1E35A7BDU * color;
                    for (int bits = firstBits; bits <= maximumBits; bits++) caches[bits][(int)(hash >> (32 - bits))] = color;
                }
                position += length;
            } else {
                uint color = Vp8lTokens.Color(token);
                uint hash = 0x1E35A7BDU * color;
                histograms[0].AddLiteral(color);
                for (int bits = firstBits; bits <= maximumBits; bits++) {
                    int slot = (int)(hash >> (32 - bits));
                    if (caches[bits][slot] == color) {
                        histograms[bits].AddCacheIndex(slot);
                    } else {
                        histograms[bits].AddLiteral(color);
                        caches[bits][slot] = color;
                    }
                }
                position++;
            }
        }
        int best = 0;
        double bestBits = double.MaxValue;
        for (int bits = 0; bits <= maximumBits; bits++) {
            if (bits != 0 && bits < firstBits) continue;
            double estimate = histograms[bits].EstimateBits() + (bits == 0 ? 1 : 5);
            if (estimate < bestBits) {
                bestBits = estimate;
                best = bits;
            }
        }
        estimatedBits = bestBits + extraBits;
        return best;
    }

    /// <summary>The prefix-code groups of an image and, with more than one, the tile map that selects them.</summary>
    private sealed class Vp8lEntropyPlan {
        internal Vp8lEntropyPlan(Vp8lHistogram[] groups, int bits, int columns, int[]? tileGroups) {
            Groups = groups;
            Bits = bits;
            Columns = columns;
            TileGroups = tileGroups;
        }

        internal Vp8lHistogram[] Groups { get; }
        internal int Bits { get; }
        internal int Columns { get; }
        internal int[]? TileGroups { get; }

        internal int GroupAt(int position, int width) {
            if (TileGroups == null) return 0;
            int y = position / width, x = position - y * width;
            return TileGroups[(y >> Bits) * Columns + (x >> Bits)];
        }
    }

    private static Vp8lEntropyPlan CreateVp8lSingleGroupPlan(Vp8lTokens? tokens, uint[] pixels, int cacheBits,
        CancellationToken cancellationToken) {
        var histogram = new Vp8lHistogram(cacheBits);
        ReplayVp8lTokens(tokens, pixels, cacheBits, cancellationToken,
            (position, token, cacheIndex) => AddVp8lToken(histogram, token, cacheIndex));
        return new Vp8lEntropyPlan(new[] { histogram }, 0, 0, null);
    }

    /// <summary>
    /// Splits the image into tiles, gives each tile its own histogram, and merges histograms while merging
    /// saves more bits than the extra header it removes costs: cheap merges within entropy bins first,
    /// then the best pair at a time, then a pass that moves each tile to the group that codes it best.
    /// </summary>
    private static Vp8lEntropyPlan CreateVp8lEntropyPlan(Vp8lTokens tokens, uint[] pixels, int width, int height,
        int cacheBits, int maximumTiles, int exactClusters, CancellationToken cancellationToken) {
        int bits = 2;
        while (bits < 9 && (long)DivideRoundUp(width, 1 << bits) * DivideRoundUp(height, 1 << bits) > maximumTiles) bits++;
        int columns = DivideRoundUp(width, 1 << bits), rows = DivideRoundUp(height, 1 << bits);
        if (columns * rows <= 1) return CreateVp8lSingleGroupPlan(tokens, pixels, cacheBits, cancellationToken);
        var tiles = new Vp8lHistogram[columns * rows];
        for (int index = 0; index < tiles.Length; index++) tiles[index] = new Vp8lHistogram(cacheBits);
        ReplayVp8lTokens(tokens, pixels, cacheBits, cancellationToken, (position, token, cacheIndex) => {
            int y = position / width, x = position - y * width;
            AddVp8lToken(tiles[(y >> bits) * columns + (x >> bits)], token, cacheIndex);
        });

        // Each cluster starts as one non-empty tile; empty tiles join whatever group is first.
        var clusters = new List<Vp8lHistogram>();
        var tileCluster = new int[tiles.Length];
        for (int index = 0; index < tiles.Length; index++) {
            if (tiles[index].IsEmpty()) {
                tileCluster[index] = -1;
                continue;
            }
            var copy = new Vp8lHistogram(cacheBits);
            copy.Add(tiles[index]);
            copy.Cost = copy.EstimateBits();
            tileCluster[index] = clusters.Count;
            clusters.Add(copy);
        }
        if (clusters.Count <= 1) return CreateVp8lSingleGroupPlan(tokens, pixels, cacheBits, cancellationToken);
        var owner = new int[clusters.Count];
        for (int index = 0; index < owner.Length; index++) owner[index] = index;

        MergeVp8lEntropyBins(clusters, owner, exactClusters, cancellationToken);
        MergeVp8lSampledPairs(clusters, owner, exactClusters, cancellationToken);
        MergeVp8lBestPairs(clusters, owner, cancellationToken);

        // Resolve each tile's surviving cluster, then let every tile move to the cluster that codes it best.
        var alive = new List<int>();
        var renumber = new int[clusters.Count];
        for (int index = 0; index < clusters.Count; index++) {
            renumber[index] = -1;
            if (owner[index] == index) {
                renumber[index] = alive.Count;
                alive.Add(index);
            }
        }
        var groups = new Vp8lHistogram[alive.Count];
        for (int index = 0; index < groups.Length; index++) groups[index] = clusters[alive[index]];
        var tileGroups = new int[tiles.Length];
        for (int index = 0; index < tiles.Length; index++) {
            tileGroups[index] = tileCluster[index] < 0 ? 0 : renumber[Find(owner, tileCluster[index])];
        }
        if (groups.Length > 1) {
            RefineVp8lTileGroups(tiles, tileGroups, ref groups, cacheBits, cancellationToken);
        }
        if (groups.Length == 1) return CreateVp8lSingleGroupPlan(tokens, pixels, cacheBits, cancellationToken);
        return new Vp8lEntropyPlan(groups, bits, columns, tileGroups);
    }

    private static int Find(int[] owner, int index) {
        while (owner[index] != index) index = owner[index] = owner[owner[index]];
        return index;
    }

    private static void Merge(List<Vp8lHistogram> clusters, int[] owner, int target, int source, double cost) {
        clusters[target].Add(clusters[source]);
        clusters[target].Cost = cost;
        owner[source] = target;
    }

    // Histograms with similar shares of literal, red and blue information merge first, cheaply.
    private static void MergeVp8lEntropyBins(List<Vp8lHistogram> clusters, int[] owner, int exactClusters,
        CancellationToken cancellationToken) {
        if (clusters.Count <= exactClusters) return;
        var binOwner = new Dictionary<int, int>();
        for (int index = 0; index < clusters.Count; index++) {
            if ((index & 63) == 0) cancellationToken.ThrowIfCancellationRequested();
            Vp8lHistogram histogram = clusters[index];
            double green = EstimateVp8lAlphabetBits(histogram.Green, null);
            double red = EstimateVp8lAlphabetBits(histogram.Red, null);
            double blue = EstimateVp8lAlphabetBits(histogram.Blue, null);
            double total = green + red + blue + 1;
            int key = Quantize(green / total) * 64 + Quantize(red / total) * 8 + Quantize(blue / total);
            if (!binOwner.TryGetValue(key, out int target)) {
                binOwner[key] = index;
                continue;
            }
            double merged = clusters[target].EstimateBits(histogram);
            if (merged < clusters[target].Cost + histogram.Cost) Merge(clusters, owner, target, index, merged);
        }

        static int Quantize(double share) => Math.Min(7, (int)(share * 8));
    }

    // Deterministic random pairs bring a large cluster count within reach of the exact search.
    private static void MergeVp8lSampledPairs(List<Vp8lHistogram> clusters, int[] owner, int exactClusters,
        CancellationToken cancellationToken) {
        var alive = new List<int>();
        for (int index = 0; index < clusters.Count; index++) if (owner[index] == index) alive.Add(index);
        uint state = 0x9E3779B9U;
        int failures = 0;
        while (alive.Count > exactClusters && failures < alive.Count * 2) {
            cancellationToken.ThrowIfCancellationRequested();
            int first = (int)(Next() % (uint)alive.Count), second = (int)(Next() % (uint)alive.Count);
            if (first == second) continue;
            int target = alive[first], source = alive[second];
            double merged = clusters[target].EstimateBits(clusters[source]);
            if (merged < clusters[target].Cost + clusters[source].Cost) {
                Merge(clusters, owner, target, source, merged);
                alive[second] = alive[alive.Count - 1];
                alive.RemoveAt(alive.Count - 1);
                failures = 0;
            } else {
                failures++;
            }
        }
        // Whatever is still above the exact limit merges with the pair that costs least.
        while (alive.Count > exactClusters) {
            cancellationToken.ThrowIfCancellationRequested();
            int target = alive[0], bestSecond = 1;
            double bestDelta = double.MaxValue, bestMerged = 0;
            for (int second = 1; second < alive.Count; second++) {
                double merged = clusters[target].EstimateBits(clusters[alive[second]]);
                double delta = merged - clusters[target].Cost - clusters[alive[second]].Cost;
                if (delta < bestDelta) {
                    bestDelta = delta;
                    bestMerged = merged;
                    bestSecond = second;
                }
            }
            Merge(clusters, owner, target, alive[bestSecond], bestMerged);
            alive[bestSecond] = alive[alive.Count - 1];
            alive.RemoveAt(alive.Count - 1);
        }

        uint Next() {
            state ^= state << 13;
            state ^= state >> 17;
            state ^= state << 5;
            return state;
        }
    }

    private static void MergeVp8lBestPairs(List<Vp8lHistogram> clusters, int[] owner, CancellationToken cancellationToken) {
        var alive = new List<int>();
        for (int index = 0; index < clusters.Count; index++) if (owner[index] == index) alive.Add(index);
        int count = alive.Count;
        if (count < 2) return;
        // gains[i, j] (i < j) is the bit saving of merging alive[i] and alive[j].
        var gains = new double[count * count];
        var merged = new double[count * count];
        for (int first = 0; first < count; first++) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int second = first + 1; second < count; second++) SetGain(first, second);
        }
        var removed = new bool[count];
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            int bestFirst = -1, bestSecond = -1;
            double bestGain = 0;
            for (int first = 0; first < count; first++) {
                if (removed[first]) continue;
                for (int second = first + 1; second < count; second++) {
                    if (removed[second]) continue;
                    double gain = gains[first * count + second];
                    if (gain > bestGain) {
                        bestGain = gain;
                        bestFirst = first;
                        bestSecond = second;
                    }
                }
            }
            if (bestFirst < 0) return;
            Merge(clusters, owner, alive[bestFirst], alive[bestSecond], merged[bestFirst * count + bestSecond]);
            removed[bestSecond] = true;
            for (int other = 0; other < count; other++) {
                if (other == bestFirst || removed[other]) continue;
                if (other < bestFirst) SetGain(other, bestFirst);
                else SetGain(bestFirst, other);
            }
        }

        void SetGain(int first, int second) {
            Vp8lHistogram a = clusters[alive[first]], b = clusters[alive[second]];
            double cost = a.EstimateBits(b);
            merged[first * count + second] = cost;
            gains[first * count + second] = a.Cost + b.Cost - cost;
        }
    }

    // Moves each tile to the group whose code is cheapest for it, keeps the result only when the estimate drops.
    private static void RefineVp8lTileGroups(Vp8lHistogram[] tiles, int[] tileGroups, ref Vp8lHistogram[] groups,
        int cacheBits, CancellationToken cancellationToken) {
        int groupCount = groups.Length;
        var tables = new float[groupCount][][];
        for (int group = 0; group < groupCount; group++) {
            tables[group] = new float[5][];
            for (int alphabet = 0; alphabet < 5; alphabet++) {
                int[] counts = groups[group].Alphabet(alphabet);
                long total = 0;
                foreach (int count in counts) total += count;
                var table = new float[counts.Length];
                double missing = Log2(total + 2) + 4;
                for (int symbol = 0; symbol < counts.Length; symbol++) {
                    table[symbol] = counts[symbol] == 0 ? (float)missing : (float)Log2((double)total / counts[symbol]);
                }
                tables[group][alphabet] = table;
            }
        }
        var moved = new int[tiles.Length];
        for (int tile = 0; tile < tiles.Length; tile++) {
            if ((tile & 63) == 0) cancellationToken.ThrowIfCancellationRequested();
            Vp8lHistogram histogram = tiles[tile];
            int best = tileGroups[tile];
            if (histogram.IsEmpty()) {
                moved[tile] = best;
                continue;
            }
            double bestBits = double.MaxValue;
            for (int group = 0; group < groupCount; group++) {
                double bits = 0;
                for (int alphabet = 0; alphabet < 5; alphabet++) {
                    int[] counts = histogram.Alphabet(alphabet);
                    float[] table = tables[group][alphabet];
                    for (int symbol = 0; symbol < counts.Length; symbol++) {
                        if (counts[symbol] != 0) bits += counts[symbol] * table[symbol];
                    }
                }
                if (bits < bestBits) {
                    bestBits = bits;
                    best = group;
                }
            }
            moved[tile] = best;
        }

        var rebuilt = new Vp8lHistogram[groupCount];
        for (int group = 0; group < groupCount; group++) rebuilt[group] = new Vp8lHistogram(cacheBits);
        for (int tile = 0; tile < tiles.Length; tile++) rebuilt[moved[tile]].Add(tiles[tile]);
        var renumber = new int[groupCount];
        var kept = new List<Vp8lHistogram>();
        double before = 0, after = 0;
        foreach (Vp8lHistogram group in groups) before += group.Cost;
        for (int group = 0; group < groupCount; group++) {
            renumber[group] = -1;
            if (rebuilt[group].IsEmpty()) continue;
            rebuilt[group].Cost = rebuilt[group].EstimateBits();
            after += rebuilt[group].Cost;
            renumber[group] = kept.Count;
            kept.Add(rebuilt[group]);
        }
        if (after >= before || kept.Count == 0) return;
        for (int tile = 0; tile < tiles.Length; tile++) {
            int group = renumber[moved[tile]];
            tileGroups[tile] = group < 0 ? 0 : group;
        }
        groups = kept.ToArray();
    }
}
