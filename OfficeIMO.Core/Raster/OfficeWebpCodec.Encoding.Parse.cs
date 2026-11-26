namespace OfficeIMO.Drawing;

public static partial class OfficeWebpCodec {
    // Inside a match at least this long, the next pixel reuses it (one shorter) instead of searching.
    private const int Vp8lReuseMatchLength = 32;
    // A copy is tried at every length up to this, and at its full length.
    private const int Vp8lParseShortLengths = 16;

    /// <summary>
    /// Cost-based parse. Each pixel's longest match is found once; a shortest-path pass then chooses
    /// literals, cache hits and copies by the bits each would take under a code fitted to the previous
    /// parse (the first one greedy). The colour cache receives every pixel however it was coded, so
    /// whether a pixel hits the cache does not depend on the parse.
    /// </summary>
    private static Vp8lTokens CreateVp8lCostedTokens(uint[] pixels, int width, int attempts, int maximumCacheBits,
        int passes, bool bothStarts, CancellationToken cancellationToken, out int cacheBits) {
        int count = pixels.Length;
        var matcher = new Vp8lMatcher(pixels, width, attempts);
        var lengths = new ushort[count];
        var distances = new int[count];
        int previousLength = 0, previousDistance = 0;
        for (int position = 0; position < count; position++) {
            if ((position & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            int length, distance;
            if (previousLength - 1 >= Vp8lReuseMatchLength) {
                length = previousLength - 1;
                distance = previousDistance;
            } else {
                length = matcher.Find(position, out distance);
            }
            lengths[position] = (ushort)length;
            distances[position] = distance;
            previousLength = length;
            previousDistance = distance;
            matcher.Insert(position);
        }
        matcher.ReleaseChains();

        Vp8lTokens? greedy = new Vp8lTokens(Math.Min(count, 1 << 16));
        for (int position = 0; position < count;) {
            int length = lengths[position];
            if (length >= 3) {
                greedy.AddCopy(length, matcher.DistanceCode(distances[position]));
                position += length;
            } else {
                greedy.AddLiteral(pixels[position++]);
            }
        }
        // The first code comes from the greedy parse or from coding every pixel as a literal, whichever
        // is cheaper. Starting from greedy copies on an image where they barely pay makes far copies look
        // cheap, and the parse keeps choosing them; starting from literals misses repeats that do pay.
        SelectVp8lCacheBits(greedy, pixels, maximumCacheBits, cancellationToken, out double greedyBits, onlyLargest: true);
        SelectVp8lCacheBits(null, pixels, maximumCacheBits, cancellationToken, out double literalBits, onlyLargest: true);
        Vp8lTokens? best = null;
        double bestBits = double.MaxValue;
        cacheBits = 0;
        Vp8lTokens?[] starts = bothStarts ? new[] { greedy, null } : new[] { greedyBits <= literalBits ? greedy : null };
        greedy = null;
        for (int startIndex = 0; startIndex < starts.Length; startIndex++) {
            Vp8lTokens? start = starts[startIndex];
            starts[startIndex] = null;
            Vp8lTokens? tokens = start;
            // Refitting pays when the first code came from greedy copies; a literal start is already close.
            int startPasses = start == null && !bothStarts ? 1 : passes;
            int passCacheBits = 0;
            for (int pass = 0; pass < startPasses; pass++) {
                passCacheBits = SelectVp8lCacheBits(tokens, pixels, maximumCacheBits, cancellationToken, out _);
                Vp8lEntropyPlan plan = CreateVp8lSingleGroupPlan(tokens, pixels, passCacheBits, cancellationToken);
                var model = new Vp8lCostModel(plan.Groups[0]);
                // The previous parse is no longer needed once its code is fitted.
                tokens = null;
                start = null;
                tokens = TraceVp8lCheapestParse(pixels, lengths, distances, matcher, passCacheBits, model, cancellationToken);
            }
            // The last parse was costed for this cache size, so the stream keeps it.
            if (!bothStarts) {
                cacheBits = passCacheBits;
                return tokens!;
            }
            int finalBits = SelectVp8lCacheBits(tokens, pixels, maximumCacheBits, cancellationToken, out double bits);
            if (bits < bestBits) {
                bestBits = bits;
                best = tokens;
                cacheBits = finalBits;
            }
        }
        return best!;
    }

    private static Vp8lTokens TraceVp8lCheapestParse(uint[] pixels, ushort[] lengths, int[] distances, Vp8lMatcher matcher,
        int cacheBits, Vp8lCostModel model, CancellationToken cancellationToken) {
        int count = pixels.Length;
        var costs = new double[count + 1];
        // steps[i] is the token that ends at pixel i: zero for a literal, otherwise the copy length.
        var steps = new ushort[count + 1];
        for (int index = 1; index <= count; index++) costs[index] = double.MaxValue;
        uint[] cache = cacheBits == 0 ? Array.Empty<uint>() : new uint[1 << cacheBits];
        int shift = 32 - cacheBits;
        for (int position = 0; position < count; position++) {
            if ((position & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            double cost = costs[position];
            uint color = pixels[position];
            double literal;
            if (cacheBits != 0) {
                int slot = (int)((0x1E35A7BDU * color) >> shift);
                literal = cache[slot] == color ? model.CacheCost(slot) : model.LiteralCost(color);
                cache[slot] = color;
            } else {
                literal = model.LiteralCost(color);
            }
            if (cost + literal < costs[position + 1]) {
                costs[position + 1] = cost + literal;
                steps[position + 1] = 0;
            }
            int length = lengths[position];
            if (length == 0) continue;
            double copy = cost + model.DistanceCost(matcher.DistanceCode(distances[position]));
            int shortest = Math.Min(length, Vp8lParseShortLengths);
            // Inside a long match carried over from the previous pixel, shorter copies were tried already.
            if (position > 0 && length >= Vp8lReuseMatchLength && lengths[position - 1] == length + 1 &&
                distances[position - 1] == distances[position]) shortest = 0;
            for (int candidate = 1; candidate <= shortest; candidate++) Relax(position + candidate, candidate);
            if (length > shortest) Relax(position + length, length);

            void Relax(int end, int candidate) {
                double total = copy + model.LengthCost(candidate);
                if (total < costs[end]) {
                    costs[end] = total;
                    steps[end] = (ushort)candidate;
                }
            }
        }

        int tokenCount = 0;
        for (int end = count; end > 0; tokenCount++) end -= steps[end] == 0 ? 1 : steps[end];
        // The path is walked from the end, so tokens are written back to front.
        var items = new ulong[tokenCount];
        int written = tokenCount;
        for (int end = count; end > 0;) {
            int step = steps[end];
            if (step == 0) {
                items[--written] = Vp8lTokens.Literal(pixels[end - 1]);
                end--;
            } else {
                int start = end - step;
                items[--written] = Vp8lTokens.Copy(step, matcher.DistanceCode(distances[start]));
                end = start;
            }
        }
        return new Vp8lTokens(items, tokenCount);
    }

    /// <summary>Bits per symbol under one histogram; unseen symbols cost more than any seen one.</summary>
    private sealed class Vp8lCostModel {
        private readonly float[] _green;
        private readonly float[] _red;
        private readonly float[] _blue;
        private readonly float[] _alpha;
        private readonly float[] _distance;
        private readonly float[] _length;

        internal Vp8lCostModel(Vp8lHistogram histogram) {
            _green = CreateCosts(histogram.Green);
            _red = CreateCosts(histogram.Red);
            _blue = CreateCosts(histogram.Blue);
            _alpha = CreateCosts(histogram.Alpha);
            _distance = CreateCosts(histogram.Distance);
            _length = new float[Vp8lMaximumCopyLength + 1];
            for (int length = 1; length <= Vp8lMaximumCopyLength; length++) {
                GetVp8lPrefix(length, Vp8lLengthPrefixCount, out int prefix, out int extra, out _);
                _length[length] = _green[256 + prefix] + extra;
            }
        }

        internal double LiteralCost(uint color) =>
            _green[(color >> 8) & 255] + _red[(color >> 16) & 255] + _blue[color & 255] + _alpha[color >> 24];

        internal double CacheCost(int slot) => _green[256 + Vp8lLengthPrefixCount + slot];

        internal double LengthCost(int length) => _length[length];

        internal double DistanceCost(int distanceCode) {
            GetVp8lPrefix(distanceCode, Vp8lDistancePrefixCount, out int prefix, out int extra, out _);
            return _distance[prefix] + extra;
        }

        private static float[] CreateCosts(int[] counts) {
            long total = 0;
            foreach (int count in counts) total += count;
            var costs = new float[counts.Length];
            double unseen = Log2(total + 2) + 2;
            for (int symbol = 0; symbol < counts.Length; symbol++) {
                costs[symbol] = counts[symbol] == 0 ? (float)unseen : (float)Log2((double)total / counts[symbol]);
            }
            return costs;
        }
    }
}
