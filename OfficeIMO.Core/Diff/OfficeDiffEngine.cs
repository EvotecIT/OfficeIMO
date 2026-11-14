using System.Diagnostics;

namespace OfficeIMO.Diff;

/// <summary>
/// Aligns two sequences of interned item ids and records the matching blocks.
/// Everything that is not inside a matching block is a deletion (old side) or an insertion (new side).
/// </summary>
internal sealed class OfficeDiffEngine {
    private const int MaxRecursionDepth = 4096;
    private const int HistogramMaxChain = 64;

    private readonly int[] _old;
    private readonly int[] _new;
    private readonly OfficeDiffAlgorithm _algorithm;
    private readonly long _maxWorkUnits;
    private readonly long _deadlineTicks;
    private readonly Stopwatch? _stopwatch;
    private readonly CancellationToken _cancellationToken;
    private readonly List<OfficeDiffMatch> _matches = new List<OfficeDiffMatch>();
    private long _work;

    private OfficeDiffEngine(int[] oldItems, int[] newItems, OfficeDiffOptions options, CancellationToken cancellationToken) {
        _old = oldItems;
        _new = newItems;
        _algorithm = options.Algorithm;
        _maxWorkUnits = options.MaxWorkUnits;
        _cancellationToken = cancellationToken;
        if (options.Timeout is TimeSpan timeout) {
            if (timeout < TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(options), "The diff timeout cannot be negative.");
            _stopwatch = Stopwatch.StartNew();
            _deadlineTicks = (long)(timeout.TotalSeconds * Stopwatch.Frequency);
        }
    }

    /// <summary>True when a work or time limit stopped the search, so the result may not be minimal.</summary>
    internal bool LimitReached { get; private set; }

    internal long WorkUnits => _work;

    internal static OfficeDiffEngine Run(int[] oldItems, int[] newItems, OfficeDiffOptions options, CancellationToken cancellationToken) {
        var engine = new OfficeDiffEngine(oldItems, newItems, options, cancellationToken);
        engine.Diff(0, oldItems.Length, 0, newItems.Length, 0);
        return engine;
    }

    internal IReadOnlyList<OfficeDiffMatch> Matches => _matches;

    private void AddMatch(int oldStart, int newStart, int length) {
        if (length <= 0) return;
        int last = _matches.Count - 1;
        if (last >= 0) {
            OfficeDiffMatch previous = _matches[last];
            if (previous.OldStart + previous.Length == oldStart && previous.NewStart + previous.Length == newStart) {
                _matches[last] = new OfficeDiffMatch(previous.OldStart, previous.NewStart, previous.Length + length);
                return;
            }
        }

        _matches.Add(new OfficeDiffMatch(oldStart, newStart, length));
    }

    private bool Exhausted() {
        if (LimitReached) return true;
        _cancellationToken.ThrowIfCancellationRequested();
        if (_maxWorkUnits > 0 && _work > _maxWorkUnits) {
            LimitReached = true;
        } else if (_stopwatch != null && _stopwatch.ElapsedTicks > _deadlineTicks) {
            LimitReached = true;
        }

        return LimitReached;
    }

    private void Diff(int oldStart, int oldEnd, int newStart, int newEnd, int depth) {
        int prefix = 0;
        while (oldStart + prefix < oldEnd && newStart + prefix < newEnd && _old[oldStart + prefix] == _new[newStart + prefix]) prefix++;
        AddMatch(oldStart, newStart, prefix);
        oldStart += prefix;
        newStart += prefix;

        int suffix = 0;
        while (oldStart < oldEnd - suffix && newStart < newEnd - suffix && _old[oldEnd - suffix - 1] == _new[newEnd - suffix - 1]) suffix++;
        oldEnd -= suffix;
        newEnd -= suffix;
        _work += prefix + suffix + 1;

        if (oldStart < oldEnd && newStart < newEnd) {
            if (depth >= MaxRecursionDepth || Exhausted()) {
                LimitReached = true;
            } else if (_algorithm == OfficeDiffAlgorithm.Patience) {
                Patience(oldStart, oldEnd, newStart, newEnd, depth);
            } else if (_algorithm == OfficeDiffAlgorithm.Histogram) {
                Histogram(oldStart, oldEnd, newStart, newEnd, depth);
            } else {
                Bisect(oldStart, oldEnd, newStart, newEnd, depth);
            }
        }

        AddMatch(oldEnd, newEnd, suffix);
    }

    // Myers' middle-snake search (An O(ND) Difference Algorithm, section 4b): walks forward and backward
    // D-paths until they overlap, then recurses on both halves, so memory stays linear in N + M.
    private void Bisect(int oldStart, int oldEnd, int newStart, int newEnd, int depth) {
        int n = oldEnd - oldStart;
        int m = newEnd - newStart;
        int maxD = (n + m + 1) / 2;
        int offset = maxD;
        int length = 2 * maxD + 2;
        int[] forward = new int[length];
        int[] backward = new int[length];
        for (int index = 0; index < length; index++) {
            forward[index] = -1;
            backward[index] = -1;
        }

        forward[offset + 1] = 0;
        backward[offset + 1] = 0;
        int delta = n - m;
        bool odd = (delta & 1) != 0;
        int forwardStart = 0, forwardEnd = 0, backwardStart = 0, backwardEnd = 0;
        _work += length;

        for (int d = 0; d < maxD; d++) {
            if (Exhausted()) break;

            for (int k = -d + forwardStart; k <= d - forwardEnd; k += 2) {
                int index = offset + k;
                int x = k == -d || (k != d && forward[index - 1] < forward[index + 1]) ? forward[index + 1] : forward[index - 1] + 1;
                int y = x - k;
                int start = x;
                while (x < n && y < m && _old[oldStart + x] == _new[newStart + y]) {
                    x++;
                    y++;
                }

                _work += 1 + x - start;
                forward[index] = x;
                if (x > n) {
                    forwardEnd += 2;
                } else if (y > m) {
                    forwardStart += 2;
                } else if (odd) {
                    int other = offset + delta - k;
                    if (other >= 0 && other < length && backward[other] != -1 && x >= n - backward[other]) {
                        Split(oldStart, oldEnd, newStart, newEnd, x, y, depth);
                        return;
                    }
                }
            }

            for (int k = -d + backwardStart; k <= d - backwardEnd; k += 2) {
                int index = offset + k;
                int x = k == -d || (k != d && backward[index - 1] < backward[index + 1]) ? backward[index + 1] : backward[index - 1] + 1;
                int y = x - k;
                int start = x;
                while (x < n && y < m && _old[oldEnd - x - 1] == _new[newEnd - y - 1]) {
                    x++;
                    y++;
                }

                _work += 1 + x - start;
                backward[index] = x;
                if (x > n) {
                    backwardEnd += 2;
                } else if (y > m) {
                    backwardStart += 2;
                } else if (!odd) {
                    int other = offset + delta - k;
                    if (other >= 0 && other < length && forward[other] != -1) {
                        int forwardX = forward[other];
                        int forwardY = offset + forwardX - other;
                        if (forwardX >= n - x) {
                            Split(oldStart, oldEnd, newStart, newEnd, forwardX, forwardY, depth);
                            return;
                        }
                    }
                }
            }
        }

        // Either the sides share nothing (the search ran out without overlap) or a limit stopped it.
        // Both cases leave the region as a whole deletion plus a whole insertion.
    }

    private void Split(int oldStart, int oldEnd, int newStart, int newEnd, int x, int y, int depth) {
        Diff(oldStart, oldStart + x, newStart, newStart + y, depth + 1);
        Diff(oldStart + x, oldEnd, newStart + y, newEnd, depth + 1);
    }

    private void Patience(int oldStart, int oldEnd, int newStart, int newEnd, int depth) {
        var counts = new Dictionary<int, PatienceCount>();
        for (int index = oldStart; index < oldEnd; index++) {
            counts.TryGetValue(_old[index], out PatienceCount count);
            count.Old++;
            count.OldIndex = index;
            counts[_old[index]] = count;
        }

        for (int index = newStart; index < newEnd; index++) {
            if (!counts.TryGetValue(_new[index], out PatienceCount count)) continue;
            count.New++;
            count.NewIndex = index;
            counts[_new[index]] = count;
        }

        _work += (oldEnd - oldStart) + (newEnd - newStart);
        var unique = new List<OfficeDiffMatch>();
        for (int index = oldStart; index < oldEnd; index++) {
            PatienceCount count = counts[_old[index]];
            if (count.Old == 1 && count.New == 1) unique.Add(new OfficeDiffMatch(index, count.NewIndex, 1));
        }

        if (unique.Count == 0) {
            Bisect(oldStart, oldEnd, newStart, newEnd, depth);
            return;
        }

        IReadOnlyList<OfficeDiffMatch> anchors = LongestIncreasingByNew(unique);
        int oldCursor = oldStart;
        int newCursor = newStart;
        foreach (OfficeDiffMatch anchor in anchors) {
            Diff(oldCursor, anchor.OldStart, newCursor, anchor.NewStart, depth + 1);
            AddMatch(anchor.OldStart, anchor.NewStart, 1);
            oldCursor = anchor.OldStart + 1;
            newCursor = anchor.NewStart + 1;
        }

        Diff(oldCursor, oldEnd, newCursor, newEnd, depth + 1);
    }

    private static IReadOnlyList<OfficeDiffMatch> LongestIncreasingByNew(List<OfficeDiffMatch> pairs) {
        // Patience sorting over the new-side indexes; pairs are already ordered by old-side index.
        var tails = new List<int>();
        var previous = new int[pairs.Count];
        for (int index = 0; index < pairs.Count; index++) {
            int value = pairs[index].NewStart;
            int low = 0, high = tails.Count;
            while (low < high) {
                int middle = (low + high) / 2;
                if (pairs[tails[middle]].NewStart < value) low = middle + 1;
                else high = middle;
            }

            previous[index] = low > 0 ? tails[low - 1] : -1;
            if (low == tails.Count) tails.Add(index);
            else tails[low] = index;
        }

        var result = new OfficeDiffMatch[tails.Count];
        int cursor = tails.Count == 0 ? -1 : tails[tails.Count - 1];
        for (int index = result.Length - 1; index >= 0; index--) {
            result[index] = pairs[cursor];
            cursor = previous[cursor];
        }

        return result;
    }

    private void Histogram(int oldStart, int oldEnd, int newStart, int newEnd, int depth) {
        var occurrences = new Dictionary<int, List<int>>();
        for (int index = oldStart; index < oldEnd; index++) {
            if (!occurrences.TryGetValue(_old[index], out List<int>? list)) {
                list = new List<int>();
                occurrences.Add(_old[index], list);
            }

            list.Add(index);
        }

        _work += oldEnd - oldStart;
        int bestOld = -1, bestNew = -1, bestLength = 0, bestCount = int.MaxValue;
        int newIndex = newStart;
        while (newIndex < newEnd) {
            // Like git's histogram diff, skip past the region just measured so long common runs are walked once, not once per item.
            int nextNew = newIndex + 1;
            if (!occurrences.TryGetValue(_new[newIndex], out List<int>? list) || list.Count > HistogramMaxChain || list.Count > bestCount) {
                newIndex = nextNew;
                continue;
            }

            foreach (int oldIndex in list) {
                int before = 0;
                while (oldIndex - before > oldStart && newIndex - before > newStart && _old[oldIndex - before - 1] == _new[newIndex - before - 1]) before++;
                int after = 1;
                while (oldIndex + after < oldEnd && newIndex + after < newEnd && _old[oldIndex + after] == _new[newIndex + after]) after++;
                int length = before + after;
                int rarest = list.Count;
                for (int offset = -before; offset < after; offset++) {
                    rarest = Math.Min(rarest, occurrences[_old[oldIndex + offset]].Count);
                }

                _work += length + 1;
                nextNew = Math.Max(nextNew, newIndex + after);
                if (rarest < bestCount || (rarest == bestCount && length > bestLength)) {
                    bestCount = rarest;
                    bestLength = length;
                    bestOld = oldIndex - before;
                    bestNew = newIndex - before;
                }
            }

            if (Exhausted()) return;
            newIndex = nextNew;
        }

        if (bestLength == 0) {
            Bisect(oldStart, oldEnd, newStart, newEnd, depth);
            return;
        }

        Diff(oldStart, bestOld, newStart, bestNew, depth + 1);
        AddMatch(bestOld, bestNew, bestLength);
        Diff(bestOld + bestLength, oldEnd, bestNew + bestLength, newEnd, depth + 1);
    }

    private struct PatienceCount {
        internal int Old;
        internal int New;
        internal int OldIndex;
        internal int NewIndex;
    }
}

/// <summary>A run of items that match on both sides.</summary>
internal readonly struct OfficeDiffMatch {
    internal OfficeDiffMatch(int oldStart, int newStart, int length) {
        OldStart = oldStart;
        NewStart = newStart;
        Length = length;
    }

    internal int OldStart { get; }
    internal int NewStart { get; }
    internal int Length { get; }
}
