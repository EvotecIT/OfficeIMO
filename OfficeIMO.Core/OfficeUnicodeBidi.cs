using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>
/// Unicode Bidirectional Algorithm (UAX #9) over Unicode scalars: paragraph level (P2–P3),
/// explicit embeddings, overrides and isolates (X1–X10), weak types (W1–W7), paired brackets
/// (N0), neutrals (N1–N2), implicit levels (I1–I2), line-level resets (L1) and reordering (L2).
/// </summary>
internal static class OfficeUnicodeBidi {
    internal const int MaximumDepth = 125;
    private const int MaximumBracketPairs = 63;

    /// <summary>Resolved embedding levels for one paragraph.</summary>
    internal sealed class Paragraph {
        internal Paragraph(int[] scalars, OfficeBidiClass[] classes, byte[] levels, bool[] removed, int paragraphLevel) {
            Scalars = scalars;
            OriginalClasses = classes;
            Levels = levels;
            Removed = removed;
            ParagraphLevel = paragraphLevel;
        }

        internal int[] Scalars { get; }
        internal OfficeBidiClass[] OriginalClasses { get; }
        /// <summary>Levels before L1. Characters removed by X9 take the level of the preceding character.</summary>
        internal byte[] Levels { get; }
        /// <summary>Characters removed by rule X9 (embedding and override controls, BN).</summary>
        internal bool[] Removed { get; }
        internal int ParagraphLevel { get; }

        /// <summary>Applies rule L1 to the line [start, end) and returns its levels.</summary>
        internal byte[] GetLineLevels(int start, int end) {
            var result = new byte[end - start];
            Array.Copy(Levels, start, result, 0, result.Length);
            bool trailing = true;
            for (int index = end - 1; index >= start; index--) {
                OfficeBidiClass type = OriginalClasses[index];
                if (type is OfficeBidiClass.B or OfficeBidiClass.S) {
                    result[index - start] = (byte)ParagraphLevel;
                    trailing = true;
                } else if (trailing && (IsWhitespaceForL1(type) || Removed[index])) {
                    result[index - start] = (byte)ParagraphLevel;
                } else {
                    trailing = false;
                }
            }
            return result;
        }
    }

    internal static OfficeBidiClass GetClass(int scalar) => OfficeUnicodeTextData.GetBidiClass(scalar);

    /// <summary>Resolves levels. <paramref name="paragraphLevel"/> is 0, 1 or null for P2–P3.</summary>
    internal static Paragraph Resolve(int[] scalars, int? paragraphLevel, CancellationToken cancellationToken = default) {
        if (scalars == null) throw new ArgumentNullException(nameof(scalars));
        int length = scalars.Length;
        var original = new OfficeBidiClass[length];
        for (int index = 0; index < length; index++) original[index] = GetClass(scalars[index]);
        int[] matchingPdi = FindMatchingPdis(original, cancellationToken, out int[] matchingInitiator);
        int baseLevel = paragraphLevel ?? FirstStrongLevel(original, 0, length, matchingPdi) ?? 0;

        var types = (OfficeBidiClass[])original.Clone();
        var levels = new byte[length];
        var removed = new bool[length];
        ApplyExplicitLevels(scalars, types, levels, removed, baseLevel, matchingPdi, cancellationToken);

        // sos/eos compare explicit embedding levels, so resolve from a snapshot taken before I1–I2.
        var explicitLevels = (byte[])levels.Clone();
        foreach (List<int> sequence in BuildIsolatingRunSequences(types, levels, removed, matchingPdi, matchingInitiator, baseLevel, cancellationToken)) {
            cancellationToken.ThrowIfCancellationRequested();
            ResolveSequence(sequence, scalars, original, types, explicitLevels, levels, removed, matchingPdi, baseLevel);
        }

        // Removed characters take the level of the preceding character (paragraph level at start).
        int previous = baseLevel;
        for (int index = 0; index < length; index++) {
            if (removed[index]) levels[index] = (byte)previous;
            else previous = levels[index];
        }
        return new Paragraph(scalars, original, levels, removed, baseLevel);
    }

    /// <summary>Rule L2: returns logical indexes in visual (left-to-right) order.</summary>
    internal static int[] ReorderLine(byte[] lineLevels) {
        int count = lineLevels.Length;
        var order = new int[count];
        for (int index = 0; index < count; index++) order[index] = index;
        if (count == 0) return order;
        int highest = 0, lowestOdd = int.MaxValue;
        foreach (byte level in lineLevels) {
            if (level > highest) highest = level;
            if ((level & 1) == 1 && level < lowestOdd) lowestOdd = level;
        }
        for (int level = highest; level >= lowestOdd && lowestOdd != int.MaxValue; level--) {
            int start = 0;
            while (start < count) {
                while (start < count && lineLevels[order[start]] < level) start++;
                int end = start;
                while (end < count && lineLevels[order[end]] >= level) end++;
                Array.Reverse(order, start, end - start);
                start = end;
            }
        }
        return order;
    }

    /// <summary>
    /// Reorders whole items (for example words) by the levels of their characters. Each item takes
    /// the lowest level among its characters, which keeps items intact under L2.
    /// </summary>
    internal static int[] ReorderItems(IReadOnlyList<int> itemLevels) {
        var levels = new byte[itemLevels.Count];
        for (int index = 0; index < levels.Length; index++) levels[index] = (byte)itemLevels[index];
        return ReorderLine(levels);
    }

    internal static bool IsRightToLeftClass(OfficeBidiClass type) =>
        type is OfficeBidiClass.R or OfficeBidiClass.AL or OfficeBidiClass.RLE or OfficeBidiClass.RLO or OfficeBidiClass.RLI;

    /// <summary>True when the text has a strong right-to-left character or an explicit RTL control.</summary>
    internal static bool ContainsRightToLeft(string? text) {
        if (string.IsNullOrEmpty(text)) return false;
        for (int index = 0; index < text!.Length; index++) {
            int scalar = text[index];
            if (char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) {
                scalar = char.ConvertToUtf32(text[index], text[index + 1]);
                index++;
            } else if (scalar < 0x0590) {
                continue;
            }
            OfficeBidiClass type = GetClass(scalar);
            if (IsRightToLeftClass(type) || type == OfficeBidiClass.AN) return true;
        }
        return false;
    }

    internal static int[] ToScalars(string text, out int[] utf16Starts) {
        var scalars = new List<int>(text.Length);
        var starts = new List<int>(text.Length);
        for (int index = 0; index < text.Length; index++) {
            starts.Add(index);
            char unit = text[index];
            if (char.IsHighSurrogate(unit) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) {
                scalars.Add(char.ConvertToUtf32(unit, text[index + 1]));
                index++;
            } else {
                scalars.Add(unit);
            }
        }
        utf16Starts = starts.ToArray();
        return scalars.ToArray();
    }

    private static bool IsWhitespaceForL1(OfficeBidiClass type) =>
        type is OfficeBidiClass.WS or OfficeBidiClass.FSI or OfficeBidiClass.LRI or OfficeBidiClass.RLI or OfficeBidiClass.PDI;

    private static bool IsIsolateInitiator(OfficeBidiClass type) =>
        type is OfficeBidiClass.LRI or OfficeBidiClass.RLI or OfficeBidiClass.FSI;

    private static bool IsRemovedByX9(OfficeBidiClass type) =>
        type is OfficeBidiClass.RLE or OfficeBidiClass.LRE or OfficeBidiClass.RLO or OfficeBidiClass.LRO or
            OfficeBidiClass.PDF or OfficeBidiClass.BN;

    private static int[] FindMatchingPdis(OfficeBidiClass[] types, CancellationToken cancellationToken, out int[] matchingInitiator) {
        var matching = new int[types.Length];
        matchingInitiator = new int[types.Length];
        for (int index = 0; index < types.Length; index++) {
            matching[index] = -1;
            matchingInitiator[index] = -1;
        }
        var open = new Stack<int>();
        for (int index = 0; index < types.Length; index++) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            OfficeBidiClass type = types[index];
            if (IsIsolateInitiator(type)) {
                open.Push(index);
            } else if (type == OfficeBidiClass.PDI && open.Count > 0) {
                int initiator = open.Pop();
                matching[initiator] = index;
                matchingInitiator[index] = initiator;
            } else if (type == OfficeBidiClass.B) {
                open.Clear();
            }
        }
        return matching;
    }

    private static int? FirstStrongLevel(OfficeBidiClass[] types, int start, int end, int[] matchingPdi) {
        for (int index = start; index < end; index++) {
            OfficeBidiClass type = types[index];
            if (type == OfficeBidiClass.L) return 0;
            if (type is OfficeBidiClass.R or OfficeBidiClass.AL) return 1;
            if (IsIsolateInitiator(type)) {
                if (matchingPdi[index] < 0) return null;
                index = matchingPdi[index];
            } else if (type == OfficeBidiClass.B) {
                return null;
            }
        }
        return null;
    }

    private readonly struct StatusEntry {
        internal StatusEntry(int level, OfficeBidiClass overrideType, bool isolate) {
            Level = level;
            Override = overrideType;
            Isolate = isolate;
        }
        internal int Level { get; }
        internal OfficeBidiClass Override { get; } // ON when no override
        internal bool Isolate { get; }
    }

    private static void ApplyExplicitLevels(int[] scalars, OfficeBidiClass[] types, byte[] levels, bool[] removed,
        int baseLevel, int[] matchingPdi, CancellationToken cancellationToken) {
        var stack = new Stack<StatusEntry>();
        stack.Push(new StatusEntry(baseLevel, OfficeBidiClass.ON, false));
        int overflowIsolates = 0, overflowEmbeddings = 0, validIsolates = 0;
        for (int index = 0; index < types.Length; index++) {
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            OfficeBidiClass type = types[index];
            StatusEntry top = stack.Peek();
            switch (type) {
                case OfficeBidiClass.RLE:
                case OfficeBidiClass.LRE:
                case OfficeBidiClass.RLO:
                case OfficeBidiClass.LRO: {
                    bool rtl = type is OfficeBidiClass.RLE or OfficeBidiClass.RLO;
                    int level = rtl ? (top.Level + 1) | 1 : (top.Level + 2) & ~1;
                    if (level <= MaximumDepth && overflowIsolates == 0 && overflowEmbeddings == 0) {
                        OfficeBidiClass overrideType = type == OfficeBidiClass.RLO ? OfficeBidiClass.R
                            : type == OfficeBidiClass.LRO ? OfficeBidiClass.L : OfficeBidiClass.ON;
                        stack.Push(new StatusEntry(level, overrideType, false));
                    } else if (overflowIsolates == 0) {
                        overflowEmbeddings++;
                    }
                    levels[index] = (byte)top.Level;
                    removed[index] = true;
                    break;
                }
                case OfficeBidiClass.RLI:
                case OfficeBidiClass.LRI:
                case OfficeBidiClass.FSI: {
                    levels[index] = (byte)top.Level;
                    if (top.Override != OfficeBidiClass.ON) types[index] = top.Override;
                    bool rtl = type == OfficeBidiClass.RLI;
                    if (type == OfficeBidiClass.FSI) {
                        int end = matchingPdi[index] >= 0 ? matchingPdi[index] : types.Length;
                        rtl = FirstStrongLevel(types, index + 1, end, matchingPdi) == 1;
                    }
                    int level = rtl ? (top.Level + 1) | 1 : (top.Level + 2) & ~1;
                    if (level <= MaximumDepth && overflowIsolates == 0 && overflowEmbeddings == 0) {
                        validIsolates++;
                        stack.Push(new StatusEntry(level, OfficeBidiClass.ON, true));
                    } else {
                        overflowIsolates++;
                    }
                    break;
                }
                case OfficeBidiClass.PDI: {
                    if (overflowIsolates > 0) {
                        overflowIsolates--;
                    } else if (validIsolates > 0) {
                        overflowEmbeddings = 0;
                        while (!stack.Peek().Isolate) stack.Pop();
                        stack.Pop();
                        validIsolates--;
                    }
                    top = stack.Peek();
                    levels[index] = (byte)top.Level;
                    if (top.Override != OfficeBidiClass.ON) types[index] = top.Override;
                    break;
                }
                case OfficeBidiClass.PDF:
                    if (overflowIsolates == 0) {
                        if (overflowEmbeddings > 0) overflowEmbeddings--;
                        else if (!top.Isolate && stack.Count >= 2) stack.Pop();
                    }
                    levels[index] = (byte)top.Level;
                    removed[index] = true;
                    break;
                case OfficeBidiClass.B:
                    levels[index] = (byte)baseLevel;
                    break;
                case OfficeBidiClass.BN:
                    levels[index] = (byte)top.Level;
                    removed[index] = true;
                    break;
                default:
                    levels[index] = (byte)top.Level;
                    if (top.Override != OfficeBidiClass.ON) types[index] = top.Override;
                    break;
            }
        }
    }

    private static List<List<int>> BuildIsolatingRunSequences(OfficeBidiClass[] types, byte[] levels, bool[] removed,
        int[] matchingPdi, int[] matchingInitiator, int baseLevel, CancellationToken cancellationToken) {
        // X10: level runs over characters that X9 retains.
        var runs = new List<List<int>>();
        List<int>? current = null;
        int currentLevel = -1;
        for (int index = 0; index < types.Length; index++) {
            if (removed[index]) continue;
            if (current == null || levels[index] != currentLevel) {
                current = new List<int>();
                runs.Add(current);
                currentLevel = levels[index];
            }
            current.Add(index);
        }

        var runOfStart = new Dictionary<int, int>();
        for (int runIndex = 0; runIndex < runs.Count; runIndex++) runOfStart[runs[runIndex][0]] = runIndex;
        var sequences = new List<List<int>>();
        for (int runIndex = 0; runIndex < runs.Count; runIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            List<int> run = runs[runIndex];
            int first = run[0];
            // A run starting with a PDI that matches an initiator continues that initiator's sequence.
            if (types[first] == OfficeBidiClass.PDI && matchingInitiator[first] >= 0 && !removed[matchingInitiator[first]]) continue;
            var sequence = new List<int>(run);
            int last = run[run.Count - 1];
            while (IsIsolateInitiator(types[last]) && matchingPdi[last] >= 0 &&
                   runOfStart.TryGetValue(matchingPdi[last], out int nextRun)) {
                List<int> continuation = runs[nextRun];
                sequence.AddRange(continuation);
                last = continuation[continuation.Count - 1];
            }
            sequences.Add(sequence);
        }
        return sequences;
    }

    private static void ResolveSequence(List<int> sequence, int[] scalars, OfficeBidiClass[] original, OfficeBidiClass[] types,
        byte[] explicitLevels, byte[] levels, bool[] removed, int[] matchingPdi, int baseLevel) {
        int count = sequence.Count;
        int level = explicitLevels[sequence[0]];
        int firstIndex = sequence[0];
        int lastIndex = sequence[count - 1];

        int previousLevel = baseLevel;
        for (int index = firstIndex - 1; index >= 0; index--) {
            if (removed[index]) continue;
            previousLevel = explicitLevels[index];
            break;
        }
        int nextLevel = baseLevel;
        if (!IsIsolateInitiator(original[lastIndex])) {
            for (int index = lastIndex + 1; index < types.Length; index++) {
                if (removed[index]) continue;
                nextLevel = explicitLevels[index];
                break;
            }
        }
        OfficeBidiClass sos = (Math.Max(previousLevel, level) & 1) == 1 ? OfficeBidiClass.R : OfficeBidiClass.L;
        OfficeBidiClass eos = (Math.Max(nextLevel, level) & 1) == 1 ? OfficeBidiClass.R : OfficeBidiClass.L;

        var t = new OfficeBidiClass[count];
        for (int index = 0; index < count; index++) t[index] = types[sequence[index]];

        // W1
        for (int index = 0; index < count; index++) {
            if (t[index] != OfficeBidiClass.NSM) continue;
            if (index == 0) t[index] = sos;
            else t[index] = IsIsolateInitiator(t[index - 1]) || t[index - 1] == OfficeBidiClass.PDI ? OfficeBidiClass.ON : t[index - 1];
        }
        // W2, W3
        OfficeBidiClass lastStrong = sos;
        for (int index = 0; index < count; index++) {
            OfficeBidiClass type = t[index];
            if (type is OfficeBidiClass.L or OfficeBidiClass.R or OfficeBidiClass.AL) lastStrong = type;
            else if (type == OfficeBidiClass.EN && lastStrong == OfficeBidiClass.AL) t[index] = OfficeBidiClass.AN;
        }
        for (int index = 0; index < count; index++) if (t[index] == OfficeBidiClass.AL) t[index] = OfficeBidiClass.R;
        // W4
        for (int index = 1; index < count - 1; index++) {
            if (t[index] == OfficeBidiClass.ES && t[index - 1] == OfficeBidiClass.EN && t[index + 1] == OfficeBidiClass.EN) {
                t[index] = OfficeBidiClass.EN;
            } else if (t[index] == OfficeBidiClass.CS && t[index - 1] == t[index + 1] &&
                       t[index - 1] is OfficeBidiClass.EN or OfficeBidiClass.AN) {
                t[index] = t[index - 1];
            }
        }
        // W5
        for (int index = 0; index < count; index++) {
            if (t[index] != OfficeBidiClass.ET) continue;
            int start = index, end = index;
            while (end < count && t[end] == OfficeBidiClass.ET) end++;
            bool adjacentEn = (start > 0 && t[start - 1] == OfficeBidiClass.EN) || (end < count && t[end] == OfficeBidiClass.EN);
            if (adjacentEn) for (int fill = start; fill < end; fill++) t[fill] = OfficeBidiClass.EN;
            index = end - 1;
        }
        // W6
        for (int index = 0; index < count; index++) {
            if (t[index] is OfficeBidiClass.ES or OfficeBidiClass.ET or OfficeBidiClass.CS) t[index] = OfficeBidiClass.ON;
        }
        // W7
        lastStrong = sos;
        for (int index = 0; index < count; index++) {
            OfficeBidiClass type = t[index];
            if (type is OfficeBidiClass.L or OfficeBidiClass.R) lastStrong = type;
            else if (type == OfficeBidiClass.EN && lastStrong == OfficeBidiClass.L) t[index] = OfficeBidiClass.L;
        }

        ResolvePairedBrackets(sequence, scalars, original, t, level, sos);
        ResolveNeutrals(t, level, sos, eos);

        // I1, I2
        for (int index = 0; index < count; index++) {
            int position = sequence[index];
            OfficeBidiClass type = t[index];
            if ((level & 1) == 0) {
                if (type == OfficeBidiClass.R) levels[position] = (byte)(level + 1);
                else if (type is OfficeBidiClass.AN or OfficeBidiClass.EN) levels[position] = (byte)(level + 2);
                else levels[position] = (byte)level;
            } else {
                levels[position] = type is OfficeBidiClass.L or OfficeBidiClass.EN or OfficeBidiClass.AN
                    ? (byte)(level + 1) : (byte)level;
            }
        }
    }

    private static bool IsNeutralOrIsolate(OfficeBidiClass type) =>
        type is OfficeBidiClass.B or OfficeBidiClass.S or OfficeBidiClass.WS or OfficeBidiClass.ON or
            OfficeBidiClass.LRI or OfficeBidiClass.RLI or OfficeBidiClass.FSI or OfficeBidiClass.PDI;

    private static OfficeBidiClass StrongForNeutrals(OfficeBidiClass type) =>
        type is OfficeBidiClass.EN or OfficeBidiClass.AN ? OfficeBidiClass.R : type;

    private static void ResolveNeutrals(OfficeBidiClass[] t, int level, OfficeBidiClass sos, OfficeBidiClass eos) {
        OfficeBidiClass embedding = (level & 1) == 1 ? OfficeBidiClass.R : OfficeBidiClass.L;
        int count = t.Length;
        for (int index = 0; index < count; index++) {
            if (!IsNeutralOrIsolate(t[index])) continue;
            int start = index, end = index;
            while (end < count && IsNeutralOrIsolate(t[end])) end++;
            OfficeBidiClass before = start == 0 ? sos : StrongForNeutrals(t[start - 1]);
            OfficeBidiClass after = end == count ? eos : StrongForNeutrals(t[end]);
            OfficeBidiClass resolved = before == after && before is OfficeBidiClass.L or OfficeBidiClass.R ? before : embedding;
            for (int fill = start; fill < end; fill++) t[fill] = resolved;
            index = end - 1;
        }
    }

    private static int CanonicalBracket(int scalar) => scalar switch {
        0x2329 => 0x3008,
        0x232A => 0x3009,
        _ => scalar
    };

    private static void ResolvePairedBrackets(List<int> sequence, int[] scalars, OfficeBidiClass[] original, OfficeBidiClass[] t,
        int level, OfficeBidiClass sos) {
        int count = sequence.Count;
        var pairs = new List<(int Open, int Close)>();
        var stack = new List<(int Pair, int Index)>();
        for (int index = 0; index < count; index++) {
            if (t[index] != OfficeBidiClass.ON) continue;
            int scalar = scalars[sequence[index]];
            if (!OfficeUnicodeTextData.TryGetPairedBracket(scalar, out int pair, out bool opening)) continue;
            if (opening) {
                if (stack.Count == MaximumBracketPairs) {
                    // UAX #9 BD16: overflow returns an empty pair list for the
                    // entire isolating run sequence, including previously found pairs.
                    return;
                }
                stack.Add((CanonicalBracket(pair), index));
            } else {
                int canonical = CanonicalBracket(scalar);
                for (int depth = stack.Count - 1; depth >= 0; depth--) {
                    if (stack[depth].Pair != canonical) continue;
                    pairs.Add((stack[depth].Index, index));
                    stack.RemoveRange(depth, stack.Count - depth);
                    break;
                }
            }
        }
        if (pairs.Count == 0) return;
        pairs.Sort(static (left, right) => left.Open.CompareTo(right.Open));

        OfficeBidiClass embedding = (level & 1) == 1 ? OfficeBidiClass.R : OfficeBidiClass.L;
        OfficeBidiClass opposite = embedding == OfficeBidiClass.L ? OfficeBidiClass.R : OfficeBidiClass.L;
        foreach ((int open, int close) in pairs) {
            bool foundEmbedding = false, foundOpposite = false;
            for (int index = open + 1; index < close; index++) {
                OfficeBidiClass strong = StrongForNeutrals(t[index]);
                if (strong == embedding) {
                    foundEmbedding = true;
                    break;
                }
                if (strong == opposite) foundOpposite = true;
            }
            OfficeBidiClass resolved;
            if (foundEmbedding) {
                resolved = embedding;
            } else if (foundOpposite) {
                OfficeBidiClass context = sos;
                for (int index = open - 1; index >= 0; index--) {
                    OfficeBidiClass strong = StrongForNeutrals(t[index]);
                    if (strong is OfficeBidiClass.L or OfficeBidiClass.R) {
                        context = strong;
                        break;
                    }
                }
                resolved = context == opposite ? opposite : embedding;
            } else {
                continue;
            }
            t[open] = resolved;
            t[close] = resolved;
            for (int index = open + 1; index < count && original[sequence[index]] == OfficeBidiClass.NSM; index++) t[index] = resolved;
            for (int index = close + 1; index < count && original[sequence[index]] == OfficeBidiClass.NSM; index++) t[index] = resolved;
        }
    }
}
