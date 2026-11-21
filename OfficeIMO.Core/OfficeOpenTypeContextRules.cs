namespace OfficeIMO.Drawing;

/// <summary>Reads GSUB 5/6 and GPOS 7/8 rules in all three formats, sharing one matcher between both tables.</summary>
/// <remarks>
/// The caller owns glyph skipping: it supplies the glyph at each input, backtrack and lookahead
/// position after its lookup flags have been applied, or -1 when no such glyph exists. Reads go
/// through the caller's bounded accessor so table limits and work budgets stay with the owner.
/// </remarks>
internal sealed class OfficeOpenTypeContextRules {
    internal const int MaximumContextGlyphs = 256;
    private const int MaximumRules = 4096;
    private readonly Func<int, int> _u16;
    private readonly Func<int, int, int> _coverage;
    private readonly Action _charge;

    /// <param name="u16">Bounds-checked big-endian read of an unsigned 16-bit value.</param>
    /// <param name="coverage">Coverage index of a glyph in a coverage table, or -1.</param>
    /// <param name="charge">Charges one unit of the owner's work budget.</param>
    internal OfficeOpenTypeContextRules(Func<int, int> u16, Func<int, int, int> coverage, Action charge) {
        _u16 = u16; _coverage = coverage; _charge = charge;
    }

    /// <summary>A matched rule: its lookup records and the number of input glyphs it consumed.</summary>
    internal readonly struct Match {
        internal Match(int records, int recordCount, int inputCount) {
            Records = records; RecordCount = recordCount; InputCount = inputCount;
        }
        internal int Records { get; }
        internal int RecordCount { get; }
        internal int InputCount { get; }
    }

    /// <summary>Finds the first rule matching at the current glyph.</summary>
    /// <param name="subtable">Start of the context subtable.</param>
    /// <param name="chained">True for GSUB 6 and GPOS 8.</param>
    /// <param name="input">Glyph at input position k (0 is the current glyph), or -1.</param>
    /// <param name="backtrack">Glyph at backtrack position k (0 precedes the input), or -1.</param>
    /// <param name="lookahead">Glyph at lookahead position k after an input of the given length, or -1.</param>
    /// <param name="match">The matched rule's lookup records and input length.</param>
    internal bool TryMatch(int subtable, bool chained, Func<int, int> input, Func<int, int> backtrack,
        Func<int, int, int> lookahead, out Match match) {
        match = default;
        int format = _u16(subtable), first = input(0);
        if (first < 0) return false;
        if (format == 3) return chained ? TryChainedFormat3(subtable, input, backtrack, lookahead, out match)
            : TryFormat3(subtable, input, out match);
        if (format != 1 && format != 2) throw new InvalidDataException("Unsupported OpenType context format.");
        int coverageIndex = _coverage(Offset(subtable, _u16(subtable + 2)), first);
        if (coverageIndex < 0) return false;
        int inputClasses = 0, backtrackClasses = 0, lookaheadClasses = 0, setCountAt;
        if (format == 1) setCountAt = subtable + 4;
        else if (!chained) { inputClasses = OptionalOffset(subtable, _u16(subtable + 4)); setCountAt = subtable + 6; }
        else {
            backtrackClasses = OptionalOffset(subtable, _u16(subtable + 4));
            inputClasses = OptionalOffset(subtable, _u16(subtable + 6));
            lookaheadClasses = OptionalOffset(subtable, _u16(subtable + 8));
            setCountAt = subtable + 10;
        }
        int setCount = _u16(setCountAt);
        int setIndex = format == 1 ? coverageIndex : ClassOf(inputClasses, first);
        if (setIndex >= setCount) return false;
        int setOffset = _u16(setCountAt + 2 + setIndex * 2);
        if (setOffset == 0) return false;
        int set = Offset(subtable, setOffset), ruleCount = Count(set, MaximumRules);
        for (int rule = 0; rule < ruleCount; rule++) {
            _charge();
            int ruleTable = Offset(set, _u16(set + 2 + rule * 2));
            Func<int, int, bool> matchesInput = format == 1 ? (glyph, value) => glyph == value
                : (glyph, value) => ClassOf(inputClasses, glyph) == value;
            if (!chained) {
                int glyphCount = Count(ruleTable, MaximumContextGlyphs), recordCount = _u16(ruleTable + 2);
                if (glyphCount == 0) throw new InvalidDataException("An empty OpenType input context is invalid.");
                if (!MatchesSequence(ruleTable + 4, 1, glyphCount, input, matchesInput)) continue;
                match = new Match(ruleTable + 4 + (glyphCount - 1) * 2, recordCount, glyphCount);
                return true;
            }
            int cursor = ruleTable;
            int backtrackCount = Count(cursor, MaximumContextGlyphs);
            Func<int, int, bool> matchesBacktrack = format == 1 ? (glyph, value) => glyph == value
                : (glyph, value) => ClassOf(backtrackClasses, glyph) == value;
            if (!MatchesSequence(cursor + 2, 0, backtrackCount, backtrack, matchesBacktrack)) continue;
            cursor = checked(cursor + 2 + backtrackCount * 2);
            int inputCount = Count(cursor, MaximumContextGlyphs);
            if (inputCount == 0) throw new InvalidDataException("An empty OpenType input context is invalid.");
            if (!MatchesSequence(cursor + 2, 1, inputCount, input, matchesInput)) continue;
            cursor = checked(cursor + 2 + (inputCount - 1) * 2);
            int lookaheadCount = Count(cursor, MaximumContextGlyphs);
            Func<int, int, bool> matchesLookahead = format == 1 ? (glyph, value) => glyph == value
                : (glyph, value) => ClassOf(lookaheadClasses, glyph) == value;
            if (!MatchesSequence(cursor + 2, 0, lookaheadCount, k => lookahead(inputCount, k), matchesLookahead)) continue;
            cursor = checked(cursor + 2 + lookaheadCount * 2);
            match = new Match(cursor + 2, _u16(cursor), inputCount);
            return true;
        }
        return false;
    }

    /// <summary>Lists every rule's lookup records so preflight can validate nested lookups.</summary>
    internal List<Match> EnumerateRules(int subtable, bool chained) {
        var result = new List<Match>();
        int format = _u16(subtable);
        if (format == 3) {
            if (!chained) {
                int glyphCount = Count(subtable + 2, MaximumContextGlyphs);
                result.Add(new Match(subtable + 6 + glyphCount * 2, _u16(subtable + 4), glyphCount));
                return result;
            }
            int cursor = subtable + 2;
            cursor = checked(cursor + 2 + Count(cursor, MaximumContextGlyphs) * 2);
            int inputCount = Count(cursor, MaximumContextGlyphs);
            cursor = checked(cursor + 2 + inputCount * 2);
            cursor = checked(cursor + 2 + Count(cursor, MaximumContextGlyphs) * 2);
            result.Add(new Match(cursor + 2, _u16(cursor), inputCount));
            return result;
        }
        if (format != 1 && format != 2) throw new InvalidDataException("Unsupported OpenType context format.");
        _ = Offset(subtable, _u16(subtable + 2));
        int setCountAt = format == 1 ? subtable + 4 : chained ? subtable + 10 : subtable + 6;
        int setCount = Count(setCountAt, MaximumRules);
        for (int setIndex = 0; setIndex < setCount; setIndex++) {
            _charge();
            int setOffset = _u16(setCountAt + 2 + setIndex * 2);
            if (setOffset == 0) continue;
            int set = Offset(subtable, setOffset), ruleCount = Count(set, MaximumRules);
            for (int rule = 0; rule < ruleCount; rule++) {
                _charge();
                int ruleTable = Offset(set, _u16(set + 2 + rule * 2));
                if (!chained) {
                    int glyphCount = Count(ruleTable, MaximumContextGlyphs);
                    result.Add(new Match(ruleTable + 4 + Math.Max(0, glyphCount - 1) * 2, _u16(ruleTable + 2), glyphCount));
                    continue;
                }
                int cursor = ruleTable;
                cursor = checked(cursor + 2 + Count(cursor, MaximumContextGlyphs) * 2);
                int inputCount = Count(cursor, MaximumContextGlyphs);
                cursor = checked(cursor + 2 + Math.Max(0, inputCount - 1) * 2);
                cursor = checked(cursor + 2 + Count(cursor, MaximumContextGlyphs) * 2);
                result.Add(new Match(cursor + 2, _u16(cursor), inputCount));
            }
        }
        return result;
    }

    private bool TryFormat3(int subtable, Func<int, int> input, out Match match) {
        match = default;
        int glyphCount = Count(subtable + 2, MaximumContextGlyphs), recordCount = _u16(subtable + 4);
        if (glyphCount == 0) throw new InvalidDataException("An empty OpenType input context is invalid.");
        if (!MatchesCoverages(subtable, subtable + 6, glyphCount, input)) return false;
        match = new Match(subtable + 6 + glyphCount * 2, recordCount, glyphCount);
        return true;
    }

    private bool TryChainedFormat3(int subtable, Func<int, int> input, Func<int, int> backtrack,
        Func<int, int, int> lookahead, out Match match) {
        match = default;
        int cursor = subtable + 2;
        int backtrackCount = Count(cursor, MaximumContextGlyphs);
        int backtrackCoverages = cursor + 2;
        cursor = checked(cursor + 2 + backtrackCount * 2);
        int inputCount = Count(cursor, MaximumContextGlyphs);
        if (inputCount == 0) throw new InvalidDataException("An empty OpenType input context is invalid.");
        int inputCoverages = cursor + 2;
        cursor = checked(cursor + 2 + inputCount * 2);
        int lookaheadCount = Count(cursor, MaximumContextGlyphs);
        int lookaheadCoverages = cursor + 2;
        cursor = checked(cursor + 2 + lookaheadCount * 2);
        if (!MatchesCoverages(subtable, inputCoverages, inputCount, input)
            || !MatchesCoverages(subtable, backtrackCoverages, backtrackCount, backtrack)
            || !MatchesCoverages(subtable, lookaheadCoverages, lookaheadCount, k => lookahead(inputCount, k))) return false;
        match = new Match(cursor + 2, _u16(cursor), inputCount);
        return true;
    }

    private bool MatchesCoverages(int subtable, int coverages, int count, Func<int, int> glyphAt) {
        for (int k = 0; k < count; k++) {
            _charge();
            int glyph = glyphAt(k);
            if (glyph < 0 || _coverage(Offset(subtable, _u16(coverages + k * 2)), glyph) < 0) return false;
        }
        return true;
    }

    // Values hold glyph IDs (format 1) or classes (format 2) for positions first..count-1.
    private bool MatchesSequence(int values, int first, int count, Func<int, int> glyphAt, Func<int, int, bool> matches) {
        for (int k = first; k < count; k++) {
            _charge();
            int glyph = glyphAt(k);
            if (glyph < 0 || !matches(glyph, _u16(values + (k - first) * 2))) return false;
        }
        return true;
    }

    private int ClassOf(int classDef, int glyph) {
        if (classDef == 0) return 0;
        int format = _u16(classDef);
        if (format == 1) {
            int start = _u16(classDef + 2), count = _u16(classDef + 4), index = glyph - start;
            return index < 0 || index >= count ? 0 : _u16(classDef + 6 + index * 2);
        }
        if (format != 2) throw new InvalidDataException("Unsupported OpenType class-definition format.");
        int ranges = _u16(classDef + 2), low = 0, high = ranges - 1;
        while (low <= high) {
            _charge();
            int middle = low + ((high - low) / 2), record = classDef + 4 + middle * 6;
            if (glyph < _u16(record)) high = middle - 1;
            else if (glyph > _u16(record + 2)) low = middle + 1;
            else return _u16(record + 4);
        }
        return 0;
    }

    private int Count(int offset, int maximum) {
        int count = _u16(offset);
        if (count > maximum) throw new InvalidDataException("An OpenType context exceeds the managed limit.");
        return count;
    }

    private int OptionalOffset(int origin, int relative) => relative == 0 ? 0 : Offset(origin, relative);

    private int Offset(int origin, int relative) {
        if (relative == 0) throw new InvalidDataException("An OpenType context offset is missing.");
        int result = checked(origin + relative);
        _ = _u16(result);
        return result;
    }
}
