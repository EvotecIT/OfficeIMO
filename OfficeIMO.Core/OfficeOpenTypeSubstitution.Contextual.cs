namespace OfficeIMO.Drawing;

internal sealed partial class OfficeOpenTypeSubstitution {
    private const int MaximumContextGlyphs = 256;
    private const int MaximumContextLookupRecords = 256;
    private const int MaximumLookupRecursion = 8;
    private const int MaximumPreflightInspections = 100_000;

    internal bool CanApply(OfficeTextFeatureSettings settings, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(settings);
        cancellationToken.ThrowIfCancellationRequested();
        try {
            return CanApplyCore(settings, cancellationToken);
        } catch (Exception exception) when (exception is InvalidDataException
                                            || exception is OverflowException
                                            || exception is ArgumentOutOfRangeException
                                            || exception is IndexOutOfRangeException) {
            return false;
        }
    }

    private bool CanApplyCore(OfficeTextFeatureSettings settings, CancellationToken cancellationToken) {
        Ensure(_featureList, 2);
        int featureCount = _reader.ReadUInt16(_featureList);
        if (featureCount > MaximumFeatureRecords) return false;
        Ensure(_featureList + 2, checked(featureCount * 6));
        int inspections = 0;
        for (int featureIndex = 0; featureIndex < featureCount; featureIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            int record = _featureList + 2 + featureIndex * 6;
            string tag = ReadTag(record);
            if (!settings.TryGetValue(tag, out int setting) || setting <= 0) continue;
            int feature = Relative(_featureList, _reader.ReadUInt16(record + 4), 4);
            int lookupCount = _reader.ReadUInt16(feature + 2);
            if (lookupCount > MaximumLookupRecords) return false;
            Ensure(feature + 4, checked(lookupCount * 2));
            for (int index = 0; index < lookupCount; index++) {
                if (!CanApplyLookup(_reader.ReadUInt16(feature + 4 + index * 2), 0, ref inspections, cancellationToken)) return false;
            }
        }
        return true;
    }

    private bool CanApplyLookup(int lookupIndex, int depth, ref int inspections, CancellationToken cancellationToken, bool allowGlyphCountChanges = true,
        OfficeOpenTypeGlyphFilter? glyphFilter = null, int inheritedFlags = -1, int inheritedSet = -1, MarkFilterUse? markFilters = null) {
        cancellationToken.ThrowIfCancellationRequested();
        if (depth >= MaximumLookupRecursion || ++inspections > MaximumPreflightInspections) return false;
        Ensure(_lookupList, 2);
        int lookupCount = _reader.ReadUInt16(_lookupList);
        if (lookupCount > MaximumLookupRecords || lookupIndex < 0 || lookupIndex >= lookupCount) return false;
        Ensure(_lookupList + 2, checked(lookupCount * 2));
        int lookup = Relative(_lookupList, _reader.ReadUInt16(_lookupList + 2 + lookupIndex * 2), 6);
        int lookupType = _reader.ReadUInt16(lookup);
        int lookupFlags = _reader.ReadUInt16(lookup + 2);
        if (lookupType < 1 || lookupType > 8) return false;
        int subtableCount = _reader.ReadUInt16(lookup + 4);
        if (subtableCount > MaximumSubtablesPerLookup) return false;
        Ensure(lookup + 6, checked(subtableCount * 2));
        int set = -1;
        if ((lookupFlags & 0x10) != 0) { Ensure(lookup + 6 + subtableCount * 2, 2); set = _reader.ReadUInt16(lookup + 6 + subtableCount * 2); }
        // The filtered run is the parent's view of the glyphs. A nested lookup whose own
        // flags differ can still act on its one target glyph: like the reference shaper,
        // a nested single, multiple or alternate substitution does not re-filter it.
        // A nested lookup that matches further glyphs would need its own view.
        bool ownView = glyphFilter != null && inheritedFlags >= 0
            && ((lookupFlags & 0xFF1E) != (inheritedFlags & 0xFF1E) || set != inheritedSet);
        if (glyphFilter == null && (lookupFlags & 0xFF1E) != 0) {
            // Flags that only skip marks change nothing in a run without marks. Latin
            // defaults accept them and verify each run before applying (MarkFilterUse).
            if (markFilters == null || (lookupFlags & 0x0006) != 0) return false;
            markFilters.Used = true;
        }
        for (int index = 0; index < subtableCount; index++) {
            if ((index & 63) == 0) cancellationToken.ThrowIfCancellationRequested();
            if (++inspections > MaximumPreflightInspections) return false;
            int subtable = Relative(lookup, _reader.ReadUInt16(lookup + 6 + index * 2), 2);
            int effectiveType = lookupType;
            if (lookupType == 7) {
                Ensure(subtable, 8);
                if (_reader.ReadUInt16(subtable) != 1) return false;
                effectiveType = _reader.ReadUInt16(subtable + 2);
                uint relative = _reader.ReadUInt32(subtable + 4);
                if (relative > int.MaxValue) return false;
                subtable = Relative(subtable, (int)relative, 2);
            }
            if (effectiveType < 1 || effectiveType > 8 || effectiveType == 7) return false;
            if (ownView && effectiveType > 3) return false;
            int format = _reader.ReadUInt16(subtable);
            if (effectiveType == 1 && format != 1 && format != 2 ||
                effectiveType is 2 or 3 or 4 && format != 1) return false;
            if (!allowGlyphCountChanges && (effectiveType == 2 || effectiveType == 4)) return false;
            if (effectiveType == 8 && (depth > 0 || subtableCount != 1)) return false;
            if ((effectiveType == 5 || effectiveType == 6) && format is < 1 or > 3) return false;
            if (effectiveType == 8 && _reader.ReadUInt16(subtable) != 1) return false;
            if ((effectiveType == 5 || effectiveType == 6)
                && !CanApplyContextLookupRecords(effectiveType, subtable, depth, ref inspections, cancellationToken, allowGlyphCountChanges,
                    glyphFilter, lookupFlags, set, markFilters)) return false;
        }
        return true;
    }

    /// <summary>Records whether an accepted lookup skips marks, so its caller must check runs for marks.</summary>
    internal sealed class MarkFilterUse {
        internal bool Used { get; set; }
    }

    private bool CanApplyContextLookupRecords(int lookupType, int subtable, int depth, ref int inspections, CancellationToken cancellationToken,
        bool allowGlyphCountChanges, OfficeOpenTypeGlyphFilter? glyphFilter, int inheritedFlags, int inheritedSet, MarkFilterUse? markFilters) {
        int counted = inspections;
        List<OfficeOpenTypeContextRules.Match> rules = ContextRules(() => {
            if (++counted > MaximumPreflightInspections) throw new InvalidDataException("GSUB preflight exceeded the managed inspection limit.");
        }).EnumerateRules(subtable, lookupType == 6);
        inspections = counted;
        foreach (OfficeOpenTypeContextRules.Match rule in rules) {
            if (rule.InputCount <= 0 || rule.RecordCount > MaximumContextLookupRecords) return false;
            Ensure(rule.Records, checked(rule.RecordCount * 4));
            for (int record = 0; record < rule.RecordCount; record++) {
                int sequenceIndex = _reader.ReadUInt16(rule.Records + record * 4);
                int nestedLookup = _reader.ReadUInt16(rule.Records + record * 4 + 2);
                // Later records address original input slots. Until slot tracking is supported,
                // decline any earlier record (including nested contexts) that can change their positions.
                if (sequenceIndex >= rule.InputCount || !CanApplyLookup(nestedLookup, depth + 1, ref inspections, cancellationToken,
                    allowGlyphCountChanges && record == rule.RecordCount - 1, glyphFilter, inheritedFlags, inheritedSet, markFilters)) return false;
            }
        }
        return true;
    }

    private OfficeOpenTypeContextRules ContextRules(Action charge) => new(
        offset => { Ensure(offset, 2); return _reader.ReadUInt16(offset); },
        (coverage, glyph) => { int operations = 0; return CoverageIndex(coverage, glyph, ref operations); },
        charge);

    private bool TryMatchContextRule(List<GlyphToken> glyphs, int index, int subtable, bool chained, ref int operations,
        out OfficeOpenTypeContextRules.Match match) {
        int counted = 0, budget = MaximumOperations - operations;
        bool matched = ContextRules(() => {
            if (++counted > budget) throw new InvalidDataException("GSUB shaping exceeded the managed operation budget.");
        }).TryMatch(subtable, chained,
            k => index + k < glyphs.Count ? glyphs[index + k].GlyphId : -1,
            k => index - 1 - k >= 0 ? glyphs[index - 1 - k].GlyphId : -1,
            (inputCount, k) => index + inputCount + k < glyphs.Count ? glyphs[index + inputCount + k].GlyphId : -1,
            out match);
        ChargeOperations(ref operations, counted);
        if (!matched) return false;
        if (match.RecordCount > MaximumContextLookupRecords) throw new InvalidDataException("A GSUB context rule exceeds the managed lookup-record limit.");
        Ensure(match.Records, checked(match.RecordCount * 4));
        return true;
    }

    private int ApplyMultiple(List<GlyphToken> glyphs, int index, int subtable, ref int operations) {
        Ensure(subtable, 6);
        if (_reader.ReadUInt16(subtable) != 1) return 0;
        int coverage = Relative(subtable, _reader.ReadUInt16(subtable + 2), 4);
        int coverageIndex = CoverageIndex(coverage, glyphs[index].GlyphId, ref operations);
        int sequenceCount = _reader.ReadUInt16(subtable + 4);
        if (sequenceCount > MaximumCoverageGlyphs || coverageIndex < 0 || coverageIndex >= sequenceCount) return 0;
        Ensure(subtable + 6, checked(sequenceCount * 2));
        int sequence = Relative(subtable, _reader.ReadUInt16(subtable + 6 + coverageIndex * 2), 2);
        int replacementCount = _reader.ReadUInt16(sequence);
        if (replacementCount <= 0 || replacementCount > MaximumContextGlyphs) return 0;
        Ensure(sequence + 2, checked(replacementCount * 2));
        if (glyphs.Count > MaximumOperations - replacementCount + 1)
            throw new InvalidDataException("GSUB shaping exceeded the managed glyph budget.");
        // InsertRange shifts the remaining tail. Charge that work as well as the
        // new tokens so a chain of multiple substitutions cannot amplify input.
        ChargeOperations(ref operations, checked(glyphs.Count - index + replacementCount));
        GlyphToken source = glyphs[index];
        var replacements = new GlyphToken[replacementCount];
        for (int replacementIndex = 0; replacementIndex < replacementCount; replacementIndex++) {
            int glyphId = _reader.ReadUInt16(sequence + 2 + replacementIndex * 2);
            if (glyphId <= 0 || glyphId >= _reader.GlyphCount) return 0;
            replacements[replacementIndex] = replacementIndex == 0
                ? source.WithGlyph(glyphId)
                : new GlyphToken(glyphId, string.Empty, source.TextIndex, source.Scalar, isUnicodeContinuation: true, clusterStart: source.ClusterStart,
                    ligatureComponentTextIndexes: source.LigatureComponentTextIndexes, ligatureIdentity: source.LigatureIdentity,
                    ligatureComponent: source.LigatureComponent, lookupOrder: source.LookupOrder, lookupIdentity: source.LookupIdentity);
        }
        glyphs.RemoveAt(index);
        glyphs.InsertRange(index, replacements);
        return replacementCount;
    }

    private int ApplyContextual(
        List<GlyphToken> glyphs,
        int index,
        int subtable,
        int featureValue,
        CancellationToken cancellationToken,
        ref int operations,
        int recursionDepth) {
        Ensure(subtable, 6);
        int format = _reader.ReadUInt16(subtable);
        if (format is 1 or 2) return ApplyContextRule(glyphs, index, subtable, chained: false, featureValue, cancellationToken, ref operations, recursionDepth);
        if (format != 3) return 0;
        int glyphCount = _reader.ReadUInt16(subtable + 2);
        int recordCount = _reader.ReadUInt16(subtable + 4);
        if (glyphCount <= 0 || glyphCount > MaximumContextGlyphs || recordCount > MaximumContextLookupRecords || index > glyphs.Count - glyphCount) return 0;
        Ensure(subtable + 6, checked(glyphCount * 2 + recordCount * 4));
        for (int input = 0; input < glyphCount; input++) {
            int coverage = Relative(subtable, _reader.ReadUInt16(subtable + 6 + input * 2), 4);
            if (CoverageIndex(coverage, glyphs[index + input].GlyphId, ref operations) < 0) return 0;
        }
        int countBefore = glyphs.Count;
        ApplyLookupRecords(glyphs, index, subtable + 6 + glyphCount * 2, recordCount, featureValue, cancellationToken, ref operations, recursionDepth);
        return Math.Max(1, glyphCount + glyphs.Count - countBefore);
    }

    private int ApplyChainedContextual(
        List<GlyphToken> glyphs,
        int index,
        int subtable,
        int featureValue,
        CancellationToken cancellationToken,
        ref int operations,
        int recursionDepth) {
        Ensure(subtable, 4);
        int format = _reader.ReadUInt16(subtable);
        if (format is 1 or 2) return ApplyContextRule(glyphs, index, subtable, chained: true, featureValue, cancellationToken, ref operations, recursionDepth);
        if (format != 3) return 0;
        int cursor = subtable + 2;
        int backtrackCount = ReadBoundedCount(cursor);
        cursor += 2;
        Ensure(cursor, checked(backtrackCount * 2 + 2));
        if (index < backtrackCount) return 0;
        for (int backtrack = 0; backtrack < backtrackCount; backtrack++) {
            int coverage = Relative(subtable, _reader.ReadUInt16(cursor + backtrack * 2), 4);
            if (CoverageIndex(coverage, glyphs[index - backtrack - 1].GlyphId, ref operations) < 0) return 0;
        }
        cursor += backtrackCount * 2;
        int inputCount = ReadBoundedCount(cursor);
        cursor += 2;
        Ensure(cursor, checked(inputCount * 2 + 2));
        if (inputCount <= 0 || index > glyphs.Count - inputCount) return 0;
        for (int input = 0; input < inputCount; input++) {
            int coverage = Relative(subtable, _reader.ReadUInt16(cursor + input * 2), 4);
            if (CoverageIndex(coverage, glyphs[index + input].GlyphId, ref operations) < 0) return 0;
        }
        cursor += inputCount * 2;
        int lookaheadCount = ReadBoundedCount(cursor);
        cursor += 2;
        Ensure(cursor, checked(lookaheadCount * 2 + 2));
        if (index + inputCount + lookaheadCount > glyphs.Count) return 0;
        for (int lookahead = 0; lookahead < lookaheadCount; lookahead++) {
            int coverage = Relative(subtable, _reader.ReadUInt16(cursor + lookahead * 2), 4);
            if (CoverageIndex(coverage, glyphs[index + inputCount + lookahead].GlyphId, ref operations) < 0) return 0;
        }
        cursor += lookaheadCount * 2;
        int recordCount = _reader.ReadUInt16(cursor);
        cursor += 2;
        if (recordCount > MaximumContextLookupRecords) throw new InvalidDataException("A GSUB chained-context rule exceeds the managed lookup-record limit.");
        Ensure(cursor, checked(recordCount * 4));
        int countBefore = glyphs.Count;
        ApplyLookupRecords(glyphs, index, cursor, recordCount, featureValue, cancellationToken, ref operations, recursionDepth);
        return Math.Max(1, inputCount + glyphs.Count - countBefore);
    }

    private int ApplyContextRule(List<GlyphToken> glyphs, int index, int subtable, bool chained, int featureValue,
        CancellationToken cancellationToken, ref int operations, int recursionDepth) {
        if (!TryMatchContextRule(glyphs, index, subtable, chained, ref operations, out OfficeOpenTypeContextRules.Match match)) return 0;
        int countBefore = glyphs.Count;
        ApplyLookupRecords(glyphs, index, match.Records, match.RecordCount, featureValue, cancellationToken, ref operations, recursionDepth);
        return Math.Max(1, match.InputCount + glyphs.Count - countBefore);
    }

    private void ApplyReverseChaining(
        List<GlyphToken> glyphs,
        int subtable,
        CancellationToken cancellationToken,
        ref int operations) {
        Ensure(subtable, 6);
        if (_reader.ReadUInt16(subtable) != 1) return;
        int coverage = Relative(subtable, _reader.ReadUInt16(subtable + 2), 4);
        int cursor = subtable + 4;
        int backtrackCount = ReadBoundedCount(cursor);
        cursor += 2;
        Ensure(cursor, checked(backtrackCount * 2 + 2));
        var backtrack = new int[backtrackCount];
        for (int i = 0; i < backtrackCount; i++) backtrack[i] = Relative(subtable, _reader.ReadUInt16(cursor + i * 2), 4);
        cursor += backtrackCount * 2;
        int lookaheadCount = ReadBoundedCount(cursor);
        cursor += 2;
        Ensure(cursor, checked(lookaheadCount * 2 + 2));
        var lookahead = new int[lookaheadCount];
        for (int i = 0; i < lookaheadCount; i++) lookahead[i] = Relative(subtable, _reader.ReadUInt16(cursor + i * 2), 4);
        cursor += lookaheadCount * 2;
        int substituteCount = _reader.ReadUInt16(cursor);
        cursor += 2;
        if (substituteCount > MaximumCoverageGlyphs) throw new InvalidDataException("A GSUB reverse substitution exceeds the managed glyph limit.");
        Ensure(cursor, checked(substituteCount * 2));
        for (int index = glyphs.Count - 1; index >= 0; index--) {
            cancellationToken.ThrowIfCancellationRequested();
            if (++operations > MaximumOperations) throw new InvalidDataException("GSUB shaping exceeded the managed operation budget.");
            int coverageIndex = CoverageIndex(coverage, glyphs[index].GlyphId, ref operations);
            if (coverageIndex < 0 || coverageIndex >= substituteCount || index < backtrackCount || index + 1 + lookaheadCount > glyphs.Count) continue;
            bool matches = true;
            for (int i = 0; i < backtrackCount && matches; i++) matches = CoverageIndex(backtrack[i], glyphs[index - i - 1].GlyphId, ref operations) >= 0;
            for (int i = 0; i < lookaheadCount && matches; i++) matches = CoverageIndex(lookahead[i], glyphs[index + i + 1].GlyphId, ref operations) >= 0;
            int replacement = _reader.ReadUInt16(cursor + coverageIndex * 2);
            if (matches && replacement > 0 && replacement < _reader.GlyphCount) glyphs[index] = glyphs[index].WithGlyph(replacement);
        }
    }

    private int ReadBoundedCount(int offset) {
        Ensure(offset, 2);
        int count = _reader.ReadUInt16(offset);
        if (count > MaximumContextGlyphs) throw new InvalidDataException("A GSUB context exceeds the managed glyph limit.");
        return count;
    }

    private void ApplyLookupRecords(
        List<GlyphToken> glyphs,
        int start,
        int records,
        int recordCount,
        int featureValue,
        CancellationToken cancellationToken,
        ref int operations,
        int recursionDepth) {
        if (recursionDepth >= MaximumLookupRecursion) throw new InvalidDataException("GSUB contextual lookup recursion exceeded the managed limit.");
        for (int record = 0; record < recordCount; record++) {
            cancellationToken.ThrowIfCancellationRequested();
            int sequenceIndex = _reader.ReadUInt16(records + record * 4);
            int lookupIndex = _reader.ReadUInt16(records + record * 4 + 2);
            int target = start + sequenceIndex;
            if (target < 0 || target >= glyphs.Count) continue;
            ApplyLookupAt(glyphs, target, lookupIndex, featureValue, cancellationToken, ref operations, recursionDepth + 1);
        }
    }

    private int ApplyLookupAt(
        List<GlyphToken> glyphs,
        int glyphIndex,
        int lookupIndex,
        int featureValue,
        CancellationToken cancellationToken,
        ref int operations,
        int recursionDepth) {
        Ensure(_lookupList, 2);
        int lookupCount = _reader.ReadUInt16(_lookupList);
        if (lookupIndex < 0 || lookupIndex >= lookupCount || glyphIndex < 0 || glyphIndex >= glyphs.Count) return 0;
        if (lookupCount > MaximumLookupRecords) throw new InvalidDataException("A GSUB lookup list exceeds the managed lookup limit.");
        Ensure(_lookupList + 2, checked(lookupCount * 2));
        int lookup = Relative(_lookupList, _reader.ReadUInt16(_lookupList + 2 + lookupIndex * 2), 6);
        int lookupType = _reader.ReadUInt16(lookup);
        int subtableCount = _reader.ReadUInt16(lookup + 4);
        if (subtableCount > MaximumSubtablesPerLookup) throw new InvalidDataException("A GSUB lookup exceeds the managed subtable limit.");
        Ensure(lookup + 6, checked(subtableCount * 2));
        for (int subtableIndex = 0; subtableIndex < subtableCount; subtableIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (++operations > MaximumOperations) throw new InvalidDataException("GSUB shaping exceeded the managed operation budget.");
            int subtable = Relative(lookup, _reader.ReadUInt16(lookup + 6 + subtableIndex * 2), 2);
            if (lookupType == 7) {
                Ensure(subtable, 8);
                if (_reader.ReadUInt16(subtable) != 1) continue;
                int extendedType = _reader.ReadUInt16(subtable + 2);
                uint relative = _reader.ReadUInt32(subtable + 4);
                if (relative > int.MaxValue) continue;
                int advance = ApplySubtableAt(glyphs, glyphIndex, extendedType, Relative(subtable, (int)relative, 2), featureValue, cancellationToken, ref operations, recursionDepth);
                if (advance > 0) return advance;
            } else {
                int advance = ApplySubtableAt(glyphs, glyphIndex, lookupType, subtable, featureValue, cancellationToken, ref operations, recursionDepth);
                if (advance > 0) return advance;
            }
        }
        return 0;
    }

    private int ApplySubtableAt(
        List<GlyphToken> glyphs, int index, int lookupType, int subtable, int featureValue,
        CancellationToken cancellationToken, ref int operations, int recursionDepth) {
        if (index < 0 || index >= glyphs.Count) return 0;
        if (lookupType == 1) return ApplySingle(glyphs, index, subtable, ref operations) ? 1 : 0;
        if (lookupType == 2) return ApplyMultiple(glyphs, index, subtable, ref operations);
        if (lookupType == 3) return ApplyAlternate(glyphs, index, subtable, featureValue, ref operations) ? 1 : 0;
        if (lookupType == 4) return ApplyLigature(glyphs, index, subtable, ref operations) ? 1 : 0;
        if (lookupType == 5) return ApplyContextual(glyphs, index, subtable, featureValue, cancellationToken, ref operations, recursionDepth);
        if (lookupType == 6) return ApplyChainedContextual(glyphs, index, subtable, featureValue, cancellationToken, ref operations, recursionDepth);
        return 0;
    }
}
