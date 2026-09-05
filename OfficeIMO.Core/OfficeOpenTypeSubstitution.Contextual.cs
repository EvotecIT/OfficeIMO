using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeOpenTypeSubstitution {
    private const int MaximumContextGlyphs = 256;
    private const int MaximumContextLookupRecords = 256;
    private const int MaximumLookupRecursion = 8;

    internal bool CanApply(OfficeTextFeatureSettings settings) {
        if (settings == null) throw new ArgumentNullException(nameof(settings));
        try {
            return CanApplyCore(settings);
        } catch (Exception exception) when (exception is InvalidDataException
                                            || exception is OverflowException
                                            || exception is ArgumentOutOfRangeException
                                            || exception is IndexOutOfRangeException) {
            return false;
        }
    }

    private bool CanApplyCore(OfficeTextFeatureSettings settings) {
        Ensure(_featureList, 2);
        int featureCount = _reader.ReadUInt16(_featureList);
        if (featureCount > MaximumFeatureRecords) return false;
        Ensure(_featureList + 2, checked(featureCount * 6));
        for (int featureIndex = 0; featureIndex < featureCount; featureIndex++) {
            int record = _featureList + 2 + featureIndex * 6;
            string tag = ReadTag(record);
            if (!settings.TryGetValue(tag, out int setting) || setting <= 0) continue;
            int feature = Relative(_featureList, _reader.ReadUInt16(record + 4), 4);
            int lookupCount = _reader.ReadUInt16(feature + 2);
            if (lookupCount > MaximumLookupRecords) return false;
            Ensure(feature + 4, checked(lookupCount * 2));
            for (int index = 0; index < lookupCount; index++) {
                if (!CanApplyLookup(_reader.ReadUInt16(feature + 4 + index * 2), 0)) return false;
            }
        }
        return true;
    }

    private bool CanApplyLookup(int lookupIndex, int depth) {
        if (depth >= MaximumLookupRecursion) return false;
        Ensure(_lookupList, 2);
        int lookupCount = _reader.ReadUInt16(_lookupList);
        if (lookupIndex < 0 || lookupIndex >= lookupCount) return false;
        Ensure(_lookupList + 2, checked(lookupCount * 2));
        int lookup = Relative(_lookupList, _reader.ReadUInt16(_lookupList + 2 + lookupIndex * 2), 6);
        int lookupType = _reader.ReadUInt16(lookup);
        int lookupFlags = _reader.ReadUInt16(lookup + 2);
        if ((lookupFlags & 0xFF1E) != 0 || lookupType < 1 || lookupType > 8) return false;
        int subtableCount = _reader.ReadUInt16(lookup + 4);
        if (subtableCount > MaximumSubtablesPerLookup) return false;
        Ensure(lookup + 6, checked(subtableCount * 2));
        for (int index = 0; index < subtableCount; index++) {
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
            if (depth > 0 && effectiveType == 8) return false;
            if ((effectiveType == 5 || effectiveType == 6) && _reader.ReadUInt16(subtable) != 3) return false;
            if (effectiveType == 8 && _reader.ReadUInt16(subtable) != 1) return false;
            if ((effectiveType == 5 || effectiveType == 6)
                && !CanApplyContextLookupRecords(effectiveType, subtable, depth)) return false;
        }
        return true;
    }

    private bool CanApplyContextLookupRecords(int lookupType, int subtable, int depth) {
        int records;
        int recordCount;
        int inputGlyphCount;
        if (lookupType == 5) {
            Ensure(subtable, 6);
            int glyphCount = _reader.ReadUInt16(subtable + 2);
            recordCount = _reader.ReadUInt16(subtable + 4);
            if (glyphCount <= 0 || glyphCount > MaximumContextGlyphs || recordCount > MaximumContextLookupRecords) return false;
            inputGlyphCount = glyphCount;
            records = checked(subtable + 6 + glyphCount * 2);
            Ensure(subtable + 6, checked(glyphCount * 2 + recordCount * 4));
        } else {
            Ensure(subtable, 4);
            int cursor = subtable + 2;
            int backtrackCount = ReadBoundedCount(cursor);
            cursor = checked(cursor + 2 + backtrackCount * 2);
            Ensure(cursor, 2);
            int inputCount = ReadBoundedCount(cursor);
            if (inputCount <= 0) return false;
            inputGlyphCount = inputCount;
            cursor = checked(cursor + 2 + inputCount * 2);
            Ensure(cursor, 2);
            int lookaheadCount = ReadBoundedCount(cursor);
            cursor = checked(cursor + 2 + lookaheadCount * 2);
            Ensure(cursor, 2);
            recordCount = _reader.ReadUInt16(cursor);
            if (recordCount > MaximumContextLookupRecords) return false;
            records = cursor + 2;
            Ensure(records, checked(recordCount * 4));
        }

        for (int record = 0; record < recordCount; record++) {
            int sequenceIndex = _reader.ReadUInt16(records + record * 4);
            int nestedLookup = _reader.ReadUInt16(records + record * 4 + 2);
            if (sequenceIndex >= inputGlyphCount || !CanApplyLookup(nestedLookup, depth + 1)) return false;
        }
        return true;
    }

    private int ApplyMultiple(List<GlyphToken> glyphs, int index, int subtable) {
        Ensure(subtable, 6);
        if (_reader.ReadUInt16(subtable) != 1) return 0;
        int coverage = Relative(subtable, _reader.ReadUInt16(subtable + 2), 4);
        int coverageIndex = CoverageIndex(coverage, glyphs[index].GlyphId);
        int sequenceCount = _reader.ReadUInt16(subtable + 4);
        if (sequenceCount > MaximumCoverageGlyphs || coverageIndex < 0 || coverageIndex >= sequenceCount) return 0;
        Ensure(subtable + 6, checked(sequenceCount * 2));
        int sequence = Relative(subtable, _reader.ReadUInt16(subtable + 6 + coverageIndex * 2), 2);
        int replacementCount = _reader.ReadUInt16(sequence);
        if (replacementCount <= 0 || replacementCount > MaximumContextGlyphs) return 0;
        Ensure(sequence + 2, checked(replacementCount * 2));
        GlyphToken source = glyphs[index];
        var replacements = new GlyphToken[replacementCount];
        for (int replacementIndex = 0; replacementIndex < replacementCount; replacementIndex++) {
            int glyphId = _reader.ReadUInt16(sequence + 2 + replacementIndex * 2);
            if (glyphId <= 0 || glyphId >= _reader.GlyphCount) return 0;
            replacements[replacementIndex] = replacementIndex == 0
                ? source.WithGlyph(glyphId)
                : new GlyphToken(glyphId, string.Empty, source.TextIndex, source.Scalar, isUnicodeContinuation: true);
        }
        glyphs.RemoveAt(index);
        glyphs.InsertRange(index, replacements);
        return replacementCount;
    }

    private void ApplyContextual(
        List<GlyphToken> glyphs,
        int index,
        int subtable,
        int featureValue,
        CancellationToken cancellationToken,
        ref int operations,
        int recursionDepth) {
        Ensure(subtable, 6);
        int format = _reader.ReadUInt16(subtable);
        if (format != 3) return;
        int glyphCount = _reader.ReadUInt16(subtable + 2);
        int recordCount = _reader.ReadUInt16(subtable + 4);
        if (glyphCount <= 0 || glyphCount > MaximumContextGlyphs || recordCount > MaximumContextLookupRecords || index > glyphs.Count - glyphCount) return;
        Ensure(subtable + 6, checked(glyphCount * 2 + recordCount * 4));
        for (int input = 0; input < glyphCount; input++) {
            int coverage = Relative(subtable, _reader.ReadUInt16(subtable + 6 + input * 2), 4);
            if (CoverageIndex(coverage, glyphs[index + input].GlyphId) < 0) return;
        }
        ApplyLookupRecords(glyphs, index, subtable + 6 + glyphCount * 2, recordCount, featureValue, cancellationToken, ref operations, recursionDepth);
    }

    private void ApplyChainedContextual(
        List<GlyphToken> glyphs,
        int index,
        int subtable,
        int featureValue,
        CancellationToken cancellationToken,
        ref int operations,
        int recursionDepth) {
        Ensure(subtable, 4);
        if (_reader.ReadUInt16(subtable) != 3) return;
        int cursor = subtable + 2;
        int backtrackCount = ReadBoundedCount(cursor);
        cursor += 2;
        Ensure(cursor, checked(backtrackCount * 2 + 2));
        if (index < backtrackCount) return;
        for (int backtrack = 0; backtrack < backtrackCount; backtrack++) {
            int coverage = Relative(subtable, _reader.ReadUInt16(cursor + backtrack * 2), 4);
            if (CoverageIndex(coverage, glyphs[index - backtrack - 1].GlyphId) < 0) return;
        }
        cursor += backtrackCount * 2;
        int inputCount = ReadBoundedCount(cursor);
        cursor += 2;
        Ensure(cursor, checked(inputCount * 2 + 2));
        if (inputCount <= 0 || index > glyphs.Count - inputCount) return;
        for (int input = 0; input < inputCount; input++) {
            int coverage = Relative(subtable, _reader.ReadUInt16(cursor + input * 2), 4);
            if (CoverageIndex(coverage, glyphs[index + input].GlyphId) < 0) return;
        }
        cursor += inputCount * 2;
        int lookaheadCount = ReadBoundedCount(cursor);
        cursor += 2;
        Ensure(cursor, checked(lookaheadCount * 2 + 2));
        if (index + inputCount + lookaheadCount > glyphs.Count) return;
        for (int lookahead = 0; lookahead < lookaheadCount; lookahead++) {
            int coverage = Relative(subtable, _reader.ReadUInt16(cursor + lookahead * 2), 4);
            if (CoverageIndex(coverage, glyphs[index + inputCount + lookahead].GlyphId) < 0) return;
        }
        cursor += lookaheadCount * 2;
        int recordCount = _reader.ReadUInt16(cursor);
        cursor += 2;
        if (recordCount > MaximumContextLookupRecords) throw new InvalidDataException("A GSUB chained-context rule exceeds the managed lookup-record limit.");
        Ensure(cursor, checked(recordCount * 4));
        ApplyLookupRecords(glyphs, index, cursor, recordCount, featureValue, cancellationToken, ref operations, recursionDepth);
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
            int coverageIndex = CoverageIndex(coverage, glyphs[index].GlyphId);
            if (coverageIndex < 0 || coverageIndex >= substituteCount || index < backtrackCount || index + 1 + lookaheadCount > glyphs.Count) continue;
            bool matches = true;
            for (int i = 0; i < backtrackCount && matches; i++) matches = CoverageIndex(backtrack[i], glyphs[index - i - 1].GlyphId) >= 0;
            for (int i = 0; i < lookaheadCount && matches; i++) matches = CoverageIndex(lookahead[i], glyphs[index + i + 1].GlyphId) >= 0;
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

    private void ApplyLookupAt(
        List<GlyphToken> glyphs,
        int glyphIndex,
        int lookupIndex,
        int featureValue,
        CancellationToken cancellationToken,
        ref int operations,
        int recursionDepth) {
        Ensure(_lookupList, 2);
        int lookupCount = _reader.ReadUInt16(_lookupList);
        if (lookupIndex < 0 || lookupIndex >= lookupCount || glyphIndex < 0 || glyphIndex >= glyphs.Count) return;
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
                ApplySubtableAt(glyphs, glyphIndex, extendedType, Relative(subtable, (int)relative, 2), featureValue, cancellationToken, ref operations, recursionDepth);
            } else {
                ApplySubtableAt(glyphs, glyphIndex, lookupType, subtable, featureValue, cancellationToken, ref operations, recursionDepth);
            }
        }
    }

    private void ApplySubtableAt(
        List<GlyphToken> glyphs,
        int index,
        int lookupType,
        int subtable,
        int featureValue,
        CancellationToken cancellationToken,
        ref int operations,
        int recursionDepth) {
        if (index < 0 || index >= glyphs.Count) return;
        if (lookupType == 1) ApplySingle(glyphs, index, subtable);
        else if (lookupType == 2) ApplyMultiple(glyphs, index, subtable);
        else if (lookupType == 3) ApplyAlternate(glyphs, index, subtable, featureValue);
        else if (lookupType == 4) ApplyLigature(glyphs, index, subtable);
        else if (lookupType == 5) ApplyContextual(glyphs, index, subtable, featureValue, cancellationToken, ref operations, recursionDepth);
        else if (lookupType == 6) ApplyChainedContextual(glyphs, index, subtable, featureValue, cancellationToken, ref operations, recursionDepth);
    }
}
