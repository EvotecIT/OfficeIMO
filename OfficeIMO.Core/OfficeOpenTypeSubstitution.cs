using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

/// <summary>Bounded GSUB single, alternate, ligature, and extension lookup interpreter.</summary>
internal sealed class OfficeOpenTypeSubstitution {
    private const int MaximumFeatureRecords = 4096;
    private const int MaximumLookupRecords = 4096;
    private const int MaximumSubtablesPerLookup = 256;
    private const int MaximumCoverageGlyphs = 65535;
    private const int MaximumOperations = 1_000_000;

    private readonly OfficeOpenTypeReader _reader;
    private readonly int _table;
    private readonly int _end;
    private readonly int _featureList;
    private readonly int _lookupList;

    private OfficeOpenTypeSubstitution(OfficeOpenTypeReader reader, int table, int length) {
        _reader = reader;
        _table = table;
        _end = checked(table + length);
        Ensure(table, 10);
        _featureList = Relative(table, reader.ReadUInt16(table + 6), 2);
        _lookupList = Relative(table, reader.ReadUInt16(table + 8), 2);
    }

    internal static OfficeOpenTypeSubstitution? TryCreate(byte[] data) {
        try {
            OfficeOpenTypeReader? reader = OfficeOpenTypeReader.TryCreate(data);
            if (reader == null || !reader.TryGetTable("GSUB", out int table, out int length) || length < 10) return null;
            return new OfficeOpenTypeSubstitution(reader, table, length);
        } catch (Exception exception) when (exception is InvalidDataException
                                            || exception is OverflowException
                                            || exception is ArgumentOutOfRangeException
                                            || exception is IndexOutOfRangeException) {
            return null;
        }
    }

    internal void Apply(List<GlyphToken> glyphs, OfficeTextFeatureSettings settings, System.Threading.CancellationToken cancellationToken) {
        if (glyphs.Count == 0 || settings.IsDefault) return;
        Ensure(_featureList, 2);
        int featureCount = _reader.ReadUInt16(_featureList);
        if (featureCount > MaximumFeatureRecords) throw new InvalidDataException("The GSUB feature list exceeds the managed shaping limit.");
        Ensure(_featureList + 2, checked(featureCount * 6));
        int operations = 0;
        for (int featureIndex = 0; featureIndex < featureCount; featureIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            int record = _featureList + 2 + featureIndex * 6;
            string tag = ReadTag(record);
            if (!settings.TryGetValue(tag, out int setting) || setting <= 0) continue;
            int feature = Relative(_featureList, _reader.ReadUInt16(record + 4), 4);
            Ensure(feature, 4);
            int lookupCount = _reader.ReadUInt16(feature + 2);
            if (lookupCount > MaximumLookupRecords) throw new InvalidDataException("The GSUB feature lookup list exceeds the managed shaping limit.");
            Ensure(feature + 4, checked(lookupCount * 2));
            for (int lookupIndex = 0; lookupIndex < lookupCount; lookupIndex++) {
                ApplyLookup(
                    glyphs,
                    _reader.ReadUInt16(feature + 4 + lookupIndex * 2),
                    setting,
                    cancellationToken,
                    ref operations);
            }
        }
    }

    private void ApplyLookup(
        List<GlyphToken> glyphs,
        int lookupIndex,
        int featureValue,
        System.Threading.CancellationToken cancellationToken,
        ref int operations) {
        Ensure(_lookupList, 2);
        int lookupCount = _reader.ReadUInt16(_lookupList);
        if (lookupCount > MaximumLookupRecords || lookupIndex < 0 || lookupIndex >= lookupCount) return;
        Ensure(_lookupList + 2, checked(lookupCount * 2));
        int lookup = Relative(_lookupList, _reader.ReadUInt16(_lookupList + 2 + lookupIndex * 2), 6);
        int lookupType = _reader.ReadUInt16(lookup);
        int subtableCount = _reader.ReadUInt16(lookup + 4);
        if (subtableCount > MaximumSubtablesPerLookup) throw new InvalidDataException("A GSUB lookup exceeds the managed subtable limit.");
        Ensure(lookup + 6, checked(subtableCount * 2));
        for (int subtableIndex = 0; subtableIndex < subtableCount; subtableIndex++) {
            int subtable = Relative(lookup, _reader.ReadUInt16(lookup + 6 + subtableIndex * 2), 2);
            ApplySubtable(glyphs, lookupType, subtable, featureValue, cancellationToken, ref operations, 0);
        }
    }

    private void ApplySubtable(
        List<GlyphToken> glyphs,
        int lookupType,
        int subtable,
        int featureValue,
        System.Threading.CancellationToken cancellationToken,
        ref int operations,
        int extensionDepth) {
        if (lookupType == 7) {
            if (extensionDepth > 0) return;
            Ensure(subtable, 8);
            if (_reader.ReadUInt16(subtable) != 1) return;
            int extendedType = _reader.ReadUInt16(subtable + 2);
            uint relative = _reader.ReadUInt32(subtable + 4);
            if (relative > int.MaxValue) return;
            ApplySubtable(glyphs, extendedType, Relative(subtable, (int)relative, 2), featureValue, cancellationToken, ref operations, extensionDepth + 1);
            return;
        }

        for (int glyphIndex = 0; glyphIndex < glyphs.Count; glyphIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            if (++operations > MaximumOperations) throw new InvalidDataException("GSUB shaping exceeded the managed operation budget.");
            if (lookupType == 1) ApplySingle(glyphs, glyphIndex, subtable);
            else if (lookupType == 3) ApplyAlternate(glyphs, glyphIndex, subtable, featureValue);
            else if (lookupType == 4 && ApplyLigature(glyphs, glyphIndex, subtable)) glyphIndex--;
        }
    }

    private void ApplySingle(List<GlyphToken> glyphs, int index, int subtable) {
        Ensure(subtable, 6);
        int format = _reader.ReadUInt16(subtable);
        int coverage = Relative(subtable, _reader.ReadUInt16(subtable + 2), 4);
        int coverageIndex = CoverageIndex(coverage, glyphs[index].GlyphId);
        if (coverageIndex < 0) return;
        int replacement;
        if (format == 1) {
            replacement = unchecked((ushort)(glyphs[index].GlyphId + _reader.ReadInt16(subtable + 4)));
        } else if (format == 2) {
            int glyphCount = _reader.ReadUInt16(subtable + 4);
            if (glyphCount > MaximumCoverageGlyphs || coverageIndex >= glyphCount) return;
            Ensure(subtable + 6, checked(glyphCount * 2));
            replacement = _reader.ReadUInt16(subtable + 6 + coverageIndex * 2);
        } else {
            return;
        }
        if (replacement > 0 && replacement < _reader.GlyphCount) glyphs[index] = glyphs[index].WithGlyph(replacement);
    }

    private void ApplyAlternate(List<GlyphToken> glyphs, int index, int subtable, int featureValue) {
        Ensure(subtable, 6);
        if (_reader.ReadUInt16(subtable) != 1) return;
        int coverage = Relative(subtable, _reader.ReadUInt16(subtable + 2), 4);
        int coverageIndex = CoverageIndex(coverage, glyphs[index].GlyphId);
        int setCount = _reader.ReadUInt16(subtable + 4);
        if (setCount > MaximumCoverageGlyphs || coverageIndex < 0 || coverageIndex >= setCount) return;
        Ensure(subtable + 6, checked(setCount * 2));
        int set = Relative(subtable, _reader.ReadUInt16(subtable + 6 + coverageIndex * 2), 2);
        int alternateCount = _reader.ReadUInt16(set);
        if (alternateCount <= 0 || alternateCount > MaximumCoverageGlyphs) return;
        Ensure(set + 2, checked(alternateCount * 2));
        int selected = Math.Min(Math.Max(1, featureValue), alternateCount) - 1;
        int replacement = _reader.ReadUInt16(set + 2 + selected * 2);
        if (replacement > 0 && replacement < _reader.GlyphCount) glyphs[index] = glyphs[index].WithGlyph(replacement);
    }

    private bool ApplyLigature(List<GlyphToken> glyphs, int index, int subtable) {
        Ensure(subtable, 6);
        if (_reader.ReadUInt16(subtable) != 1) return false;
        int coverage = Relative(subtable, _reader.ReadUInt16(subtable + 2), 4);
        int coverageIndex = CoverageIndex(coverage, glyphs[index].GlyphId);
        int setCount = _reader.ReadUInt16(subtable + 4);
        if (setCount > MaximumCoverageGlyphs || coverageIndex < 0 || coverageIndex >= setCount) return false;
        Ensure(subtable + 6, checked(setCount * 2));
        int set = Relative(subtable, _reader.ReadUInt16(subtable + 6 + coverageIndex * 2), 2);
        int ligatureCount = _reader.ReadUInt16(set);
        if (ligatureCount > MaximumCoverageGlyphs) return false;
        Ensure(set + 2, checked(ligatureCount * 2));
        int bestLigature = 0;
        int bestComponentCount = 0;
        for (int ligatureIndex = 0; ligatureIndex < ligatureCount; ligatureIndex++) {
            int ligature = Relative(set, _reader.ReadUInt16(set + 2 + ligatureIndex * 2), 4);
            int replacement = _reader.ReadUInt16(ligature);
            int componentCount = _reader.ReadUInt16(ligature + 2);
            if (replacement <= 0 || replacement >= _reader.GlyphCount || componentCount < 2 || componentCount > glyphs.Count - index || componentCount <= bestComponentCount) continue;
            Ensure(ligature + 4, checked((componentCount - 1) * 2));
            bool matches = true;
            for (int component = 1; component < componentCount; component++) {
                if (glyphs[index + component].GlyphId != _reader.ReadUInt16(ligature + 4 + (component - 1) * 2)) {
                    matches = false;
                    break;
                }
            }
            if (matches) {
                bestLigature = replacement;
                bestComponentCount = componentCount;
            }
        }
        if (bestComponentCount == 0) return false;
        string unicode = string.Empty;
        for (int component = 0; component < bestComponentCount; component++) unicode += glyphs[index + component].UnicodeText;
        GlyphToken first = glyphs[index];
        glyphs[index] = new GlyphToken(bestLigature, unicode, first.TextIndex, first.Scalar);
        glyphs.RemoveRange(index + 1, bestComponentCount - 1);
        return true;
    }

    private int CoverageIndex(int coverage, int glyphId) {
        Ensure(coverage, 4);
        int format = _reader.ReadUInt16(coverage);
        int count = _reader.ReadUInt16(coverage + 2);
        if (count > MaximumCoverageGlyphs) throw new InvalidDataException("A GSUB coverage table exceeds the managed glyph limit.");
        if (format == 1) {
            Ensure(coverage + 4, checked(count * 2));
            int low = 0;
            int high = count - 1;
            while (low <= high) {
                int middle = low + (high - low) / 2;
                int candidate = _reader.ReadUInt16(coverage + 4 + middle * 2);
                if (glyphId < candidate) high = middle - 1;
                else if (glyphId > candidate) low = middle + 1;
                else return middle;
            }
            return -1;
        }
        if (format == 2) {
            Ensure(coverage + 4, checked(count * 6));
            for (int rangeIndex = 0; rangeIndex < count; rangeIndex++) {
                int range = coverage + 4 + rangeIndex * 6;
                int start = _reader.ReadUInt16(range);
                int end = _reader.ReadUInt16(range + 2);
                if (glyphId >= start && glyphId <= end) return checked(_reader.ReadUInt16(range + 4) + glyphId - start);
            }
        }
        return -1;
    }

    private int Relative(int origin, int relative, int minimumLength) {
        int offset = checked(origin + relative);
        Ensure(offset, minimumLength);
        return offset;
    }

    private void Ensure(int offset, int length) {
        if (offset < _table || length < 0 || offset > _end - length) throw new InvalidDataException("The GSUB table is truncated.");
        _reader.EnsureAvailable(offset, length);
    }

    private string ReadTag(int offset) {
        Ensure(offset, 4);
        uint tag = _reader.ReadUInt32(offset);
        return new string(new[] {
            (char)((tag >> 24) & 0xFF),
            (char)((tag >> 16) & 0xFF),
            (char)((tag >> 8) & 0xFF),
            (char)(tag & 0xFF)
        });
    }

    internal readonly struct GlyphToken {
        internal GlyphToken(int glyphId, string unicodeText, int textIndex, int scalar) {
            GlyphId = glyphId;
            UnicodeText = unicodeText;
            TextIndex = textIndex;
            Scalar = scalar;
        }

        internal int GlyphId { get; }
        internal string UnicodeText { get; }
        internal int TextIndex { get; }
        internal int Scalar { get; }
        internal GlyphToken WithGlyph(int glyphId) => new GlyphToken(glyphId, UnicodeText, TextIndex, Scalar);
    }
}
