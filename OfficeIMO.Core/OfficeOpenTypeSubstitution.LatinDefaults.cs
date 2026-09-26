using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeOpenTypeSubstitution {
    // Latin defaults must come from the selected Script/LangSys, not unrelated Arabic features.
    internal IReadOnlyList<string>? GetLatinDefaultFeatureTags() {
        int[]? indexes = GetLatinDefaultFeatureIndexes(out _);
        if (indexes == null) return null;
        var tags = new List<string>(indexes.Length);
        foreach (int index in indexes) tags.Add(ReadTag(_featureList + 2 + index * 6));
        return tags;
    }

    private int[]? GetLatinDefaultFeatureIndexes(out int requiredFeature) {
        requiredFeature = -1;
        try {
            int scriptList = Relative(_table, _reader.ReadUInt16(_table + 4), 2);
            int count = _reader.ReadUInt16(scriptList);
            if (count > MaximumFeatureRecords) return null;
            Ensure(scriptList + 2, checked(count * 6));
            int script = 0;
            for (int index = 0; index < count; index++) {
                int record = scriptList + 2 + index * 6;
                string tag = ReadTag(record);
                if (tag == "DFLT" && script == 0 || tag == "latn") {
                    script = Relative(scriptList, _reader.ReadUInt16(record + 4), 4);
                    if (tag == "latn") break;
                }
            }
            if (script == 0 || _reader.ReadUInt16(script) == 0) return Array.Empty<int>();
            int language = Relative(script, _reader.ReadUInt16(script), 6);
            int featureCount = _reader.ReadUInt16(language + 4);
            if (featureCount > MaximumFeatureRecords) return null;
            Ensure(language + 6, checked(featureCount * 2));
            Ensure(_featureList, 2);
            int totalFeatures = _reader.ReadUInt16(_featureList);
            if (totalFeatures > MaximumFeatureRecords) return null;
            Ensure(_featureList + 2, checked(totalFeatures * 6));
            var features = new SortedSet<int>();
            int required = _reader.ReadUInt16(language + 2);
            if (required != ushort.MaxValue) { features.Add(required); requiredFeature = required; }
            for (int index = 0; index < featureCount; index++) features.Add(_reader.ReadUInt16(language + 6 + index * 2));
            var result = new int[features.Count];
            int output = 0;
            foreach (int feature in features) {
                if (feature >= totalFeatures) return null;
                result[output++] = feature;
            }
            return result;
        } catch (Exception exception) when (exception is InvalidDataException || exception is OverflowException ||
            exception is ArgumentOutOfRangeException || exception is IndexOutOfRangeException) { return null; }
    }

    internal bool ApplyLatinDefaults(List<GlyphToken> glyphs, OfficeTextFeatureSettings settings, CancellationToken cancellationToken) {
        try { return ApplyLatinDefaultsCore(glyphs, settings, cancellationToken); }
        catch (Exception exception) when (exception is InvalidDataException || exception is OverflowException ||
            exception is ArgumentOutOfRangeException || exception is IndexOutOfRangeException) { return false; }
    }

    private static bool IsLatinDefaultScalar(int scalar) =>
        scalar <= 0x024F || scalar >= 0x0300 && scalar <= 0x036F ||
        scalar >= 0x1AB0 && scalar <= 0x1AFF || scalar >= 0x1DC0 && scalar <= 0x1DFF ||
        scalar >= 0x20D0 && scalar <= 0x20FF || scalar >= 0xFE20 && scalar <= 0xFE2F ||
        scalar >= 0x1E00 && scalar <= 0x1EFF || scalar >= 0x2C60 && scalar <= 0x2C7F ||
        scalar >= 0xA720 && scalar <= 0xA7FF || scalar >= 0xAB30 && scalar <= 0xAB6F ||
        scalar >= 0x10780 && scalar <= 0x107BF || scalar >= 0x1DF00 && scalar <= 0x1DFFF;

    private bool ApplyLatinDefaultsCore(List<GlyphToken> glyphs, OfficeTextFeatureSettings settings, CancellationToken cancellationToken) {
        int[]? features = GetLatinDefaultFeatureIndexes(out int requiredFeature);
        if (features == null) return false;
        var lookups = new SortedDictionary<int, int>();
        int inspections = 0;
        foreach (int index in features) {
            cancellationToken.ThrowIfCancellationRequested();
            int record = _featureList + 2 + index * 6;
            string tag = ReadTag(record);
            int setting = index == requiredFeature ? 1 : settings.TryGetValue(tag, out int explicitValue) ? explicitValue
                : tag == "liga" || tag == "clig" || tag == "rlig" ? 1 : 0;
            if (setting <= 0) continue;
            int feature = Relative(_featureList, _reader.ReadUInt16(record + 4), 4);
            int count = _reader.ReadUInt16(feature + 2);
            if (count > MaximumLookupRecords) return false;
            Ensure(feature + 4, checked(count * 2));
            for (int lookup = 0; lookup < count; lookup++) {
                int lookupIndex = _reader.ReadUInt16(feature + 4 + lookup * 2);
                if (!CanApplyLookup(lookupIndex, 0, ref inspections)) return false;
                lookups[lookupIndex] = setting;
            }
        }
        int operations = 0;
        // Script-specific lookups must not consume neighboring non-Latin or presentation glyphs.
        var shaped = new List<GlyphToken>(glyphs.Count);
        for (int index = 0; index < glyphs.Count;) {
            if (!IsLatinDefaultScalar(glyphs[index].Scalar)) { shaped.Add(glyphs[index++]); continue; }
            var segment = new List<GlyphToken>();
            do { segment.Add(glyphs[index++]); }
            while (index < glyphs.Count && IsLatinDefaultScalar(glyphs[index].Scalar));
            foreach (var lookup in lookups) ApplyLookup(segment, lookup.Key, lookup.Value, cancellationToken, ref operations);
            shaped.AddRange(segment);
        }
        glyphs.Clear(); glyphs.AddRange(shaped);
        return true;
    }
}
