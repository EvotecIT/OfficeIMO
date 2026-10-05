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
        // The selected default Script/LangSys belongs to immutable font data.
        // Resolve it once per cached font, including malformed-table failures.
        var features = _latinDefaultFeatures.Value;
        requiredFeature = features.Required;
        return features.Indexes;
    }

    private int[]? ReadLatinDefaultFeatureIndexes(out int requiredFeature) {
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

    internal bool ApplyLatinDefaults(List<GlyphToken> glyphs, OfficeTextFeatureSettings settings, CancellationToken cancellationToken, string? sourceText = null) {
        try { return ApplyLatinDefaultsCore(glyphs, settings, cancellationToken, sourceText); }
        catch (Exception exception) when (exception is InvalidDataException || exception is OverflowException ||
            exception is ArgumentOutOfRangeException || exception is IndexOutOfRangeException) { return false; }
    }

    private bool ApplyLatinDefaultsCore(List<GlyphToken> glyphs, OfficeTextFeatureSettings settings, CancellationToken cancellationToken, string? sourceText) {
        bool? asciiResult = TryApplyAsciiLatinDefaults(glyphs, settings, cancellationToken, sourceText);
        if (asciiResult.HasValue) return asciiResult.Value;
        var scalars = new int[glyphs.Count];
        var breakBefore = new bool[glyphs.Count];
        for (int index = 0; index < glyphs.Count; index++) {
            scalars[index] = glyphs[index].Scalar;
            // Non-painting controls removed by the provider still delimit source segments.
            breakBefore[index] = index == 0 ? glyphs[index].TextIndex > 0 : glyphs[index].TextIndex >
                glyphs[index - 1].TextIndex + glyphs[index - 1].UnicodeText.Length;
        }
        bool trailingBoundary = glyphs.Count > 0 && sourceText != null && sourceText.Length >
            glyphs[glyphs.Count - 1].TextIndex + glyphs[glyphs.Count - 1].UnicodeText.Length;
        bool[] eligible = GetLatinDefaultEligibility(scalars, breakBefore, trailingBoundary);
        // Empty input is also used to preflight the font's selected lookups.
        if (glyphs.Count > 0 && !Array.Exists(eligible, value => value)) return true;
        KeyValuePair<int, int>[]? lookups = settings.IsDefault ? GetLatinDefaultLookups(cancellationToken) : BuildLatinDefaultLookups(settings, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        if (lookups == null) return false;
        if (lookups.Length == 0) return true;
        int operations = 0;
        // Script-specific lookups must not consume neighboring non-Latin or presentation glyphs.
        var shaped = new List<GlyphToken>(glyphs.Count);
        for (int index = 0; index < glyphs.Count;) {
            if (!eligible[index]) { shaped.Add(glyphs[index++]); continue; }
            var segment = new List<GlyphToken>();
            do { segment.Add(glyphs[index++]); }
            while (index < glyphs.Count && eligible[index] && !breakBefore[index]);
            foreach (var lookup in lookups) ApplyLookup(segment, lookup.Key, lookup.Value, cancellationToken, ref operations);
            shaped.AddRange(segment);
        }
        glyphs.Clear(); glyphs.AddRange(shaped);
        return true;
    }

    private bool? TryApplyAsciiLatinDefaults(List<GlyphToken> glyphs, OfficeTextFeatureSettings settings,
        CancellationToken cancellationToken, string? sourceText) {
        if (glyphs.Count == 0 || sourceText == null || sourceText.Length != glyphs.Count) return null;
        bool hasLatin = false;
        for (int index = 0; index < glyphs.Count; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            GlyphToken glyph = glyphs[index];
            if ((uint)(glyph.Scalar - 32) > 94U || glyph.TextIndex != index || glyph.UnicodeText.Length != 1) return null;
            hasLatin |= glyph.Scalar >= 'A' && glyph.Scalar <= 'Z' || glyph.Scalar >= 'a' && glyph.Scalar <= 'z';
        }
        if (!hasLatin) return true;
        KeyValuePair<int, int>[]? lookups = settings.IsDefault ? GetLatinDefaultLookups(cancellationToken) : BuildLatinDefaultLookups(settings, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        if (lookups == null) return false;
        if (lookups.Length == 0) return true;
        // A single uninterrupted ASCII segment has the same eligibility as the general
        // script resolver. Keep its working copy so a rejected operation leaves input intact.
        var segment = new List<GlyphToken>(glyphs);
        int operations = 0;
        foreach (var lookup in lookups) ApplyLookup(segment, lookup.Key, lookup.Value, cancellationToken, ref operations);
        glyphs.Clear(); glyphs.AddRange(segment);
        return true;
    }

    internal bool CanApplyLatinDefaults(OfficeTextFeatureSettings settings, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            bool supported = (settings.IsDefault ? GetLatinDefaultLookups(cancellationToken) : BuildLatinDefaultLookups(settings, cancellationToken)) != null;
            cancellationToken.ThrowIfCancellationRequested();
            return supported;
        } catch (Exception exception) when (exception is InvalidDataException || exception is OverflowException ||
            exception is ArgumentOutOfRangeException || exception is IndexOutOfRangeException) { return false; }
    }

    private KeyValuePair<int, int>[]? GetLatinDefaultLookups(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        LatinDefaultLookupCache? cached = Volatile.Read(ref _latinDefaultLookups);
        if (cached != null) return cached.Lookups;
        KeyValuePair<int, int>[]? lookups;
        try { lookups = BuildLatinDefaultLookups(OfficeTextFeatureSettings.Default, cancellationToken); }
        catch (Exception exception) when (exception is InvalidDataException || exception is OverflowException ||
            exception is ArgumentOutOfRangeException || exception is IndexOutOfRangeException) { lookups = null; }
        cancellationToken.ThrowIfCancellationRequested();
        var completed = new LatinDefaultLookupCache(lookups);
        // Publish completed immutable results only. A cancelled caller does not
        // poison initialization or make another caller wait for its work.
        return (Interlocked.CompareExchange(ref _latinDefaultLookups, completed, null) ?? completed).Lookups;
    }

    private KeyValuePair<int, int>[]? BuildLatinDefaultLookups(OfficeTextFeatureSettings settings, CancellationToken cancellationToken) {
        int[]? features = GetLatinDefaultFeatureIndexes(out int requiredFeature);
        if (features == null) return null;
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
            if (count > MaximumLookupRecords) return null;
            Ensure(feature + 4, checked(count * 2));
            for (int lookup = 0; lookup < count; lookup++) {
                if ((lookup & 63) == 0) cancellationToken.ThrowIfCancellationRequested();
                int lookupIndex = _reader.ReadUInt16(feature + 4 + lookup * 2);
                if (!CanApplyLookup(lookupIndex, 0, ref inspections, cancellationToken)) return null;
                lookups[lookupIndex] = setting;
            }
        }
        var result = new KeyValuePair<int, int>[lookups.Count];
        int destination = 0;
        foreach (var lookup in lookups) result[destination++] = lookup;
        return result;
    }
}
