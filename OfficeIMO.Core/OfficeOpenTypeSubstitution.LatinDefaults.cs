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

    private int[]? ReadLatinDefaultFeatureIndexes(out int requiredFeature) =>
        OfficeOpenTypeLayoutFeatures.GetLatinDefaultIndexes(_reader, _table, _end - _table, out requiredFeature);

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
        LatinDefaultLookupCache selected = ResolveLatinDefaultLookups(settings, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        if (selected.Lookups == null) return false;
        if (selected.Lookups.Length == 0) return true;
        int operations = 0;
        // Script-specific lookups must not consume neighboring non-Latin or presentation glyphs.
        var shaped = new List<GlyphToken>(glyphs.Count);
        for (int index = 0; index < glyphs.Count;) {
            if (!eligible[index]) { shaped.Add(glyphs[index++]); continue; }
            var segment = new List<GlyphToken>();
            do { segment.Add(glyphs[index++]); }
            while (index < glyphs.Count && eligible[index] && !breakBefore[index]);
            if (!ApplyLatinDefaultLookups(segment, selected, cancellationToken, ref operations)) return false;
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
        LatinDefaultLookupCache selected = ResolveLatinDefaultLookups(settings, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        if (selected.Lookups == null) return false;
        if (selected.Lookups.Length == 0) return true;
        // A single uninterrupted ASCII segment has the same eligibility as the general
        // script resolver. Keep its working copy so a rejected operation leaves input intact.
        var segment = new List<GlyphToken>(glyphs);
        int operations = 0;
        if (!ApplyLatinDefaultLookups(segment, selected, cancellationToken, ref operations)) return false;
        glyphs.Clear(); glyphs.AddRange(segment);
        return true;
    }

    // Lookups that skip marks are applied unfiltered, which is exact only while the
    // segment holds no mark. A segment that has or gains one declines the run.
    private bool ApplyLatinDefaultLookups(List<GlyphToken> segment, LatinDefaultLookupCache selected,
        CancellationToken cancellationToken, ref int operations) {
        foreach (var lookup in selected.Lookups!) {
            if (selected.FiltersMarks && ContainsMarkGlyph(segment, cancellationToken)) return false;
            ApplyLookup(segment, lookup.Key, lookup.Value, cancellationToken, ref operations);
        }
        return !selected.FiltersMarks || !ContainsMarkGlyph(segment, cancellationToken);
    }

    private bool ContainsMarkGlyph(List<GlyphToken> glyphs, CancellationToken cancellationToken) {
        OfficeOpenTypeGlyphFilter? filter = _glyphFilter.Value;
        if (filter == null) return true;
        for (int index = 0; index < glyphs.Count; index++) {
            if ((index & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
            GlyphToken glyph = glyphs[index];
            if (filter.GetGlyphClass(glyph.GlyphId, glyph.Scalar, glyph.LigatureComponentTextIndexes?.Count > 1, NoCharge) == 3) return true;
        }
        return false;
    }

    private static readonly Action NoCharge = static () => { };

    internal bool CanApplyLatinDefaults(OfficeTextFeatureSettings settings, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        try {
            bool supported = ResolveLatinDefaultLookups(settings, cancellationToken).Lookups != null;
            cancellationToken.ThrowIfCancellationRequested();
            return supported;
        } catch (Exception exception) when (exception is InvalidDataException || exception is OverflowException ||
            exception is ArgumentOutOfRangeException || exception is IndexOutOfRangeException) { return false; }
    }

    // Default settings and the most recent explicit settings keep their selected lookups.
    // Document runs usually repeat one explicit feature set, such as Word's kerning and ligatures.
    private LatinDefaultLookupCache ResolveLatinDefaultLookups(OfficeTextFeatureSettings settings, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        LatinDefaultLookupCache? cached = Volatile.Read(ref settings.IsDefault ? ref _latinDefaultLookups : ref _latinExplicitLookups);
        if (cached != null && cached.Settings.Equals(settings)) return cached;
        LatinDefaultLookupCache completed;
        try { completed = BuildLatinDefaultLookups(settings, cancellationToken); }
        catch (Exception exception) when (exception is InvalidDataException || exception is OverflowException ||
            exception is ArgumentOutOfRangeException || exception is IndexOutOfRangeException) { completed = new LatinDefaultLookupCache(settings, null, false); }
        cancellationToken.ThrowIfCancellationRequested();
        // Publish completed immutable results only. A cancelled caller does not
        // poison the cache or make another caller wait for its work.
        if (settings.IsDefault) Volatile.Write(ref _latinDefaultLookups, completed);
        else Volatile.Write(ref _latinExplicitLookups, completed);
        return completed;
    }

    private LatinDefaultLookupCache BuildLatinDefaultLookups(OfficeTextFeatureSettings settings, CancellationToken cancellationToken) {
        int[]? features = GetLatinDefaultFeatureIndexes(out int requiredFeature);
        if (features == null) return new LatinDefaultLookupCache(settings, null, false);
        var lookups = new SortedDictionary<int, int>();
        var markFilters = new MarkFilterUse();
        int inspections = 0;
        foreach (int index in features) {
            cancellationToken.ThrowIfCancellationRequested();
            int record = _featureList + 2 + index * 6;
            int setting = LatinDefaultFeatureSetting(index, requiredFeature, ReadTag(record), settings);
            if (setting <= 0) continue;
            int feature = Relative(_featureList, _reader.ReadUInt16(record + 4), 4);
            int count = _reader.ReadUInt16(feature + 2);
            if (count > MaximumLookupRecords) return new LatinDefaultLookupCache(settings, null, false);
            Ensure(feature + 4, checked(count * 2));
            for (int lookup = 0; lookup < count; lookup++) {
                if ((lookup & 63) == 0) cancellationToken.ThrowIfCancellationRequested();
                int lookupIndex = _reader.ReadUInt16(feature + 4 + lookup * 2);
                if (!CanApplyLookup(lookupIndex, 0, ref inspections, cancellationToken, markFilters: markFilters))
                    return new LatinDefaultLookupCache(settings, null, false);
                lookups[lookupIndex] = setting;
            }
        }
        var result = new KeyValuePair<int, int>[lookups.Count];
        int destination = 0;
        foreach (var lookup in lookups) result[destination++] = lookup;
        return new LatinDefaultLookupCache(settings, result, markFilters.Used);
    }

    private static int LatinDefaultFeatureSetting(int featureIndex, int requiredFeature, string tag, OfficeTextFeatureSettings settings) =>
        featureIndex == requiredFeature ? 1 : settings.TryGetValue(tag, out int explicitValue) ? explicitValue
            : tag == "liga" || tag == "clig" || tag == "rlig" ? 1 : 0;

    /// <summary>
    /// Reports whether a selected Latin lookup that managed shaping declines could start a match on one of these glyphs.
    /// </summary>
    /// <remarks>
    /// A declined lookup only simplifies output where its first-glyph coverage meets the run, or where a supported
    /// lookup changes the run first. Tables whose features, lookups or coverage cannot be read are treated as applicable.
    /// </remarks>
    internal bool DeclinedLatinDefaultLookupMayApply(ICollection<int> glyphIds, OfficeTextFeatureSettings settings, CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(glyphIds);
        ArgumentNullException.ThrowIfNull(settings);
        try {
            int[]? features = GetLatinDefaultFeatureIndexes(out int requiredFeature);
            if (features == null) return true;
            int inspections = 0;
            int operations = 0;
            bool declined = false, applied = false;
            foreach (int index in features) {
                cancellationToken.ThrowIfCancellationRequested();
                int record = _featureList + 2 + index * 6;
                if (LatinDefaultFeatureSetting(index, requiredFeature, ReadTag(record), settings) <= 0) continue;
                int feature = Relative(_featureList, _reader.ReadUInt16(record + 4), 4);
                int count = _reader.ReadUInt16(feature + 2);
                if (count > MaximumLookupRecords) return true;
                Ensure(feature + 4, checked(count * 2));
                for (int lookup = 0; lookup < count; lookup++) {
                    int lookupIndex = _reader.ReadUInt16(feature + 4 + lookup * 2);
                    if (CanApplyLookup(lookupIndex, 0, ref inspections, cancellationToken)) {
                        if (!applied) applied = LookupMayStartAt(lookupIndex, glyphIds, ref operations, cancellationToken);
                        continue;
                    }
                    if (LookupMayStartAt(lookupIndex, glyphIds, ref operations, cancellationToken)) return true;
                    declined = true;
                }
            }
            // A supported lookup that changes this run can hand a declined lookup glyphs the run did not start with.
            return declined && applied;
        } catch (Exception exception) when (exception is InvalidDataException || exception is OverflowException ||
            exception is ArgumentOutOfRangeException || exception is IndexOutOfRangeException) { return true; }
    }

    private bool LookupMayStartAt(int lookupIndex, ICollection<int> glyphIds, ref int operations, CancellationToken cancellationToken) {
        Ensure(_lookupList, 2);
        int lookupCount = _reader.ReadUInt16(_lookupList);
        if (lookupCount > MaximumLookupRecords || lookupIndex < 0 || lookupIndex >= lookupCount) return true;
        Ensure(_lookupList + 2, checked(lookupCount * 2));
        int lookup = Relative(_lookupList, _reader.ReadUInt16(_lookupList + 2 + lookupIndex * 2), 6);
        int lookupType = _reader.ReadUInt16(lookup);
        int subtableCount = _reader.ReadUInt16(lookup + 4);
        if (subtableCount > MaximumSubtablesPerLookup) return true;
        Ensure(lookup + 6, checked(subtableCount * 2));
        for (int index = 0; index < subtableCount; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            int subtable = Relative(lookup, _reader.ReadUInt16(lookup + 6 + index * 2), 4);
            int type = lookupType;
            if (lookupType == 7) {
                Ensure(subtable, 8);
                if (_reader.ReadUInt16(subtable) != 1) return true;
                type = _reader.ReadUInt16(subtable + 2);
                uint relative = _reader.ReadUInt32(subtable + 4);
                if (relative > int.MaxValue) return true;
                subtable = Relative(subtable, (int)relative, 4);
            }
            int coverage = FirstGlyphCoverage(type, subtable);
            if (coverage < 0) return true;
            foreach (int glyphId in glyphIds) {
                if (CoverageIndex(coverage, glyphId, ref operations) >= 0) return true;
            }
        }
        return false;
    }

    // Offset of the coverage that a subtable's first input glyph must match, or -1 when unknown.
    private int FirstGlyphCoverage(int lookupType, int subtable) {
        int format = _reader.ReadUInt16(subtable);
        switch (lookupType) {
            case 1:
            case 2:
            case 3:
            case 4:
            case 8:
                return Relative(subtable, _reader.ReadUInt16(subtable + 2), 4);
            case 5:
            case 6:
                if (format == 1 || format == 2) return Relative(subtable, _reader.ReadUInt16(subtable + 2), 4);
                if (format != 3) return -1;
                int inputCount;
                int firstInput;
                if (lookupType == 5) {
                    Ensure(subtable, 8);
                    inputCount = _reader.ReadUInt16(subtable + 2);
                    firstInput = subtable + 6;
                } else {
                    int backtrack = ReadBoundedCount(subtable + 2);
                    int cursor = checked(subtable + 4 + backtrack * 2);
                    Ensure(cursor, 4);
                    inputCount = _reader.ReadUInt16(cursor);
                    firstInput = cursor + 2;
                }
                return inputCount == 0 ? -1 : Relative(subtable, _reader.ReadUInt16(firstInput), 4);
            default:
                return -1;
        }
    }
}
