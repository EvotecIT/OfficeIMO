namespace OfficeIMO.Drawing;

internal sealed partial class OfficeOpenTypeSubstitution {
    /// <summary>
    /// Reports whether a Latin feature's single or alternate substitutions start at this glyph,
    /// through lookups the managed shaper applies.
    /// </summary>
    /// <remarks>Unreadable or unsupported tables report false, so callers keep their own fallback.</remarks>
    internal bool LatinFeatureSubstitutes(string featureTag, int glyphId, CancellationToken cancellationToken) {
        try {
            int[]? features = GetLatinDefaultFeatureIndexes(out _);
            if (features == null) return false;
            int inspections = 0;
            int operations = 0;
            foreach (int index in features) {
                cancellationToken.ThrowIfCancellationRequested();
                int record = _featureList + 2 + index * 6;
                if (ReadTag(record) != featureTag) continue;
                int feature = Relative(_featureList, _reader.ReadUInt16(record + 4), 4);
                int count = _reader.ReadUInt16(feature + 2);
                if (count > MaximumLookupRecords) return false;
                Ensure(feature + 4, checked(count * 2));
                for (int lookup = 0; lookup < count; lookup++) {
                    int lookupIndex = _reader.ReadUInt16(feature + 4 + lookup * 2);
                    if (!CanApplyLookup(lookupIndex, 0, ref inspections, cancellationToken)) return false;
                    if (SingleOrAlternateLookupCovers(lookupIndex, glyphId, ref operations, cancellationToken)) return true;
                }
            }
            return false;
        } catch (Exception exception) when (exception is InvalidDataException || exception is OverflowException ||
            exception is ArgumentOutOfRangeException || exception is IndexOutOfRangeException) { return false; }
    }

    private bool SingleOrAlternateLookupCovers(int lookupIndex, int glyphId, ref int operations, CancellationToken cancellationToken) {
        Ensure(_lookupList, 2);
        int lookupCount = _reader.ReadUInt16(_lookupList);
        if (lookupCount > MaximumLookupRecords || lookupIndex < 0 || lookupIndex >= lookupCount) return false;
        Ensure(_lookupList + 2, checked(lookupCount * 2));
        int lookup = Relative(_lookupList, _reader.ReadUInt16(_lookupList + 2 + lookupIndex * 2), 6);
        int lookupType = _reader.ReadUInt16(lookup);
        int subtableCount = _reader.ReadUInt16(lookup + 4);
        if (subtableCount > MaximumSubtablesPerLookup) return false;
        Ensure(lookup + 6, checked(subtableCount * 2));
        for (int index = 0; index < subtableCount; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            int subtable = Relative(lookup, _reader.ReadUInt16(lookup + 6 + index * 2), 4);
            int type = lookupType;
            if (lookupType == 7) {
                Ensure(subtable, 8);
                if (_reader.ReadUInt16(subtable) != 1) return false;
                type = _reader.ReadUInt16(subtable + 2);
                uint relative = _reader.ReadUInt32(subtable + 4);
                if (relative > int.MaxValue) return false;
                subtable = Relative(subtable, (int)relative, 4);
            }
            if (type != 1 && type != 3) continue;
            int coverage = Relative(subtable, _reader.ReadUInt16(subtable + 2), 4);
            if (CoverageIndex(coverage, glyphId, ref operations) >= 0) return true;
        }
        return false;
    }
}
