using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>Shared bounded legacy kern and GPOS pair-adjustment reader.</summary>
internal sealed class OfficeOpenTypeKerning {
    private readonly byte[] _data;
    private readonly int _kern;
    private readonly int _gpos;
    private readonly bool _includeExtendedGpos;

    internal OfficeOpenTypeKerning(byte[] data, int kern, int gpos, bool includeExtendedGpos = false) {
        _data = data ?? throw new ArgumentNullException(nameof(data));
        _kern = kern;
        _gpos = gpos;
        _includeExtendedGpos = includeExtendedGpos;
    }

    internal static OfficeOpenTypeKerning FromReader(OfficeOpenTypeReader reader) {
        if (reader == null) throw new ArgumentNullException(nameof(reader));
        return new OfficeOpenTypeKerning(
            reader.Data,
            reader.TryGetTable("kern", out int kern, out _) ? kern : -1,
            reader.TryGetTable("GPOS", out int gpos, out _) ? gpos : -1,
            includeExtendedGpos: true);
    }

    internal int Adjustment(int left, int right) {
        if ((uint)left > ushort.MaxValue || (uint)right > ushort.MaxValue) return 0;
        return checked(KernPairAdjustment((ushort)left, (ushort)right) +
                       GposPairAdjustment((ushort)left, (ushort)right));
    }

    private int KernPairAdjustment(ushort left, ushort right) {
        if (_kern < 0 || !InBounds(_kern, 4) || ReadUInt16(_kern) != 0) return 0;
        int count = ReadUInt16(_kern + 2);
        int position = _kern + 4;
        int adjustment = 0;
        for (int table = 0; table < count; table++) {
            if (!InBounds(position, 6)) break;
            int length = ReadUInt16(position + 2);
            int coverage = ReadUInt16(position + 4);
            int next = checked(position + length);
            if (length < 14 || next <= position || next > _data.Length) break;
            if ((coverage >> 8) == 0) adjustment = checked(adjustment + KerningFormat0(position, left, right));
            position = next;
        }
        return adjustment;
    }

    private int KerningFormat0(int table, ushort left, ushort right) {
        int pairs = ReadUInt16(table + 6);
        int pairOffset = table + 14;
        uint key = ((uint)left << 16) | right;
        int low = 0;
        int high = pairs - 1;
        while (low <= high) {
            int mid = low + ((high - low) / 2);
            int record = checked(pairOffset + (mid * 6));
            if (!InBounds(record, 6)) return 0;
            uint candidate = ReadUInt32(record);
            if (candidate == key) return ReadInt16(record + 4);
            if (candidate < key) low = mid + 1;
            else high = mid - 1;
        }
        return 0;
    }

    private int GposPairAdjustment(ushort left, ushort right) {
        if (_gpos < 0 || !InBounds(_gpos, 10) || ReadUInt16(_gpos) != 1) return 0;
        int featureList = checked(_gpos + ReadUInt16(_gpos + 6));
        int lookupList = checked(_gpos + ReadUInt16(_gpos + 8));
        if (!InBounds(featureList, 2) || !InBounds(lookupList, 2)) return 0;

        int adjustment = 0;
        var seen = new HashSet<ushort>();
        foreach (ushort lookupIndex in GposFeatureLookupIndexes(featureList, "kern")) {
            if (seen.Add(lookupIndex)) {
                adjustment = checked(adjustment + GposPairAdjustmentFromLookup(lookupList, lookupIndex, left, right));
            }
        }
        return adjustment;
    }

    private IEnumerable<ushort> GposFeatureLookupIndexes(int featureList, string featureTag) {
        int featureCount = ReadUInt16(featureList);
        for (int index = 0; index < featureCount; index++) {
            int record = checked(featureList + 2 + (index * 6));
            if (!InBounds(record, 6)) yield break;
            if (!TagEquals(record, featureTag)) continue;
            int feature = checked(featureList + ReadUInt16(record + 4));
            if (!InBounds(feature, 4)) yield break;
            int lookupCount = ReadUInt16(feature + 2);
            for (int lookup = 0; lookup < lookupCount; lookup++) {
                int offset = checked(feature + 4 + (lookup * 2));
                if (!InBounds(offset, 2)) yield break;
                yield return ReadUInt16(offset);
            }
        }
    }

    private int GposPairAdjustmentFromLookup(int lookupList, ushort lookupIndex, ushort left, ushort right) {
        int lookupCount = ReadUInt16(lookupList);
        if (lookupIndex >= lookupCount) return 0;
        int lookupOffset = checked(lookupList + 2 + (lookupIndex * 2));
        if (!InBounds(lookupOffset, 2)) return 0;
        int lookup = checked(lookupList + ReadUInt16(lookupOffset));
        if (!InBounds(lookup, 6)) return 0;
        ushort lookupType = ReadUInt16(lookup);
        if (lookupType != 2 && lookupType != 9) return 0;
        if (lookupType == 9 && !_includeExtendedGpos) return 0;

        int adjustment = 0;
        int subtableCount = ReadUInt16(lookup + 4);
        for (int index = 0; index < subtableCount; index++) {
            int subtableOffset = checked(lookup + 6 + (index * 2));
            if (!InBounds(subtableOffset, 2)) break;
            int subtable = checked(lookup + ReadUInt16(subtableOffset));
            if (lookupType == 9) {
                if (!InBounds(subtable, 8) || ReadUInt16(subtable) != 1 || ReadUInt16(subtable + 2) != 2) continue;
                uint extensionOffset = ReadUInt32(subtable + 4);
                if (extensionOffset > int.MaxValue) continue;
                subtable = checked(subtable + (int)extensionOffset);
            }
            adjustment = checked(adjustment + GposPairAdjustmentFromSubtable(subtable, left, right));
        }
        return adjustment;
    }

    private int GposPairAdjustmentFromSubtable(int subtable, ushort left, ushort right) {
        if (!InBounds(subtable, 10)) return 0;
        ushort format = ReadUInt16(subtable);
        if (format == 2) return _includeExtendedGpos ? GposClassPairAdjustment(subtable, left, right) : 0;
        if (format != 1) return 0;
        int coverage = checked(subtable + ReadUInt16(subtable + 2));
        ushort valueFormat1 = ReadUInt16(subtable + 4);
        ushort valueFormat2 = ReadUInt16(subtable + 6);
        int pairSetCount = ReadUInt16(subtable + 8);
        int coverageIndex = CoverageIndex(coverage, left);
        if (coverageIndex < 0 || coverageIndex >= pairSetCount) return 0;

        int pairSetOffset = checked(subtable + 10 + (coverageIndex * 2));
        if (!InBounds(pairSetOffset, 2)) return 0;
        int pairSet = checked(subtable + ReadUInt16(pairSetOffset));
        if (!InBounds(pairSet, 2)) return 0;

        int value1Size = ValueRecordSize(valueFormat1);
        int value2Size = ValueRecordSize(valueFormat2);
        int recordSize = checked(2 + value1Size + value2Size);
        int low = 0;
        int high = ReadUInt16(pairSet) - 1;
        while (low <= high) {
            int mid = low + ((high - low) / 2);
            int record = checked(pairSet + 2 + (mid * recordSize));
            if (!InBounds(record, recordSize)) return 0;
            ushort candidate = ReadUInt16(record);
            if (candidate == right) return ReadValueRecordXAdvance(record + 2, valueFormat1);
            if (candidate < right) low = mid + 1;
            else high = mid - 1;
        }
        return 0;
    }

    private int GposClassPairAdjustment(int subtable, ushort left, ushort right) {
        if (!InBounds(subtable, 16)) return 0;
        int coverage = checked(subtable + ReadUInt16(subtable + 2));
        if (CoverageIndex(coverage, left) < 0) return 0;

        ushort valueFormat1 = ReadUInt16(subtable + 4);
        ushort valueFormat2 = ReadUInt16(subtable + 6);
        int classDef1 = checked(subtable + ReadUInt16(subtable + 8));
        int classDef2 = checked(subtable + ReadUInt16(subtable + 10));
        int class1Count = ReadUInt16(subtable + 12);
        int class2Count = ReadUInt16(subtable + 14);
        int class1 = ClassDefinition(classDef1, left);
        int class2 = ClassDefinition(classDef2, right);
        if (class1 < 0 || class2 < 0 || class1 >= class1Count || class2 >= class2Count) return 0;

        int value1Size = ValueRecordSize(valueFormat1);
        int value2Size = ValueRecordSize(valueFormat2);
        int recordSize = checked(value1Size + value2Size);
        if (recordSize == 0) return 0;
        int recordIndex = checked((class1 * class2Count) + class2);
        int record = checked(subtable + 16 + (recordIndex * recordSize));
        return InBounds(record, recordSize) ? ReadValueRecordXAdvance(record, valueFormat1) : 0;
    }

    private int ClassDefinition(int classDef, ushort glyph) {
        if (!InBounds(classDef, 4)) return -1;
        ushort format = ReadUInt16(classDef);
        if (format == 1) {
            ushort startGlyph = ReadUInt16(classDef + 2);
            if (!InBounds(classDef, 6)) return -1;
            int glyphCount = ReadUInt16(classDef + 4);
            int index = glyph - startGlyph;
            if (index < 0 || index >= glyphCount) return 0;
            int valueOffset = checked(classDef + 6 + (index * 2));
            return InBounds(valueOffset, 2) ? ReadUInt16(valueOffset) : -1;
        }

        if (format != 2) return -1;
        int rangeCount = ReadUInt16(classDef + 2);
        int low = 0;
        int high = rangeCount - 1;
        while (low <= high) {
            int mid = low + ((high - low) / 2);
            int range = checked(classDef + 4 + (mid * 6));
            if (!InBounds(range, 6)) return -1;
            ushort start = ReadUInt16(range);
            ushort end = ReadUInt16(range + 2);
            if (glyph < start) high = mid - 1;
            else if (glyph > end) low = mid + 1;
            else return ReadUInt16(range + 4);
        }
        return 0;
    }

    private int CoverageIndex(int coverage, ushort glyph) {
        if (!InBounds(coverage, 4)) return -1;
        ushort format = ReadUInt16(coverage);
        if (format == 1) {
            int count = ReadUInt16(coverage + 2);
            int low = 0;
            int high = count - 1;
            while (low <= high) {
                int mid = low + ((high - low) / 2);
                int offset = checked(coverage + 4 + (mid * 2));
                if (!InBounds(offset, 2)) return -1;
                ushort candidate = ReadUInt16(offset);
                if (candidate == glyph) return mid;
                if (candidate < glyph) low = mid + 1;
                else high = mid - 1;
            }
            return -1;
        }

        if (format != 2) return -1;
        int rangeCount = ReadUInt16(coverage + 2);
        for (int index = 0; index < rangeCount; index++) {
            int range = checked(coverage + 4 + (index * 6));
            if (!InBounds(range, 6)) return -1;
            ushort start = ReadUInt16(range);
            ushort end = ReadUInt16(range + 2);
            if (glyph < start || glyph > end) continue;
            return checked(ReadUInt16(range + 4) + glyph - start);
        }
        return -1;
    }

    private int ReadValueRecordXAdvance(int offset, ushort valueFormat) {
        if ((valueFormat & 0x0001) != 0) offset += 2;
        if ((valueFormat & 0x0002) != 0) offset += 2;
        if ((valueFormat & 0x0004) == 0) return 0;
        return InBounds(offset, 2) ? ReadInt16(offset) : 0;
    }

    private static int ValueRecordSize(ushort valueFormat) {
        int size = 0;
        for (int bit = 1; bit <= 0x0080; bit <<= 1) {
            if ((valueFormat & bit) != 0) size += 2;
        }
        return size;
    }

    private bool TagEquals(int offset, string tag) =>
        InBounds(offset, 4) &&
        _data[offset] == tag[0] &&
        _data[offset + 1] == tag[1] &&
        _data[offset + 2] == tag[2] &&
        _data[offset + 3] == tag[3];

    private bool InBounds(int offset, int count) => offset >= 0 && count >= 0 && offset <= _data.Length - count;
    private ushort ReadUInt16(int offset) => (ushort)((_data[offset] << 8) | _data[offset + 1]);
    private short ReadInt16(int offset) => unchecked((short)ReadUInt16(offset));
    private uint ReadUInt32(int offset) =>
        ((uint)_data[offset] << 24) |
        ((uint)_data[offset + 1] << 16) |
        ((uint)_data[offset + 2] << 8) |
        _data[offset + 3];
}
