using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

internal readonly struct OfficeOpenTypePairPositioning {
    internal OfficeOpenTypePairPositioning(
        int firstGlyphXPlacement,
        int firstGlyphXAdvance,
        int secondGlyphXPlacement,
        int secondGlyphXAdvance,
        bool secondValueRecordPresent = false) {
        FirstGlyphXPlacement = firstGlyphXPlacement;
        FirstGlyphXAdvance = firstGlyphXAdvance;
        SecondGlyphXPlacement = secondGlyphXPlacement;
        SecondGlyphXAdvance = secondGlyphXAdvance;
        SecondValueRecordPresent = secondValueRecordPresent;
    }

    internal int FirstGlyphXPlacement { get; }
    internal int FirstGlyphXAdvance { get; }
    internal int SecondGlyphXPlacement { get; }
    internal int SecondGlyphXAdvance { get; }
    internal bool SecondValueRecordPresent { get; }
    internal int TotalXAdvance => checked(FirstGlyphXAdvance + SecondGlyphXAdvance);

    internal OfficeOpenTypePairPositioning Add(OfficeOpenTypePairPositioning other) => new(
        checked(FirstGlyphXPlacement + other.FirstGlyphXPlacement),
        checked(FirstGlyphXAdvance + other.FirstGlyphXAdvance),
        checked(SecondGlyphXPlacement + other.SecondGlyphXPlacement),
        checked(SecondGlyphXAdvance + other.SecondGlyphXAdvance),
        SecondValueRecordPresent || other.SecondValueRecordPresent);
}

internal readonly struct OfficeOpenTypeGlyphPositioning {
    internal OfficeOpenTypeGlyphPositioning(int xPlacement, int xAdvance) {
        XPlacement = xPlacement;
        XAdvance = xAdvance;
    }

    internal int XPlacement { get; }
    internal int XAdvance { get; }

    internal OfficeOpenTypeGlyphPositioning Add(int xPlacement, int xAdvance) => new(
        checked(XPlacement + xPlacement),
        checked(XAdvance + xAdvance));
}

/// <summary>Shared bounded legacy kern and GPOS pair-adjustment reader.</summary>
internal sealed class OfficeOpenTypeKerning {
    private readonly byte[] _data;
    private readonly int _kern;
    private readonly int _gpos;
    private readonly bool _includeExtendedGpos;

    internal OfficeOpenTypeKerning(byte[] data, int kern, int gpos, bool includeExtendedGpos = true) {
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

    internal int Adjustment(int left, int right) => Positioning(left, right, "DFLT").TotalXAdvance;

    internal int Adjustment(int left, int right, int leftScalar, int rightScalar) =>
        Positioning(left, right, ResolveScriptTag(leftScalar, rightScalar)).TotalXAdvance;

    internal int Adjustment(int left, int right, string scriptTag) =>
        Positioning(left, right, scriptTag).TotalXAdvance;

    internal OfficeOpenTypePairPositioning Positioning(int left, int right, int leftScalar, int rightScalar) =>
        Positioning(left, right, ResolveScriptTag(leftScalar, rightScalar));

    internal OfficeOpenTypePairPositioning Positioning(int left, int right, string scriptTag) {
        if ((uint)left > ushort.MaxValue || (uint)right > ushort.MaxValue) return default;
        if (string.IsNullOrEmpty(scriptTag) || scriptTag.Length != 4) scriptTag = "DFLT";
        return TryGposPairAdjustment((ushort)left, (ushort)right, scriptTag, out OfficeOpenTypePairPositioning gposAdjustment)
            ? gposAdjustment
            : new OfficeOpenTypePairPositioning(0, KernPairAdjustment((ushort)left, (ushort)right), 0, 0);
    }

    internal OfficeOpenTypeGlyphPositioning[] PositionRun(
        IReadOnlyList<int> glyphs,
        IReadOnlyList<int> scalars) {
        if (glyphs == null) throw new ArgumentNullException(nameof(glyphs));
        if (scalars == null) throw new ArgumentNullException(nameof(scalars));
        if (glyphs.Count != scalars.Count) throw new ArgumentException("Glyph and scalar runs must have the same length.");

        var result = new OfficeOpenTypeGlyphPositioning[glyphs.Count];
        if (glyphs.Count < 2) return result;
        for (int index = 0; index < glyphs.Count; index++) {
            if ((uint)glyphs[index] > ushort.MaxValue) return result;
        }

        int pairCount = glyphs.Count - 1;
        var pairLookups = new List<ushort>[pairCount];
        var orderedLookups = new List<ushort>();
        var knownLookups = new HashSet<ushort>();
        bool hasGposLayout = TryGetGposLayoutTables(out int scriptList, out int featureList, out int lookupList);
        if (hasGposLayout) {
            for (int index = 0; index < pairCount; index++) {
                string scriptTag = ResolveScriptTag(scalars[index], scalars[index + 1]);
                var lookups = new List<ushort>();
                var pairSeen = new HashSet<ushort>();
                foreach (ushort lookupIndex in GposFeatureLookupIndexes(scriptList, featureList, "kern", scriptTag)) {
                    if (!pairSeen.Add(lookupIndex)) continue;
                    lookups.Add(lookupIndex);
                    if (knownLookups.Add(lookupIndex)) orderedLookups.Add(lookupIndex);
                }
                pairLookups[index] = lookups;
            }
        }

        var gposApplied = new bool[pairCount];
        foreach (ushort lookupIndex in orderedLookups) {
            for (int index = 0; index < pairCount;) {
                List<ushort> activeLookups = pairLookups[index];
                if (activeLookups.Contains(lookupIndex) &&
                    TryGposPairAdjustmentFromLookup(
                        lookupList,
                        lookupIndex,
                        (ushort)glyphs[index],
                        (ushort)glyphs[index + 1],
                        out OfficeOpenTypePairPositioning adjustment)) {
                    result[index] = result[index].Add(
                        adjustment.FirstGlyphXPlacement,
                        adjustment.FirstGlyphXAdvance);
                    result[index + 1] = result[index + 1].Add(
                        adjustment.SecondGlyphXPlacement,
                        adjustment.SecondGlyphXAdvance);
                    gposApplied[index] = true;
                    index += adjustment.SecondValueRecordPresent ? 2 : 1;
                } else {
                    index++;
                }
            }
        }

        for (int index = 0; index < pairCount; index++) {
            if (gposApplied[index]) continue;
            int legacyAdjustment = KernPairAdjustment((ushort)glyphs[index], (ushort)glyphs[index + 1]);
            result[index] = result[index].Add(0, legacyAdjustment);
        }
        return result;
    }

    private bool TryGetGposLayoutTables(out int scriptList, out int featureList, out int lookupList) {
        scriptList = 0;
        featureList = 0;
        lookupList = 0;
        if (_gpos < 0 || !InBounds(_gpos, 10) || ReadUInt16(_gpos) != 1) return false;
        scriptList = checked(_gpos + ReadUInt16(_gpos + 4));
        featureList = checked(_gpos + ReadUInt16(_gpos + 6));
        lookupList = checked(_gpos + ReadUInt16(_gpos + 8));
        return InBounds(scriptList, 2) && InBounds(featureList, 2) && InBounds(lookupList, 2);
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
            bool isFormat0 = (coverage >> 8) == 0;
            bool isHorizontalOrdinary = (coverage & 0x0007) == 0x0001;
            if (isFormat0 && isHorizontalOrdinary) {
                int value = KerningFormat0(position, left, right);
                adjustment = (coverage & 0x0008) != 0
                    ? value
                    : checked(adjustment + value);
            }
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

    private bool TryGposPairAdjustment(ushort left, ushort right, string scriptTag, out OfficeOpenTypePairPositioning adjustment) {
        adjustment = default;
        if (!TryGetGposLayoutTables(out int scriptList, out int featureList, out int lookupList)) return false;

        bool applied = false;
        var seen = new HashSet<ushort>();
        foreach (ushort lookupIndex in GposFeatureLookupIndexes(scriptList, featureList, "kern", scriptTag)) {
            if (seen.Add(lookupIndex) &&
                TryGposPairAdjustmentFromLookup(lookupList, lookupIndex, left, right, out OfficeOpenTypePairPositioning lookupAdjustment)) {
                adjustment = adjustment.Add(lookupAdjustment);
                applied = true;
            }
        }
        return applied;
    }

    private IEnumerable<ushort> GposFeatureLookupIndexes(
        int scriptList,
        int featureList,
        string featureTag,
        string scriptTag) {
        int script = FindScript(scriptList, scriptTag);
        if (script < 0 && !string.Equals(scriptTag, "DFLT", StringComparison.Ordinal)) {
            script = FindScript(scriptList, "DFLT");
        }
        if (script < 0 || !InBounds(script, 4)) yield break;

        int defaultLangSysOffset = ReadUInt16(script);
        if (defaultLangSysOffset == 0) yield break;
        int langSys = checked(script + defaultLangSysOffset);
        if (!InBounds(langSys, 6)) yield break;

        int featureCount = ReadUInt16(featureList);
        var featureIndexes = new List<ushort>();
        ushort requiredFeature = ReadUInt16(langSys + 2);
        if (requiredFeature != ushort.MaxValue) featureIndexes.Add(requiredFeature);
        int langSysFeatureCount = ReadUInt16(langSys + 4);
        for (int index = 0; index < langSysFeatureCount; index++) {
            int offset = checked(langSys + 6 + (index * 2));
            if (!InBounds(offset, 2)) yield break;
            ushort featureIndex = ReadUInt16(offset);
            if (!featureIndexes.Contains(featureIndex)) featureIndexes.Add(featureIndex);
        }

        foreach (ushort featureIndex in featureIndexes) {
            if (featureIndex >= featureCount) continue;
            int record = checked(featureList + 2 + (featureIndex * 6));
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

    private int FindScript(int scriptList, string scriptTag) {
        int scriptCount = ReadUInt16(scriptList);
        for (int index = 0; index < scriptCount; index++) {
            int record = checked(scriptList + 2 + (index * 6));
            if (!InBounds(record, 6)) return -1;
            if (!TagEquals(record, scriptTag)) continue;
            int script = checked(scriptList + ReadUInt16(record + 4));
            return InBounds(script, 4) ? script : -1;
        }
        return -1;
    }

    private static string ResolveScriptTag(int leftScalar, int rightScalar) {
        string right = ScriptTagForScalar(rightScalar);
        if (right != "DFLT") return right;
        return ScriptTagForScalar(leftScalar);
    }

    private static string ScriptTagForScalar(int scalar) {
        if ((scalar >= 0x0041 && scalar <= 0x024F) || (scalar >= 0x1E00 && scalar <= 0x1EFF)) return "latn";
        if (scalar >= 0x0370 && scalar <= 0x03FF) return "grek";
        if (scalar >= 0x0400 && scalar <= 0x052F) return "cyrl";
        if (scalar >= 0x0530 && scalar <= 0x058F) return "armn";
        if (scalar >= 0x0590 && scalar <= 0x05FF) return "hebr";
        if ((scalar >= 0x0600 && scalar <= 0x06FF) || (scalar >= 0x0750 && scalar <= 0x077F) ||
            (scalar >= 0x08A0 && scalar <= 0x08FF)) return "arab";
        if (scalar >= 0x0700 && scalar <= 0x074F) return "syrc";
        if (scalar >= 0x0900 && scalar <= 0x097F) return "deva";
        if (scalar >= 0x0980 && scalar <= 0x09FF) return "beng";
        if (scalar >= 0x0A00 && scalar <= 0x0A7F) return "guru";
        if (scalar >= 0x0A80 && scalar <= 0x0AFF) return "gujr";
        if (scalar >= 0x0B00 && scalar <= 0x0B7F) return "orya";
        if (scalar >= 0x0B80 && scalar <= 0x0BFF) return "taml";
        if (scalar >= 0x0C00 && scalar <= 0x0C7F) return "telu";
        if (scalar >= 0x0C80 && scalar <= 0x0CFF) return "knda";
        if (scalar >= 0x0D00 && scalar <= 0x0D7F) return "mlym";
        if (scalar >= 0x0D80 && scalar <= 0x0DFF) return "sinh";
        if (scalar >= 0x0E00 && scalar <= 0x0E7F) return "thai";
        if (scalar >= 0x0E80 && scalar <= 0x0EFF) return "lao ";
        if (scalar >= 0x0F00 && scalar <= 0x0FFF) return "tibt";
        if (scalar >= 0x1000 && scalar <= 0x109F) return "mymr";
        if (scalar >= 0x10A0 && scalar <= 0x10FF) return "geor";
        if (scalar >= 0x1200 && scalar <= 0x137F) return "ethi";
        if (scalar >= 0x1780 && scalar <= 0x17FF) return "khmr";
        if (scalar >= 0x1800 && scalar <= 0x18AF) return "mong";
        if (scalar >= 0x3040 && scalar <= 0x30FF) return "kana";
        if (scalar >= 0x3100 && scalar <= 0x312F) return "bopo";
        if ((scalar >= 0x3400 && scalar <= 0x4DBF) || (scalar >= 0x4E00 && scalar <= 0x9FFF)) return "hani";
        if ((scalar >= 0xAC00 && scalar <= 0xD7AF) || (scalar >= 0x1100 && scalar <= 0x11FF)) return "hang";
        return "DFLT";
    }

    private bool TryGposPairAdjustmentFromLookup(
        int lookupList,
        ushort lookupIndex,
        ushort left,
        ushort right,
        out OfficeOpenTypePairPositioning adjustment) {
        adjustment = default;
        int lookupCount = ReadUInt16(lookupList);
        if (lookupIndex >= lookupCount) return false;
        int lookupOffset = checked(lookupList + 2 + (lookupIndex * 2));
        if (!InBounds(lookupOffset, 2)) return false;
        int lookup = checked(lookupList + ReadUInt16(lookupOffset));
        if (!InBounds(lookup, 6)) return false;
        ushort lookupType = ReadUInt16(lookup);
        if (lookupType != 2 && lookupType != 9) return false;
        if (lookupType == 9 && !_includeExtendedGpos) return false;

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
            if (TryGposPairAdjustmentFromSubtable(subtable, left, right, out OfficeOpenTypePairPositioning subtableAdjustment)) {
                adjustment = subtableAdjustment;
                return true;
            }
        }
        return false;
    }

    private bool TryGposPairAdjustmentFromSubtable(
        int subtable,
        ushort left,
        ushort right,
        out OfficeOpenTypePairPositioning adjustment) {
        adjustment = default;
        if (!InBounds(subtable, 10)) return false;
        ushort format = ReadUInt16(subtable);
        if (format == 2) {
            return _includeExtendedGpos &&
                   TryGposClassPairAdjustment(subtable, left, right, out adjustment);
        }
        if (format != 1) return false;
        int coverage = checked(subtable + ReadUInt16(subtable + 2));
        ushort valueFormat1 = ReadUInt16(subtable + 4);
        ushort valueFormat2 = ReadUInt16(subtable + 6);
        int pairSetCount = ReadUInt16(subtable + 8);
        int coverageIndex = CoverageIndex(coverage, left);
        if (coverageIndex < 0 || coverageIndex >= pairSetCount) return false;

        int pairSetOffset = checked(subtable + 10 + (coverageIndex * 2));
        if (!InBounds(pairSetOffset, 2)) return false;
        int pairSet = checked(subtable + ReadUInt16(pairSetOffset));
        if (!InBounds(pairSet, 2)) return false;

        int value1Size = ValueRecordSize(valueFormat1);
        int value2Size = ValueRecordSize(valueFormat2);
        int recordSize = checked(2 + value1Size + value2Size);
        int low = 0;
        int high = ReadUInt16(pairSet) - 1;
        while (low <= high) {
            int mid = low + ((high - low) / 2);
            int record = checked(pairSet + 2 + (mid * recordSize));
            if (!InBounds(record, recordSize)) return false;
            ushort candidate = ReadUInt16(record);
            if (candidate == right) {
                OfficeOpenTypePairValue first = ReadValueRecordHorizontal(record + 2, valueFormat1);
                OfficeOpenTypePairValue second = ReadValueRecordHorizontal(record + 2 + value1Size, valueFormat2);
                adjustment = new OfficeOpenTypePairPositioning(
                    first.XPlacement,
                    first.XAdvance,
                    second.XPlacement,
                    second.XAdvance,
                    secondValueRecordPresent: valueFormat2 != 0);
                return true;
            }
            if (candidate < right) low = mid + 1;
            else high = mid - 1;
        }
        return false;
    }

    private bool TryGposClassPairAdjustment(int subtable, ushort left, ushort right, out OfficeOpenTypePairPositioning adjustment) {
        adjustment = default;
        if (!InBounds(subtable, 16)) return false;
        int coverage = checked(subtable + ReadUInt16(subtable + 2));
        if (CoverageIndex(coverage, left) < 0) return false;

        ushort valueFormat1 = ReadUInt16(subtable + 4);
        ushort valueFormat2 = ReadUInt16(subtable + 6);
        int classDef1 = checked(subtable + ReadUInt16(subtable + 8));
        int classDef2 = checked(subtable + ReadUInt16(subtable + 10));
        int class1Count = ReadUInt16(subtable + 12);
        int class2Count = ReadUInt16(subtable + 14);
        int class1 = ClassDefinition(classDef1, left);
        int class2 = ClassDefinition(classDef2, right);
        if (class1 < 0 || class2 < 0 || class1 >= class1Count || class2 >= class2Count) return false;

        int value1Size = ValueRecordSize(valueFormat1);
        int value2Size = ValueRecordSize(valueFormat2);
        int recordSize = checked(value1Size + value2Size);
        if (recordSize == 0) return false;
        int recordIndex = checked((class1 * class2Count) + class2);
        int record = checked(subtable + 16 + (recordIndex * recordSize));
        if (!InBounds(record, recordSize)) return false;
        OfficeOpenTypePairValue first = ReadValueRecordHorizontal(record, valueFormat1);
        OfficeOpenTypePairValue second = ReadValueRecordHorizontal(record + value1Size, valueFormat2);
        adjustment = new OfficeOpenTypePairPositioning(
            first.XPlacement,
            first.XAdvance,
            second.XPlacement,
            second.XAdvance,
            secondValueRecordPresent: valueFormat2 != 0);
        return true;
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

    private OfficeOpenTypePairValue ReadValueRecordHorizontal(int offset, ushort valueFormat) {
        int xPlacement = 0;
        int xAdvance = 0;
        if ((valueFormat & 0x0001) != 0) {
            xPlacement = ReadInt16(offset);
            offset += 2;
        }
        if ((valueFormat & 0x0002) != 0) offset += 2;
        if ((valueFormat & 0x0004) != 0) xAdvance = ReadInt16(offset);
        return new OfficeOpenTypePairValue(xPlacement, xAdvance);
    }

    private static int ValueRecordSize(ushort valueFormat) {
        int size = 0;
        for (int bit = 1; bit <= 0x0080; bit <<= 1) {
            if ((valueFormat & bit) != 0) size += 2;
        }
        return size;
    }

    private readonly struct OfficeOpenTypePairValue {
        internal OfficeOpenTypePairValue(int xPlacement, int xAdvance) {
            XPlacement = xPlacement;
            XAdvance = xAdvance;
        }

        internal int XPlacement { get; }
        internal int XAdvance { get; }
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
