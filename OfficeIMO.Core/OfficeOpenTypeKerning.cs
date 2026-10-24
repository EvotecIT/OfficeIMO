using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace OfficeIMO.Drawing;

internal readonly struct OfficeOpenTypePairPositioning {
    internal OfficeOpenTypePairPositioning(
        int firstGlyphXPlacement,
        int firstGlyphXAdvance,
        int secondGlyphXPlacement,
        int secondGlyphXAdvance,
        bool secondValueRecordPresent = false,
        int firstGlyphYPlacement = 0,
        int firstGlyphYAdvance = 0,
        int secondGlyphYPlacement = 0,
        int secondGlyphYAdvance = 0) {
        FirstGlyphXPlacement = firstGlyphXPlacement;
        FirstGlyphXAdvance = firstGlyphXAdvance;
        SecondGlyphXPlacement = secondGlyphXPlacement;
        SecondGlyphXAdvance = secondGlyphXAdvance;
        SecondValueRecordPresent = secondValueRecordPresent;
        FirstGlyphYPlacement = firstGlyphYPlacement;
        FirstGlyphYAdvance = firstGlyphYAdvance;
        SecondGlyphYPlacement = secondGlyphYPlacement;
        SecondGlyphYAdvance = secondGlyphYAdvance;
    }

    internal int FirstGlyphXPlacement { get; }
    internal int FirstGlyphXAdvance { get; }
    internal int SecondGlyphXPlacement { get; }
    internal int SecondGlyphXAdvance { get; }
    internal bool SecondValueRecordPresent { get; }
    internal int FirstGlyphYPlacement { get; }
    internal int FirstGlyphYAdvance { get; }
    internal int SecondGlyphYPlacement { get; }
    internal int SecondGlyphYAdvance { get; }
    internal int TotalXAdvance => checked(FirstGlyphXAdvance + SecondGlyphXAdvance);

    internal OfficeOpenTypePairPositioning Add(OfficeOpenTypePairPositioning other) => new(
        checked(FirstGlyphXPlacement + other.FirstGlyphXPlacement),
        checked(FirstGlyphXAdvance + other.FirstGlyphXAdvance),
        checked(SecondGlyphXPlacement + other.SecondGlyphXPlacement),
        checked(SecondGlyphXAdvance + other.SecondGlyphXAdvance),
        SecondValueRecordPresent || other.SecondValueRecordPresent,
        checked(FirstGlyphYPlacement + other.FirstGlyphYPlacement),
        checked(FirstGlyphYAdvance + other.FirstGlyphYAdvance),
        checked(SecondGlyphYPlacement + other.SecondGlyphYPlacement),
        checked(SecondGlyphYAdvance + other.SecondGlyphYAdvance));
}

internal readonly struct OfficeOpenTypeGlyphPositioning {
    internal OfficeOpenTypeGlyphPositioning(int xPlacement, int xAdvance, int yPlacement = 0, int yAdvance = 0) {
        XPlacement = xPlacement;
        XAdvance = xAdvance;
        YPlacement = yPlacement;
        YAdvance = yAdvance;
    }

    internal int XPlacement { get; }
    internal int XAdvance { get; }
    internal int YPlacement { get; }
    internal int YAdvance { get; }

    internal OfficeOpenTypeGlyphPositioning Add(int xPlacement, int xAdvance, int yPlacement = 0, int yAdvance = 0) => new(
        checked(XPlacement + xPlacement),
        checked(XAdvance + xAdvance),
        checked(YPlacement + yPlacement),
        checked(YAdvance + yAdvance));
}

/// <summary>Shared bounded legacy kern and GPOS pair-adjustment reader.</summary>
internal sealed partial class OfficeOpenTypeKerning {
    private readonly byte[] _data;
    private readonly int _kern;
    private readonly int _gpos;
    private readonly bool _includeExtendedGpos;
    private readonly Func<int, int, int>? _variationDelta;
    private readonly bool _requireVariationDelta;

    internal OfficeOpenTypeKerning(
        byte[] data,
        int kern,
        int gpos,
        bool includeExtendedGpos = true,
        Func<int, int, int>? variationDelta = null,
        bool requireVariationDelta = false) {
        _data = data ?? throw new ArgumentNullException(nameof(data));
        _kern = kern;
        _gpos = gpos;
        _includeExtendedGpos = includeExtendedGpos;
        _variationDelta = variationDelta;
        _requireVariationDelta = requireVariationDelta;
    }

    internal static OfficeOpenTypeKerning FromReader(
        OfficeOpenTypeReader reader,
        OfficeFontVariationModel? variations = null) {
        if (reader == null) throw new ArgumentNullException(nameof(reader));
        Func<int, int, int>? variationDelta = TryCreateGdefVariationEvaluator(
            reader,
            variations ?? OfficeFontVariationModel.None);
        return new OfficeOpenTypeKerning(
            reader.Data,
            reader.TryGetTable("kern", out int kern, out _) ? kern : -1,
            reader.TryGetTable("GPOS", out int gpos, out _) ? gpos : -1,
            includeExtendedGpos: true,
            variationDelta,
            requireVariationDelta: variations?.IsVariable == true);
    }

    private static Func<int, int, int>? TryCreateGdefVariationEvaluator(
        OfficeOpenTypeReader reader,
        OfficeFontVariationModel variations) {
        if (!variations.IsVariable || !reader.TryGetTable("GDEF", out int gdef, out int length)) return null;
        int end = checked(gdef + length);
        if (length < 4 || reader.ReadUInt16(gdef) != 1) {
            throw new InvalidDataException("The OpenType GDEF table header is invalid.");
        }
        int minorVersion = reader.ReadUInt16(gdef + 2);
        if (minorVersion < 3) return null;
        if (length < 18) throw new InvalidDataException("The OpenType GDEF variation-store header is truncated.");
        uint relativeOffset = reader.ReadUInt32(gdef + 14);
        if (relativeOffset == 0) return null;
        if (relativeOffset > int.MaxValue) {
            throw new InvalidDataException("The OpenType GDEF variation-store offset is invalid.");
        }
        int storeOffset = checked(gdef + (int)relativeOffset);
        if (storeOffset < gdef || storeOffset >= end) {
            throw new InvalidDataException("The OpenType GDEF variation-store offset is invalid.");
        }
        OfficeOpenTypeItemVariationStore store = OfficeOpenTypeItemVariationStore.Parse(
            reader,
            storeOffset,
            end,
            variations);
        return store.Evaluate;
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
        var budget = new PositioningWorkBudget(CancellationToken.None);
        return TryGposPairAdjustment((ushort)left, (ushort)right, scriptTag, budget, out OfficeOpenTypePairPositioning gposAdjustment)
            ? gposAdjustment
            : new OfficeOpenTypePairPositioning(0, KernPairAdjustment((ushort)left, (ushort)right, budget), 0, 0);
    }

    internal OfficeOpenTypeGlyphPositioning[] PositionRun(
        IReadOnlyList<int> glyphs,
        IReadOnlyList<int> scalars,
        CancellationToken cancellationToken = default,
        bool[]? synthesizedSpaces = null) {
        if (glyphs == null) throw new ArgumentNullException(nameof(glyphs));
        if (scalars == null) throw new ArgumentNullException(nameof(scalars));
        if (glyphs.Count != scalars.Count) throw new ArgumentException("Glyph and scalar runs must have the same length.");
        if (synthesizedSpaces != null && synthesizedSpaces.Length != glyphs.Count)
            throw new ArgumentException("Synthesized-space flags must match the glyph run.", nameof(synthesizedSpaces));

        var result = new OfficeOpenTypeGlyphPositioning[glyphs.Count];
        if (glyphs.Count < 2) return result;
        for (int index = 0; index < glyphs.Count; index++) {
            if ((uint)glyphs[index] > ushort.MaxValue) return result;
        }

        int pairCount = glyphs.Count - 1;
        var budget = new PositioningWorkBudget(cancellationToken);
        var pairLookups = new HashSet<ushort>[pairCount];
        var orderedLookups = new List<ushort>();
        var knownLookups = new HashSet<ushort>();
        var lookupsByScript = new Dictionary<string, HashSet<ushort>>(StringComparer.Ordinal);
        bool hasGposLayout = TryGetGposLayoutTables(out int scriptList, out int featureList, out int lookupList);
        if (hasGposLayout) {
            for (int index = 0; index < pairCount; index++) {
                if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (OfficeMissingSpaceMetrics.IsSynthesized(synthesizedSpaces, index) ||
                    OfficeMissingSpaceMetrics.IsSynthesized(synthesizedSpaces, index + 1)) {
                    pairLookups[index] = new HashSet<ushort>();
                    continue;
                }
                string scriptTag = ResolveScriptTag(scalars[index], scalars[index + 1]);
                if (!lookupsByScript.TryGetValue(scriptTag, out HashSet<ushort>? lookups)) {
                    lookups = new HashSet<ushort>();
                    foreach (ushort lookupIndex in GposFeatureLookupIndexes(scriptList, featureList, "kern", scriptTag, budget)) {
                        if (!lookups.Add(lookupIndex)) continue;
                        if (knownLookups.Add(lookupIndex)) orderedLookups.Add(lookupIndex);
                    }
                    lookupsByScript.Add(scriptTag, lookups);
                }
                pairLookups[index] = lookups;
            }
        }

        var gposApplied = new bool[pairCount];
        foreach (ushort lookupIndex in orderedLookups) {
            for (int index = 0; index < pairCount;) {
                budget.Charge();
                HashSet<ushort> activeLookups = pairLookups[index];
                if (activeLookups.Contains(lookupIndex) &&
                    TryGposPairAdjustmentFromLookup(
                        lookupList,
                        lookupIndex,
                        (ushort)glyphs[index],
                        (ushort)glyphs[index + 1],
                        budget,
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
            if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
            // A synthesized blank is not glyph zero: its advance must not pick
            // up unrelated .notdef pair positioning on either side.
            if (OfficeMissingSpaceMetrics.IsSynthesized(synthesizedSpaces, index) ||
                OfficeMissingSpaceMetrics.IsSynthesized(synthesizedSpaces, index + 1)) continue;
            if (gposApplied[index]) continue;
            int legacyAdjustment = KernPairAdjustment((ushort)glyphs[index], (ushort)glyphs[index + 1], budget);
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

    private int KernPairAdjustment(ushort left, ushort right, PositioningWorkBudget budget) {
        if (_kern < 0 || !InBounds(_kern, 4) || ReadUInt16(_kern) != 0) return 0;
        int count = ReadUInt16(_kern + 2);
        int position = _kern + 4;
        int adjustment = 0;
        for (int table = 0; table < count; table++) {
            budget.Charge();
            if (!InBounds(position, 6)) break;
            int length = ReadUInt16(position + 2);
            int coverage = ReadUInt16(position + 4);
            int next = checked(position + length);
            if (length < 14 || next <= position || next > _data.Length) break;
            bool isFormat0 = (coverage >> 8) == 0;
            bool isHorizontalOrdinary = (coverage & 0x0007) == 0x0001;
            if (isFormat0 && isHorizontalOrdinary) {
                int value = KerningFormat0(position, left, right, budget);
                adjustment = (coverage & 0x0008) != 0
                    ? value
                    : checked(adjustment + value);
            }
            position = next;
        }
        return adjustment;
    }

    private int KerningFormat0(int table, ushort left, ushort right, PositioningWorkBudget budget) {
        int pairs = ReadUInt16(table + 6);
        int pairOffset = table + 14;
        uint key = ((uint)left << 16) | right;
        int low = 0;
        int high = pairs - 1;
        while (low <= high) {
            budget.Charge();
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

    private bool TryGposPairAdjustment(ushort left, ushort right, string scriptTag,
        PositioningWorkBudget budget, out OfficeOpenTypePairPositioning adjustment) {
        adjustment = default;
        if (!TryGetGposLayoutTables(out int scriptList, out int featureList, out int lookupList)) return false;

        bool applied = false;
        var seen = new HashSet<ushort>();
        foreach (ushort lookupIndex in GposFeatureLookupIndexes(scriptList, featureList, "kern", scriptTag, budget)) {
            if (seen.Add(lookupIndex) &&
                TryGposPairAdjustmentFromLookup(lookupList, lookupIndex, left, right, budget,
                    out OfficeOpenTypePairPositioning lookupAdjustment)) {
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
        string scriptTag,
        PositioningWorkBudget budget,
        string? language = null) {
        int script = FindScript(scriptList, scriptTag, budget);
        if (script < 0 && !string.Equals(scriptTag, "DFLT", StringComparison.Ordinal)) {
            script = FindScript(scriptList, "DFLT", budget);
        }
        if (script < 0 || !InPositioningBounds(script, 4, budget)) yield break;

        int langSys = FindPositioningLanguageSystem(script, language, budget);
        if (langSys < 0) yield break;

        int featureCount = ReadUInt16(featureList);
        var featureIndexes = new List<ushort>();
        var seenFeatures = new HashSet<ushort>();
        ushort requiredFeature = ReadUInt16(langSys + 2);
        if (requiredFeature != ushort.MaxValue && seenFeatures.Add(requiredFeature)) featureIndexes.Add(requiredFeature);
        int langSysFeatureCount = ReadUInt16(langSys + 4);
        for (int index = 0; index < langSysFeatureCount; index++) {
            budget.Charge();
            int offset = checked(langSys + 6 + (index * 2));
            if (!InPositioningBounds(offset, 2, budget)) yield break;
            ushort featureIndex = ReadUInt16(offset);
            if (seenFeatures.Add(featureIndex)) featureIndexes.Add(featureIndex);
        }

        foreach (ushort featureIndex in featureIndexes) {
            budget.Charge();
            if (featureIndex >= featureCount) continue;
            int record = checked(featureList + 2 + (featureIndex * 6));
            if (!InPositioningBounds(record, 6, budget)) yield break;
            if (!TagEquals(record, featureTag)) continue;
            int feature = checked(featureList + ReadUInt16(record + 4));
            if (!InPositioningBounds(feature, 4, budget)) yield break;
            int lookupCount = ReadUInt16(feature + 2);
            for (int lookup = 0; lookup < lookupCount; lookup++) {
                budget.Charge();
                int offset = checked(feature + 4 + (lookup * 2));
                if (!InPositioningBounds(offset, 2, budget)) yield break;
                yield return ReadUInt16(offset);
            }
        }
    }

    private int FindScript(int scriptList, string scriptTag, PositioningWorkBudget budget) {
        int scriptCount = ReadUInt16(scriptList);
        for (int index = 0; index < scriptCount; index++) {
            budget.Charge();
            int record = checked(scriptList + 2 + (index * 6));
            if (!InPositioningBounds(record, 6, budget)) return -1;
            if (!TagEquals(record, scriptTag)) continue;
            int script = checked(scriptList + ReadUInt16(record + 4));
            return InPositioningBounds(script, 4, budget) ? script : -1;
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
        PositioningWorkBudget budget,
        out OfficeOpenTypePairPositioning adjustment) {
        adjustment = default;
        int lookupCount = ReadUInt16(lookupList);
        if (lookupIndex >= lookupCount) return false;
        int lookupOffset = checked(lookupList + 2 + (lookupIndex * 2));
        if (!InPositioningBounds(lookupOffset, 2, budget)) return false;
        int lookup = checked(lookupList + ReadUInt16(lookupOffset));
        if (!InPositioningBounds(lookup, 6, budget)) return false;
        ushort lookupType = ReadUInt16(lookup);
        if (lookupType != 2 && lookupType != 9) return false;
        if (lookupType == 9 && !_includeExtendedGpos) return false;

        int subtableCount = ReadUInt16(lookup + 4);
        for (int index = 0; index < subtableCount; index++) {
            budget.Charge();
            int subtableOffset = checked(lookup + 6 + (index * 2));
            if (!InPositioningBounds(subtableOffset, 2, budget)) break;
            int subtable = checked(lookup + ReadUInt16(subtableOffset));
            if (lookupType == 9) {
                if (!InPositioningBounds(subtable, 8, budget) || ReadUInt16(subtable) != 1 || ReadUInt16(subtable + 2) != 2) continue;
                uint extensionOffset = ReadUInt32(subtable + 4);
                if (extensionOffset > int.MaxValue) continue;
                subtable = checked(subtable + (int)extensionOffset);
            }
            if (TryGposPairAdjustmentFromSubtable(subtable, left, right, budget,
                    out OfficeOpenTypePairPositioning subtableAdjustment)) {
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
        PositioningWorkBudget budget,
        out OfficeOpenTypePairPositioning adjustment) {
        adjustment = default;
        if (!InPositioningBounds(subtable, 10, budget)) return false;
        ushort format = ReadUInt16(subtable);
        if (format == 2) {
            return _includeExtendedGpos &&
                   TryGposClassPairAdjustment(subtable, left, right, budget, out adjustment);
        }
        if (format != 1) return false;
        int coverage = checked(subtable + ReadUInt16(subtable + 2));
        ushort valueFormat1 = ReadUInt16(subtable + 4);
        ushort valueFormat2 = ReadUInt16(subtable + 6);
        int pairSetCount = ReadUInt16(subtable + 8);
        int coverageIndex = CoverageIndex(coverage, left, budget);
        if (coverageIndex < 0 || coverageIndex >= pairSetCount) return false;

        int pairSetOffset = checked(subtable + 10 + (coverageIndex * 2));
        if (!InPositioningBounds(pairSetOffset, 2, budget)) return false;
        int pairSet = checked(subtable + ReadUInt16(pairSetOffset));
        if (!InPositioningBounds(pairSet, 2, budget)) return false;

        int value1Size = ValueRecordSize(valueFormat1);
        int value2Size = ValueRecordSize(valueFormat2);
        int recordSize = checked(2 + value1Size + value2Size);
        int low = 0;
        int high = ReadUInt16(pairSet) - 1;
        while (low <= high) {
            budget.Charge();
            int mid = low + ((high - low) / 2);
            int record = checked(pairSet + 2 + (mid * recordSize));
            if (!InPositioningBounds(record, recordSize, budget)) return false;
            ushort candidate = ReadUInt16(record);
            if (candidate == right) {
                OfficeOpenTypePairValue first = ReadValueRecord(pairSet, record + 2, valueFormat1, budget);
                OfficeOpenTypePairValue second = ReadValueRecord(pairSet, record + 2 + value1Size, valueFormat2, budget);
                adjustment = new OfficeOpenTypePairPositioning(
                    first.XPlacement,
                    first.XAdvance,
                    second.XPlacement,
                    second.XAdvance,
                    secondValueRecordPresent: valueFormat2 != 0,
                    first.YPlacement, first.YAdvance, second.YPlacement, second.YAdvance);
                return true;
            }
            if (candidate < right) low = mid + 1;
            else high = mid - 1;
        }
        return false;
    }

    private bool TryGposClassPairAdjustment(int subtable, ushort left, ushort right,
        PositioningWorkBudget budget, out OfficeOpenTypePairPositioning adjustment) {
        adjustment = default;
        if (!InPositioningBounds(subtable, 16, budget)) return false;
        int coverage = checked(subtable + ReadUInt16(subtable + 2));
        if (CoverageIndex(coverage, left, budget) < 0) return false;

        ushort valueFormat1 = ReadUInt16(subtable + 4);
        ushort valueFormat2 = ReadUInt16(subtable + 6);
        int classDef1 = checked(subtable + ReadUInt16(subtable + 8));
        int classDef2 = checked(subtable + ReadUInt16(subtable + 10));
        int class1Count = ReadUInt16(subtable + 12);
        int class2Count = ReadUInt16(subtable + 14);
        int class1 = ClassDefinition(classDef1, left, budget);
        int class2 = ClassDefinition(classDef2, right, budget);
        if (class1 < 0 || class2 < 0 || class1 >= class1Count || class2 >= class2Count) return false;

        int value1Size = ValueRecordSize(valueFormat1);
        int value2Size = ValueRecordSize(valueFormat2);
        int recordSize = checked(value1Size + value2Size);
        if (recordSize == 0) return false;
        int recordIndex = checked((class1 * class2Count) + class2);
        int record = checked(subtable + 16 + (recordIndex * recordSize));
        if (!InPositioningBounds(record, recordSize, budget)) return false;
        OfficeOpenTypePairValue first = ReadValueRecord(subtable, record, valueFormat1, budget);
        OfficeOpenTypePairValue second = ReadValueRecord(subtable, record + value1Size, valueFormat2, budget);
        adjustment = new OfficeOpenTypePairPositioning(
            first.XPlacement,
            first.XAdvance,
            second.XPlacement,
            second.XAdvance,
            secondValueRecordPresent: valueFormat2 != 0,
            first.YPlacement, first.YAdvance, second.YPlacement, second.YAdvance);
        return true;
    }

    private int ClassDefinition(int classDef, ushort glyph, PositioningWorkBudget budget) {
        if (!InPositioningBounds(classDef, 4, budget)) return -1;
        ushort format = ReadUInt16(classDef);
        if (format == 1) {
            ushort startGlyph = ReadUInt16(classDef + 2);
            if (!InPositioningBounds(classDef, 6, budget)) return -1;
            int glyphCount = ReadUInt16(classDef + 4);
            int index = glyph - startGlyph;
            if (index < 0 || index >= glyphCount) return 0;
            int valueOffset = checked(classDef + 6 + (index * 2));
            return InPositioningBounds(valueOffset, 2, budget) ? ReadUInt16(valueOffset) : -1;
        }

        if (format != 2) return -1;
        int rangeCount = ReadUInt16(classDef + 2);
        int low = 0;
        int high = rangeCount - 1;
        while (low <= high) {
            budget.Charge();
            int mid = low + ((high - low) / 2);
            int range = checked(classDef + 4 + (mid * 6));
            if (!InPositioningBounds(range, 6, budget)) return -1;
            ushort start = ReadUInt16(range);
            ushort end = ReadUInt16(range + 2);
            if (glyph < start) high = mid - 1;
            else if (glyph > end) low = mid + 1;
            else return ReadUInt16(range + 4);
        }
        return 0;
    }

    private int CoverageIndex(int coverage, ushort glyph, PositioningWorkBudget budget) {
        if (!InPositioningBounds(coverage, 4, budget)) return -1;
        ushort format = ReadUInt16(coverage);
        if (format == 1) {
            int count = ReadUInt16(coverage + 2);
            int low = 0;
            int high = count - 1;
            while (low <= high) {
                budget.Charge();
                int mid = low + ((high - low) / 2);
                int offset = checked(coverage + 4 + (mid * 2));
                if (!InPositioningBounds(offset, 2, budget)) return -1;
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
            budget.Charge();
            int range = checked(coverage + 4 + (index * 6));
            if (!InPositioningBounds(range, 6, budget)) return -1;
            ushort start = ReadUInt16(range);
            ushort end = ReadUInt16(range + 2);
            if (glyph < start || glyph > end) continue;
            return checked(ReadUInt16(range + 4) + glyph - start);
        }
        return -1;
    }

    private OfficeOpenTypePairValue ReadValueRecord(int subtable, int offset, ushort valueFormat, PositioningWorkBudget budget) {
        if ((valueFormat & 0xFF00) != 0 || !InPositioningBounds(offset, ValueRecordSize(valueFormat), budget))
            throw new InvalidDataException("A GPOS value record is unsupported or truncated.");
        int xPlacement = 0;
        int xAdvance = 0;
        int yPlacement = 0;
        int yAdvance = 0;
        if ((valueFormat & 0x0001) != 0) {
            xPlacement = ReadInt16(offset);
            offset += 2;
        }
        if ((valueFormat & 0x0002) != 0) {
            yPlacement = ReadInt16(offset);
            offset += 2;
        }
        if ((valueFormat & 0x0004) != 0) {
            xAdvance = ReadInt16(offset);
            offset += 2;
        }
        if ((valueFormat & 0x0008) != 0) {
            yAdvance = ReadInt16(offset);
            offset += 2;
        }
        if ((valueFormat & 0x0010) != 0) {
            xPlacement = checked(xPlacement + ReadVariationAdjustment(subtable, ReadUInt16(offset), budget));
            offset += 2;
        }
        if ((valueFormat & 0x0020) != 0) {
            yPlacement = checked(yPlacement + ReadVariationAdjustment(subtable, ReadUInt16(offset), budget));
            offset += 2;
        }
        if ((valueFormat & 0x0040) != 0) {
            xAdvance = checked(xAdvance + ReadVariationAdjustment(subtable, ReadUInt16(offset), budget));
            offset += 2;
        }
        if ((valueFormat & 0x0080) != 0) {
            yAdvance = checked(yAdvance + ReadVariationAdjustment(subtable, ReadUInt16(offset), budget));
        }
        return new OfficeOpenTypePairValue(xPlacement, xAdvance, yPlacement, yAdvance);
    }

    private int ReadVariationAdjustment(int subtable, ushort relativeOffset, PositioningWorkBudget budget) {
        if (relativeOffset == 0) return 0;
        int variationIndex = checked(subtable + relativeOffset);
        if (!InPositioningBounds(variationIndex, 6, budget)) {
            throw new InvalidDataException("A GPOS device or variation-index table is truncated.");
        }
        if (ReadUInt16(variationIndex + 4) != 0x8000) return 0;
        if (_variationDelta == null) {
            if (_requireVariationDelta) {
                throw new InvalidDataException("A GPOS VariationIndex requires a GDEF ItemVariationStore.");
            }
            return 0;
        }
        return _variationDelta(ReadUInt16(variationIndex), ReadUInt16(variationIndex + 2));
    }

    private static int ValueRecordSize(ushort valueFormat) {
        int size = 0;
        for (int bit = 1; bit <= 0x0080; bit <<= 1) {
            if ((valueFormat & bit) != 0) size += 2;
        }
        return size;
    }

    private sealed class PositioningWorkBudget {
        private const long MaximumOperations = 8_000_000L;
        private readonly CancellationToken _cancellationToken;
        private long _operations;
        private readonly int _positioningStart;
        private readonly int _positioningEnd;

        internal PositioningWorkBudget(CancellationToken cancellationToken, int positioningStart = 0, int positioningEnd = int.MaxValue) {
            _cancellationToken = cancellationToken;
            _positioningStart = positioningStart;
            _positioningEnd = positioningEnd;
        }

        internal bool Contains(int offset, int count) => count >= 0 && offset >= _positioningStart
            && offset <= _positioningEnd - count;

        internal void Charge() {
            if (++_operations > MaximumOperations) {
                throw new InvalidDataException("OpenType positioning exceeded the managed work budget.");
            }
            if ((_operations & 1023L) == 0L) _cancellationToken.ThrowIfCancellationRequested();
        }
    }

    private readonly struct OfficeOpenTypePairValue {
        internal OfficeOpenTypePairValue(int xPlacement, int xAdvance, int yPlacement = 0, int yAdvance = 0) {
            XPlacement = xPlacement;
            XAdvance = xAdvance;
            YPlacement = yPlacement;
            YAdvance = yAdvance;
        }

        internal int XPlacement { get; }
        internal int XAdvance { get; }
        internal int YPlacement { get; }
        internal int YAdvance { get; }
    }

    private bool TagEquals(int offset, string tag) =>
        InBounds(offset, 4) &&
        _data[offset] == tag[0] &&
        _data[offset + 1] == tag[1] &&
        _data[offset + 2] == tag[2] &&
        _data[offset + 3] == tag[3];

    private bool InPositioningBounds(int offset, int count, PositioningWorkBudget budget) =>
        InBounds(offset, count) && budget.Contains(offset, count);

    private bool InBounds(int offset, int count) => offset >= 0 && count >= 0 && offset <= _data.Length - count;
    private ushort ReadUInt16(int offset) => (ushort)((_data[offset] << 8) | _data[offset + 1]);
    private short ReadInt16(int offset) => unchecked((short)ReadUInt16(offset));
    private uint ReadUInt32(int offset) =>
        ((uint)_data[offset] << 24) |
        ((uint)_data[offset + 1] << 16) |
        ((uint)_data[offset + 2] << 8) |
        _data[offset + 3];
}
