using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.TestAssets;

internal static partial class ManagedTextShapingTestAssets {
    /// <summary>A valid font that selects many identity substitutions for each isolated Arabic segment.</summary>
    internal static byte[] CreateAlternatingScriptBudgetFont(int lookupCount = 256) {
        var data = new MarkTableBuilder(); data.U16(1); data.U16(0);
        int scriptsOffset = data.Reserve(), featuresOffset = data.Reserve(), lookupsOffset = data.Reserve();
        data.Offset(scriptsOffset, 0); int scriptList = data.Position;
        data.U16(1); data.Tag("arab"); int scriptOffset = data.Reserve();
        data.Offset(scriptOffset, scriptList); int script = data.Position; int language = data.Reserve(); data.U16(0);
        data.Offset(language, script); data.U16(0); data.U16(ushort.MaxValue); data.U16(1); data.U16(0);
        data.Offset(featuresOffset, 0); int features = data.Position;
        data.U16(1); data.Tag("isol"); int feature = data.Reserve();
        data.Offset(feature, features); data.U16(0); data.U16(lookupCount);
        for (int index = 0; index < lookupCount; index++) data.U16(index);
        data.Offset(lookupsOffset, 0); int lookupList = data.Position; data.U16(lookupCount);
        int[] lookups = new int[lookupCount];
        for (int index = 0; index < lookupCount; index++) lookups[index] = data.Reserve();
        for (int index = 0; index < lookupCount; index++) {
            data.Offset(lookups[index], lookupList);
            data.Bytes(new byte[] { 0, 1, 0, 0, 0, 1, 0, 8, 0, 2, 0, 8, 0, 1, 0, 1, 0, 1, 0, 1, 0, 1 });
        }
        return CreateFontFromCmap(CreateFormat12Cmap(0x0628, 1, 'A', 2), glyphCount: 3, gsub: data.ToArray());
    }

    /// <summary>Valid GPOS combinations whose lookup ordering and mark adjustments are observable through the public provider.</summary>
    internal static byte[] CreateMarkCompositionFont(int valueFormat = 0, bool pair = false,
        bool stackBeforeBase = false, string extraFeature = "kern", bool requiredExtraFeature = false) {
        var lookups = new List<byte[]>();
        if (stackBeforeBase) lookups.Add(AttachmentLookup(6, AttachmentSubtable(6, false, false), false, 0));
        int baseLookup = lookups.Count;
        lookups.Add(AttachmentLookup(4, AttachmentSubtable(4, false, false), false, 0));
        int adjustmentLookup = -1;
        if (valueFormat != 0) {
            adjustmentLookup = lookups.Count;
            var subtable = new MarkTableBuilder();
            subtable.U16(1); int coverage = subtable.Reserve();
            subtable.U16(pair ? 0 : valueFormat);
            if (pair) {
                subtable.U16(valueFormat); subtable.U16(1); int pairOffset = subtable.Reserve();
                subtable.Offset(pairOffset, 0); subtable.U16(1); subtable.U16(2); subtable.U16(77);
                subtable.Offset(coverage, 0); subtable.Coverage(new[] { 1 });
            } else {
                subtable.U16(77); subtable.Offset(coverage, 0); subtable.Coverage(new[] { 2 });
            }
            lookups.Add(AttachmentLookup((ushort)(pair ? 2 : 1), subtable.ToArray(), false, 0));
        }
        var features = new SortedDictionary<string, int[]>(StringComparer.Ordinal) { ["mark"] = new[] { baseLookup } };
        if (stackBeforeBase) features.Add("mkmk", new[] { 0 });
        if (adjustmentLookup >= 0) features.Add(extraFeature, new[] { adjustmentLookup });
        var data = new MarkTableBuilder();
        data.U16(1); data.U16(0); int scriptsOffset = data.Reserve(), featuresOffset = data.Reserve(), lookupsOffset = data.Reserve();
        data.Offset(scriptsOffset, 0); int scriptList = data.Position;
        data.U16(1); data.Tag("DFLT"); int scriptOffset = data.Reserve();
        data.Offset(scriptOffset, scriptList); int script = data.Position; int language = data.Reserve(); data.U16(0);
        data.Offset(language, script); data.U16(0);
        string[] tags = features.Keys.ToArray();
        data.U16(requiredExtraFeature ? Array.IndexOf(tags, extraFeature) : ushort.MaxValue);
        data.U16(tags.Length); for (int index = 0; index < tags.Length; index++) data.U16(index);
        data.Offset(featuresOffset, 0); int featureList = data.Position; data.U16(tags.Length);
        int[] featureOffsets = new int[tags.Length];
        for (int index = 0; index < tags.Length; index++) { data.Tag(tags[index]); featureOffsets[index] = data.Reserve(); }
        for (int index = 0; index < tags.Length; index++) {
            data.Offset(featureOffsets[index], featureList); data.U16(0); data.U16(features[tags[index]].Length);
            foreach (int lookup in features[tags[index]]) data.U16(lookup);
        }
        data.Offset(lookupsOffset, 0); int lookupList = data.Position; data.U16(lookups.Count);
        int[] offsets = lookups.Select(_ => data.Reserve()).ToArray();
        for (int index = 0; index < lookups.Count; index++) { data.Offset(offsets[index], lookupList); data.Bytes(lookups[index]); }
        return CreateFontFromCmap(CreateDistinctFormat12Cmap(new[] { (int)'A', 0x0301, 0x0302, 0x0303 }),
            glyphCount: 6, gpos: data.ToArray(), additionalTables: new[] { ("GDEF", MarkFilterGdef(false)) });
    }
}
