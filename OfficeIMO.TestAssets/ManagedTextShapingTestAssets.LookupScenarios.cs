using System;
using System.Linq;

namespace OfficeIMO.TestAssets;

internal static partial class ManagedTextShapingTestAssets {
    internal static byte[] CreateFontWithInheritedMarkSubstitution() {
        byte[] gsub = CreateMultipleGsub(scriptTag: "latn");
        WriteUInt16(gsub, 74, 0);
        WriteUInt16(gsub, 52, 5);
        return CreateFontFromCmap(CreateDistinctFormat12Cmap(new[] { (int)',', (int)'1', (int)'A', 0x00AD, 0x0301, 0x03B1, 0x0410 }), glyphCount: 9, gsub: gsub);
    }

    internal static byte[] CreateFontWithOrderedLigatures(bool longestFirst) {
        byte[] original = CreateLigatureGsub("liga", 1, 2, 4, "latn");
        var gsub = new byte[92];
        Array.Copy(original, 0, gsub, 0, 46);
        Array.Copy(original, 56, gsub, 66, 26);
        WriteUInt16(gsub, 4, 72); WriteUInt16(gsub, 40, 28);
        WriteUInt16(gsub, 46, 2);
        WriteUInt16(gsub, 48, (ushort)(longestFirst ? 12 : 6));
        WriteUInt16(gsub, 50, (ushort)(longestFirst ? 6 : 12));
        WriteUInt16(gsub, 52, 4); WriteUInt16(gsub, 54, 2); WriteUInt16(gsub, 56, 2);
        WriteUInt16(gsub, 58, 5); WriteUInt16(gsub, 60, 3); WriteUInt16(gsub, 62, 2); WriteUInt16(gsub, 64, 3);
        return CreateFontFromCmap(CreateDistinctFormat12Cmap(new[] { (int)'f', (int)'i', (int)'n' }), glyphCount: 7, gsub: gsub);
    }
    internal static byte[] CreateMixedWhitespaceLigatureFont() => CreateFontFromCmap(
        CreateDistinctFormat12Cmap(new[] { 32, (int)'A', (int)'B' }), glyphCount: 5,
        gsub: CreateLigatureGsub("liga", 2, 1, 4, "latn", 0));
    internal static byte[] CreateMixedLatinLigatureFont(int neighbor) => CreateFontFromCmap(
        CreateDistinctFormat12Cmap(new[] { 32, (int)'f', (int)'i', neighbor }), glyphCount: 6,
        gsub: CreateLigatureGsub("liga", 2, 3, 5, "latn", 0));
    internal static byte[] CreateFontWithLookupScanningScenario(bool nested, bool extension, bool insertedOutput) {
        byte[] multiple = insertedOutput
            ? MultipleScenarioSubtable(new ushort[] { 1, 2 }, new[] { new ushort[] { 3, 2 }, new ushort[] { 4, 5 } })
            : MultipleScenarioSubtable(new ushort[] { 1 }, new[] { new ushort[] { 3, 4 } });
        byte[] lookup = ScenarioLookup(2, extension, insertedOutput ? new[] { multiple } : new[] {
            multiple, MultipleScenarioSubtable(new ushort[] { 3 }, new[] { new ushort[] { 5, 6 } }) });
        byte[][] lookups = nested ? new[] {
            ScenarioLookup(5, extension, new[] { ContextScenarioSubtable(insertedOutput ? new ushort[] { 1, 2 } : new ushort[] { 1 }) }), lookup
        } : new[] { lookup };
        return CreateLookupScenarioFont(lookups);
    }

    internal static byte[] CreateFontWithContextInputOrLookahead(bool lookahead) {
        var context = new byte[lookahead ? 30 : 26];
        WriteUInt16(context, 0, 3);
        if (lookahead) {
            WriteUInt16(context, 4, 1); WriteUInt16(context, 6, 18);
            WriteUInt16(context, 8, 1); WriteUInt16(context, 10, 24);
            WriteUInt16(context, 12, 1); WriteUInt16(context, 16, 1);
            WriteUInt16(context, 18, 1); WriteUInt16(context, 20, 1); WriteUInt16(context, 22, 1);
            WriteUInt16(context, 24, 1); WriteUInt16(context, 26, 1); WriteUInt16(context, 28, 1);
        } else {
            WriteUInt16(context, 2, 2); WriteUInt16(context, 4, 1);
            WriteUInt16(context, 6, 14); WriteUInt16(context, 8, 20); WriteUInt16(context, 12, 1);
            WriteUInt16(context, 14, 1); WriteUInt16(context, 16, 1); WriteUInt16(context, 18, 1);
            WriteUInt16(context, 20, 1); WriteUInt16(context, 22, 1); WriteUInt16(context, 24, 1);
        }
        var single = new byte[14];
        WriteUInt16(single, 0, 2); WriteUInt16(single, 2, 8); WriteUInt16(single, 4, 1);
        WriteUInt16(single, 6, 2); WriteUInt16(single, 8, 1); WriteUInt16(single, 10, 1); WriteUInt16(single, 12, 1);
        return CreateLookupScenarioFont(new[] { ScenarioLookup(lookahead ? (ushort)6 : (ushort)5, false, new[] { context }),
            ScenarioLookup(1, false, new[] { single }) });
    }

    private static byte[] CreateLookupScenarioFont(byte[][] lookups) {
        int size = 26 + 2 + lookups.Length * 2 + lookups.Sum(item => item.Length);
        var gsub = new byte[size];
        WriteUInt32(gsub, 0, 0x00010000); WriteUInt16(gsub, 4, 10); WriteUInt16(gsub, 6, 12); WriteUInt16(gsub, 8, 26);
        WriteUInt16(gsub, 12, 1); WriteTag(gsub, 14, "calt"); WriteUInt16(gsub, 18, 8); WriteUInt16(gsub, 22, 1);
        WriteUInt16(gsub, 26, (ushort)lookups.Length);
        int cursor = 28 + lookups.Length * 2;
        for (int index = 0; index < lookups.Length; index++) {
            WriteUInt16(gsub, 28 + index * 2, (ushort)(cursor - 26));
            Array.Copy(lookups[index], 0, gsub, cursor, lookups[index].Length); cursor += lookups[index].Length;
        }
        return CreateFontFromCmap(CreateFormat12Cmap('A', 1, 'B', 2), glyphCount: 7, gsub: gsub);
    }

    private static byte[] ScenarioLookup(ushort type, bool extension, byte[][] subtables) {
        int header = 6 + subtables.Length * 2;
        var lookup = new byte[header + subtables.Sum(s => s.Length + (extension ? 8 : 0))];
        WriteUInt16(lookup, 0, extension ? (ushort)7 : type); WriteUInt16(lookup, 4, (ushort)subtables.Length);
        int cursor = header;
        for (int index = 0; index < subtables.Length; index++) {
            WriteUInt16(lookup, 6 + index * 2, (ushort)cursor);
            if (extension) {
                WriteUInt16(lookup, cursor, 1); WriteUInt16(lookup, cursor + 2, type); WriteUInt32(lookup, cursor + 4, 8);
                cursor += 8;
            }
            Array.Copy(subtables[index], 0, lookup, cursor, subtables[index].Length); cursor += subtables[index].Length;
        }
        return lookup;
    }

    private static byte[] MultipleScenarioSubtable(ushort[] coverage, ushort[][] replacements) {
        int coverageOffset = 6 + coverage.Length * 2;
        int cursor = coverageOffset + 4 + coverage.Length * 2;
        var data = new byte[cursor + replacements.Sum(r => 2 + r.Length * 2)];
        WriteUInt16(data, 0, 1); WriteUInt16(data, 2, (ushort)coverageOffset); WriteUInt16(data, 4, (ushort)coverage.Length);
        WriteUInt16(data, coverageOffset, 1); WriteUInt16(data, coverageOffset + 2, (ushort)coverage.Length);
        for (int index = 0; index < coverage.Length; index++) {
            WriteUInt16(data, coverageOffset + 4 + index * 2, coverage[index]);
            WriteUInt16(data, 6 + index * 2, (ushort)cursor); WriteUInt16(data, cursor, (ushort)replacements[index].Length);
            for (int item = 0; item < replacements[index].Length; item++) WriteUInt16(data, cursor + 2 + item * 2, replacements[index][item]);
            cursor += 2 + replacements[index].Length * 2;
        }
        return data;
    }

    private static byte[] ContextScenarioSubtable(ushort[] coverage) {
        var data = new byte[16 + coverage.Length * 2];
        WriteUInt16(data, 0, 3); WriteUInt16(data, 2, 1); WriteUInt16(data, 4, 1); WriteUInt16(data, 6, 12);
        WriteUInt16(data, 10, 1); WriteUInt16(data, 12, 1); WriteUInt16(data, 14, (ushort)coverage.Length);
        for (int index = 0; index < coverage.Length; index++) WriteUInt16(data, 16 + index * 2, coverage[index]);
        return data;
    }
    internal static byte[] CreateFontWithRequiredLigature(string featureTag, int first = 'f', int second = 'i', ushort flags = 0) {
        byte[] gsub = CreateLigatureGsub(featureTag, 1, 2, 3, "latn", flags);
        WriteUInt16(gsub, 76, 0);
        return CreateFontFromCmap(CreateFormat12Cmap(first, 1, second, 2, 32, 4), glyphCount: 5, gsub: gsub);
    }

    internal static byte[] CreateFontWithOversizedLookupList() {
        byte[] original = CreateLigatureGsub("liga", 1, 2, 3, "latn", 0);
        const int lookup = 8222;
        var gsub = new byte[8274];
        Array.Copy(original, 0, gsub, 0, 26);
        WriteUInt16(gsub, 4, 8254);
        WriteUInt16(gsub, 26, 4097);
        WriteUInt16(gsub, 28, 8196);
        Array.Copy(original, 30, gsub, lookup, 32);
        Array.Copy(original, 62, gsub, 8254, 20);
        return CreateFontFromCmap(CreateFormat12Cmap('f', 1, 'i', 2, 32, 4), glyphCount: 5, gsub: gsub);
    }

    // A contextual rule expands the first input, then substitutes the second input.
    // The latter record must never target the newly inserted continuation glyph.
    internal static byte[] CreateFontWithContextualExpansion(bool expansionLast) {
        var gsub = new byte[130];
        WriteUInt32(gsub, 0, 0x00010000);
        WriteUInt16(gsub, 4, 10); WriteUInt16(gsub, 6, 12); WriteUInt16(gsub, 8, 26);
        WriteUInt16(gsub, 12, 1); WriteTag(gsub, 14, "calt"); WriteUInt16(gsub, 18, 8);
        WriteUInt16(gsub, 22, 1);
        WriteUInt16(gsub, 26, 3); WriteUInt16(gsub, 28, 8); WriteUInt16(gsub, 30, 46); WriteUInt16(gsub, 32, 76);
        WriteUInt16(gsub, 34, 5); WriteUInt16(gsub, 38, 1); WriteUInt16(gsub, 40, 8);
        WriteUInt16(gsub, 42, 3); WriteUInt16(gsub, 44, 2); WriteUInt16(gsub, 46, 2);
        WriteUInt16(gsub, 48, 18); WriteUInt16(gsub, 50, 24);
        WriteUInt16(gsub, 52, expansionLast ? (ushort)1 : (ushort)0);
        WriteUInt16(gsub, 54, expansionLast ? (ushort)2 : (ushort)1);
        WriteUInt16(gsub, 56, expansionLast ? (ushort)0 : (ushort)1);
        WriteUInt16(gsub, 58, expansionLast ? (ushort)1 : (ushort)2);
        WriteUInt16(gsub, 60, 1); WriteUInt16(gsub, 62, 1); WriteUInt16(gsub, 64, 1);
        WriteUInt16(gsub, 66, 1); WriteUInt16(gsub, 68, 1); WriteUInt16(gsub, 70, 2);
        Array.Copy(CreateMultipleGsub(), 30, gsub, 72, 30);
        Array.Copy(CreateContextualGsub(), 66, gsub, 102, 22);
        return CreateFontFromCmap(CreateFormat12Cmap('A', 1, 'B', 2), glyphCount: 5, gsub: gsub);
    }

}
