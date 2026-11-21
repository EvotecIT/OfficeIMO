namespace OfficeIMO.TestAssets;

internal static partial class ManagedTextShapingTestAssets {
    /// <summary>Emoji sequences the <see cref="CreateEmojiSequenceFont"/> fixture joins through one ccmp ligature lookup.</summary>
    internal static readonly (string Text, int Ligature)[] EmojiSequenceLigatures = {
        ("\U0001F3CB\u200D\u2642", EmojiFirstLigatureGlyph),
        ("\U0001F44D\U0001F3FB", EmojiFirstLigatureGlyph + 1),
        ("#\u20E3", EmojiFirstLigatureGlyph + 2),
        ("\U0001F1F5\U0001F1F1", EmojiFirstLigatureGlyph + 3),
        ("\U0001F3F4\U000E0067\U000E0062\U000E0065\U000E006E\U000E0067\U000E007F", EmojiFirstLigatureGlyph + 4),
        ("\U0001F468\u200D\U0001F469", EmojiFirstLigatureGlyph + 5)
    };

    /// <summary>Advance, in 1000-unit em, of every emoji and ligature glyph in the fixture.</summary>
    internal const int EmojiAdvance = 1000;

    private static readonly int[] EmojiFixtureScalars = {
        0x20, '#', 'A', 'f', 'i', 0x200C, 0x200D, 0x20E3, 0x2642, 0x2764, 0xFE0E, 0xFE0F,
        0x1F1F1, 0x1F1F5, 0x1F3CB, 0x1F3F4, 0x1F3FB, 0x1F44D, 0x1F468, 0x1F469,
        0xE0062, 0xE0065, 0xE0067, 0xE006E, 0xE007F
    };

    private const int EmojiFirstLigatureGlyph = 26;

    /// <summary>
    /// A small outline emoji font in the shape of common color emoji fonts: a DFLT/latn ccmp ligature lookup
    /// lists ZWJ and tag characters as components but omits variation selectors. Its ZWJ, ZWNJ, tag and
    /// variation-selector glyphs are inked with nonzero advances, so painting any of them is observable.
    /// The space glyph is blank.
    /// </summary>
    /// <param name="bitmapOnlyGlyphs">Glyphs given an empty outline in a font that also carries a CBDT table, as bitmap emoji fonts do.</param>
    /// <param name="redFirstLigature">Gives the first ligature a single red COLR layer while its components stay plain, as in color emoji fonts.</param>
    /// <param name="additionalTables">Further tables, such as an fvar that makes the font variable.</param>
    internal static byte[] CreateEmojiSequenceFont(IReadOnlyCollection<int>? bitmapOnlyGlyphs = null, bool redFirstLigature = false,
        IReadOnlyList<(string Tag, byte[] Data)>? additionalTables = null) {
        int glyphCount = EmojiFirstLigatureGlyph + EmojiSequenceLigatures.Length;
        var advances = new Dictionary<int, ushort>();
        for (int glyph = 1; glyph < glyphCount; glyph++) advances[glyph] = EmojiAdvance;
        advances[EmojiGlyph(0x20)] = 250;
        foreach (int letter in new[] { (int)'#', 'A', 'f', 'i' }) advances[EmojiGlyph(letter)] = 500;
        foreach (int control in new[] { 0x200C, 0x200D, 0xFE0E, 0xFE0F, 0xE0062, 0xE0065, 0xE0067, 0xE006E, 0xE007F })
            advances[EmojiGlyph(control)] = 600;
        var ligatures = new List<(ushort[] Components, ushort Ligature)>();
        foreach ((string text, int ligature) in EmojiSequenceLigatures) {
            var components = new List<ushort>();
            for (int index = 0; index < text.Length; index++) {
                int scalar = char.ConvertToUtf32(text, index);
                if (scalar > 0xFFFF) index++;
                components.Add(checked((ushort)EmojiGlyph(scalar)));
            }
            ligatures.Add((components.ToArray(), checked((ushort)ligature)));
        }
        var blanks = new List<int> { EmojiGlyph(0x20) };
        if (bitmapOnlyGlyphs != null) blanks.AddRange(bitmapOnlyGlyphs);
        return CreateFontFromCmap(CreateDistinctFormat12Cmap(EmojiFixtureScalars), glyphCount: glyphCount,
            gsub: CreateLigatureSetGsub("ccmp", ligatures), glyphAdvances: advances, blankGlyphIds: blanks,
            colr: redFirstLigature ? CreateSingleLayerColr(EmojiFirstLigatureGlyph, paletteEntry: 0) : null,
            cpal: redFirstLigature ? CreateCpalV1() : null,
            additionalTables: bitmapOnlyGlyphs == null ? additionalTables
                : new[] { ("CBDT", new byte[8]) }.Concat(additionalTables ?? Array.Empty<(string, byte[])>()).ToArray());
    }

    // COLR version 0 with one base glyph painted by its own outline in one palette entry.
    private static byte[] CreateSingleLayerColr(int glyph, int paletteEntry) {
        var table = new byte[24];
        WriteUInt16(table, 2, 1);
        WriteUInt32(table, 4, 14);
        WriteUInt32(table, 8, 20);
        WriteUInt16(table, 12, 1);
        WriteUInt16(table, 14, checked((ushort)glyph));
        WriteUInt16(table, 18, 1);
        WriteUInt16(table, 20, checked((ushort)glyph));
        WriteUInt16(table, 22, checked((ushort)paletteEntry));
        return table;
    }

    /// <summary>The fixture glyph a scalar maps to.</summary>
    internal static int EmojiGlyph(int scalar) => Array.IndexOf(EmojiFixtureScalars, scalar) + 1;

    /// <summary>One feature with one LigatureSubstFormat1 lookup, registered for DFLT and latn.</summary>
    private static byte[] CreateLigatureSetGsub(string featureTag, IReadOnlyList<(ushort[] Components, ushort Ligature)> ligatures) {
        var sets = new SortedDictionary<ushort, List<(ushort[] Components, ushort Ligature)>>();
        foreach (var ligature in ligatures) {
            if (!sets.TryGetValue(ligature.Components[0], out var set)) sets.Add(ligature.Components[0], set = new());
            set.Add(ligature);
        }
        // LigatureSubstFormat1 subtable: header, coverage, then each LigatureSet and its Ligatures.
        var subtable = new List<byte>();
        void U16(List<byte> target, int value) { target.Add((byte)(value >> 8)); target.Add((byte)value); }
        void Patch(List<byte> target, int at, int value) { target[at] = (byte)(value >> 8); target[at + 1] = (byte)value; }
        U16(subtable, 1);
        int coverageField = subtable.Count; U16(subtable, 0);
        U16(subtable, sets.Count);
        int setFields = subtable.Count;
        foreach (var _ in sets) U16(subtable, 0);
        Patch(subtable, coverageField, subtable.Count);
        U16(subtable, 1); U16(subtable, sets.Count);
        foreach (ushort first in sets.Keys) U16(subtable, first);
        int setIndex = 0;
        foreach (var set in sets.Values) {
            int setStart = subtable.Count;
            Patch(subtable, setFields + setIndex++ * 2, setStart);
            U16(subtable, set.Count);
            int ligatureFields = subtable.Count;
            foreach (var _ in set) U16(subtable, 0);
            for (int index = 0; index < set.Count; index++) {
                Patch(subtable, ligatureFields + index * 2, subtable.Count - setStart);
                U16(subtable, set[index].Ligature);
                U16(subtable, set[index].Components.Length);
                for (int component = 1; component < set[index].Components.Length; component++) U16(subtable, set[index].Components[component]);
            }
        }

        var data = new List<byte>();
        U16(data, 1); U16(data, 0); // version 1.0
        U16(data, 0); U16(data, 0); U16(data, 0); // ScriptList, FeatureList, LookupList offsets
        Patch(data, 4, data.Count);
        U16(data, 2);
        foreach (string script in new[] { "DFLT", "latn" }) {
            foreach (char character in script) data.Add((byte)character);
            U16(data, 2 + 2 * 6 + (script == "DFLT" ? 0 : 12));
        }
        for (int script = 0; script < 2; script++) {
            U16(data, 4); U16(data, 0); // Script: DefaultLangSys offset, LangSysCount
            U16(data, 0); U16(data, 0xFFFF); U16(data, 1); U16(data, 0); // LangSys: one feature
        }
        Patch(data, 6, data.Count);
        U16(data, 1);
        foreach (char character in featureTag) data.Add((byte)character);
        U16(data, 8);
        U16(data, 0); U16(data, 1); U16(data, 0); // Feature: params, one lookup, index 0
        Patch(data, 8, data.Count);
        U16(data, 1); U16(data, 4); // LookupList: one lookup at offset 4
        U16(data, 4); U16(data, 0); U16(data, 1); U16(data, 8); // Lookup: type 4, flags, one subtable at 8
        data.AddRange(subtable);
        return data.ToArray();
    }
}
