namespace OfficeIMO.TestAssets;

internal static partial class ManagedTextShapingTestAssets {
    /// <summary>A proven blank control can expand into independently blank or inked GSUB output glyphs.</summary>
    internal static byte[] CreateControlSubstitutionFont(bool blankFirstReplacement, bool blankSecondReplacement) {
        var blanks = new List<int> { 1 };
        if (blankFirstReplacement) blanks.Add(3);
        if (blankSecondReplacement) blanks.Add(4);
        return CreateFontFromCmap(CreateDistinctFormat12Cmap(new[] { 0x20, (int)'A' }), glyphCount: 5,
            gsub: CreateMultipleGsub(scriptTag: "latn", featureTag: "ccmp"), blankGlyphIds: blanks);
    }

    /// <summary>A font whose actual space/control contours independently determine safe nonpainting admission.</summary>
    /// <param name="variable">Adds a one-axis <c>fvar</c> table, so the font takes the nominal variable-font route.</param>
    /// <param name="distanceAdvance">Adds a GPOS <c>dist</c> single adjustment of this many units to the advance of <c>A</c>.</param>
    internal static byte[] CreateControlAdmissionFont(int control = 0x200C, bool blankSpace = false,
        bool blankControl = false, bool withSubstitution = false, bool withKerning = false,
        bool variable = false, short? distanceAdvance = null) {
        int[] scalars = new[] { 0x20, (int)'A', 'V', 'a', 'b', 'c', 0x0628, 0xFE8F, 0xFE90, 0xFE91, 0xFE92,
            control, 0x202C, 0x202E }.Distinct().OrderBy(value => value).ToArray();
        int Glyph(int scalar) => Array.IndexOf(scalars, scalar) + 1;
        byte[]? gsub = null;
        if (withSubstitution) {
            byte[] basis = CreateAlternatingScriptBudgetFont(1);
            int tables = ReadUInt16(basis, 4);
            for (int index = 0; index < tables; index++) {
                int record = 12 + index * 16;
                if (ReadUInt32(basis, record) != 0x47535542) continue;
                gsub = new byte[checked((int)ReadUInt32(basis, record + 12))];
                Array.Copy(basis, checked((int)ReadUInt32(basis, record + 8)), gsub, 0, gsub.Length);
                int lookups = ReadUInt16(gsub, 8), lookup = lookups + ReadUInt16(gsub, lookups + 2);
                int subtable = lookup + ReadUInt16(gsub, lookup + 6), coverage = subtable + ReadUInt16(gsub, subtable + 2);
                WriteUInt16(gsub, subtable + 6, checked((ushort)Glyph(0x0628)));
                WriteUInt16(gsub, coverage + 4, checked((ushort)Glyph(0x0628)));
                break;
            }
            if (gsub == null) throw new InvalidOperationException("The fixture GSUB table is unavailable.");
        }
        var blanks = new List<int>();
        if (blankSpace) blanks.Add(Glyph(0x20));
        if (blankControl) blanks.Add(Glyph(control));
        byte[]? kern = withKerning ? CreateKernTable(checked((ushort)Glyph('A')), checked((ushort)Glyph('V')), -120) : null;
        byte[]? gpos = distanceAdvance.HasValue ? CreateSingleAdvanceGpos("dist", checked((ushort)Glyph('A')), distanceAdvance.Value) : null;
        return CreateFontFromCmap(CreateDistinctFormat12Cmap(scalars), glyphCount: scalars.Length + 1,
            gsub: gsub, kern: kern, gpos: gpos, blankGlyphIds: blanks,
            additionalTables: variable ? new[] { ("fvar", CreateWeightAxisFvar()) } : null);
    }

    // DFLT script, one feature, one SinglePosFormat1 lookup adjusting XAdvance of one glyph.
    private static byte[] CreateSingleAdvanceGpos(string featureTag, ushort glyph, short advance) {
        var data = new byte[70];
        WriteUInt32(data, 0, 0x00010000);
        WriteUInt16(data, 4, 10); WriteUInt16(data, 6, 30); WriteUInt16(data, 8, 44);
        WriteUInt16(data, 10, 1); WriteTag(data, 12, "DFLT"); WriteUInt16(data, 16, 8);
        WriteUInt16(data, 18, 4); WriteUInt16(data, 20, 0);
        WriteUInt16(data, 22, 0); WriteUInt16(data, 24, 0xFFFF); WriteUInt16(data, 26, 1); WriteUInt16(data, 28, 0);
        WriteUInt16(data, 30, 1); WriteTag(data, 32, featureTag); WriteUInt16(data, 36, 8);
        WriteUInt16(data, 38, 0); WriteUInt16(data, 40, 1); WriteUInt16(data, 42, 0);
        WriteUInt16(data, 44, 1); WriteUInt16(data, 46, 4);
        WriteUInt16(data, 48, 1); WriteUInt16(data, 50, 0); WriteUInt16(data, 52, 1); WriteUInt16(data, 54, 8);
        WriteUInt16(data, 56, 1); WriteUInt16(data, 58, 8); WriteUInt16(data, 60, 4); WriteUInt16(data, 62, unchecked((ushort)advance));
        WriteUInt16(data, 64, 1); WriteUInt16(data, 66, 1); WriteUInt16(data, 68, glyph);
        return data;
    }

    /// <summary>The emoji sequence fixture with a weight axis: variable, with GSUB and no GPOS or feature variations.</summary>
    internal static byte[] CreateVariableEmojiSequenceFont() =>
        CreateEmojiSequenceFont(additionalTables: new[] { ("fvar", CreateWeightAxisFvar()) });

    private static byte[] CreateWeightAxisFvar() {
        var fvar = new byte[36];
        WriteUInt16(fvar, 0, 1);
        WriteUInt16(fvar, 4, 16);
        WriteUInt16(fvar, 6, 2);
        WriteUInt16(fvar, 8, 1);
        WriteUInt16(fvar, 10, 20);
        WriteUInt16(fvar, 14, 8);
        WriteTag(fvar, 16, "wght");
        WriteUInt32(fvar, 20, 100U << 16);
        WriteUInt32(fvar, 24, 400U << 16);
        WriteUInt32(fvar, 28, 900U << 16);
        WriteUInt16(fvar, 34, 256);
        return fvar;
    }
}
