namespace OfficeIMO.TestAssets;

internal static partial class ManagedTextShapingTestAssets {
    internal const int CursiveAcute = 0x0301;
    internal const int CursiveHiriq = 0x05B4;

    /// <summary>
    /// Cursive attachment fixture: Latin a, b, c and Hebrew alef, bet, gimel chain through entry and
    /// exit anchors at different heights; an acute and a hiriq attach to each letter afterwards.
    /// </summary>
    internal static byte[] CreateCursiveAttachmentFont(bool rightToLeftFlag) {
        int[] scalars = { 'a', 'b', 'c', CursiveAcute, CursiveHiriq, 0x05D0, 0x05D1, 0x05D2 };
        // Glyph IDs follow scalar order: a=1, b=2, c=3, acute=4, hiriq=5, alef=6, bet=7, gimel=8.
        int[] letters = { 1, 2, 3, 6, 7, 8 }, marks = { 4, 5 };
        var cursive = new LayoutBuilder();
        cursive.U16(1); int coverage = cursive.Reserve(); cursive.U16(letters.Length);
        var anchors = new List<(int Offset, int X, int Y)>();
        for (int index = 0; index < letters.Length; index++) {
            int position = index % 3;
            int entry = cursive.Reserve(), exit = cursive.Reserve();
            if (position > 0) anchors.Add((entry, 40 + position * 10, 60 * position));
            if (position < 2) anchors.Add((exit, 520 - position * 30, 150 + 90 * position));
        }
        cursive.Offset(coverage, 0); cursive.Coverage(letters);
        foreach ((int offset, int x, int y) in anchors) { cursive.Offset(offset, 0); cursive.Anchor(x, y); }
        byte[] gpos = Layout(new[] { ("curs", new[] { 0 }), ("mark", new[] { 1 }) }, new[] {
            Lookup(3, (ushort)(0x0008 | (rightToLeftFlag ? 0x0001 : 0)), cursive.ToArray()),
            Lookup(4, 0, MarkToBase(marks, letters))
        });
        return CreateFontFromCmap(CreateDistinctFormat12Cmap(scalars), glyphCount: 9, gpos: gpos,
            additionalTables: new[] { ("GDEF", GlyphClasses(9, marks, markAttachmentClasses: false)) },
            glyphAdvances: new Dictionary<int, ushort> { [4] = 0, [5] = 0, [1] = 560, [2] = 530, [3] = 500, [6] = 560, [7] = 530, [8] = 500 },
            strictHeaders: true);
    }

    /// <summary>
    /// Context fixture over a, b, c, x, y and a combining acute: class- and glyph-based GSUB rules
    /// whose nested lookups carry different flags, and class- and glyph-based GPOS rules whose
    /// mark-attachment-type flag has no MarkAttachClassDef to read.
    /// </summary>
    internal static byte[] CreateContextFormatsFont() {
        int[] scalars = { 'a', 'b', 'c', 'x', 'y', CursiveAcute };
        // a=1, b=2, c=3, x=4, y=5, acute=6.
        int[] bases = { 1, 2, 3, 4, 5 }, marks = { 6 };

        // GSUB 6 format 2, IgnoreMarks: b between a and c becomes x through a nested lookup without flags.
        var chain = new LayoutBuilder();
        chain.U16(2); int chainCoverage = chain.Reserve(), backtrackClasses = chain.Reserve(), inputClasses = chain.Reserve(),
            lookaheadClasses = chain.Reserve();
        chain.U16(3); chain.U16(0); chain.U16(0); int classSet = chain.Reserve();
        chain.Offset(chainCoverage, 0); chain.Coverage(new[] { 2 });
        int classDef = chain.Position; chain.U16(1); chain.U16(1); chain.U16(3); chain.U16(1); chain.U16(2); chain.U16(3);
        chain.OffsetTo(backtrackClasses, 0, classDef); chain.OffsetTo(inputClasses, 0, classDef); chain.OffsetTo(lookaheadClasses, 0, classDef);
        chain.Offset(classSet, 0); int set = chain.Position; chain.U16(1); int rule = chain.Reserve();
        chain.Offset(rule, set); chain.U16(1); chain.U16(1); chain.U16(1); chain.U16(1); chain.U16(3); chain.U16(1); chain.U16(0); chain.U16(1);

        // GSUB 5 format 1: c followed by a turns the a into y through a nested lookup using a mark filtering set.
        var context = new LayoutBuilder();
        context.U16(1); int contextCoverage = context.Reserve(); context.U16(1); int ruleSet = context.Reserve();
        context.Offset(contextCoverage, 0); context.Coverage(new[] { 3 });
        context.Offset(ruleSet, 0); int rules = context.Position; context.U16(1); int glyphRule = context.Reserve();
        context.Offset(glyphRule, rules); context.U16(2); context.U16(1); context.U16(1); context.U16(1); context.U16(3);

        // GPOS 7 format 2 with a mark-attachment-type flag: a followed by x widens the x.
        var positionContext = new LayoutBuilder();
        positionContext.U16(2); int positionCoverage = positionContext.Reserve(), positionClasses = positionContext.Reserve();
        positionContext.U16(2); positionContext.U16(0); int positionSet = positionContext.Reserve();
        positionContext.Offset(positionCoverage, 0); positionContext.Coverage(new[] { 1 });
        positionContext.Offset(positionClasses, 0); positionContext.U16(1); positionContext.U16(1); positionContext.U16(4);
        positionContext.U16(1); positionContext.U16(0); positionContext.U16(0); positionContext.U16(2);
        positionContext.Offset(positionSet, 0); int positionRules = positionContext.Position; positionContext.U16(1);
        int positionRule = positionContext.Reserve();
        positionContext.Offset(positionRule, positionRules);
        positionContext.U16(2); positionContext.U16(1); positionContext.U16(2); positionContext.U16(1); positionContext.U16(1);

        // GPOS 8 format 1: c followed by y moves the c.
        var positionChain = new LayoutBuilder();
        positionChain.U16(1); int chainPositionCoverage = positionChain.Reserve(); positionChain.U16(1); int chainRuleSet = positionChain.Reserve();
        positionChain.Offset(chainPositionCoverage, 0); positionChain.Coverage(new[] { 3 });
        positionChain.Offset(chainRuleSet, 0); int chainRules = positionChain.Position; positionChain.U16(1); int chainRule = positionChain.Reserve();
        positionChain.Offset(chainRule, chainRules);
        positionChain.U16(0); positionChain.U16(1); positionChain.U16(1); positionChain.U16(5); positionChain.U16(1); positionChain.U16(0); positionChain.U16(3);

        byte[] gsub = Layout(new[] { ("calt", new[] { 0, 2 }) }, new[] {
            Lookup(6, 0x0008, chain.ToArray()),
            Lookup(1, 0, SingleSubstitution(new[] { 2 }, new[] { 4 })),
            Lookup(5, 0, context.ToArray()),
            Lookup(1, 0x0010, SingleSubstitution(new[] { 1 }, new[] { 5 }), markFilteringSet: 0)
        });
        byte[] gpos = Layout(new[] { ("kern", new[] { 0, 2 }), ("mark", new[] { 4 }) }, new[] {
            Lookup(7, 0x0100, positionContext.ToArray()),
            Lookup(1, 0, SinglePositioning(new[] { 4 }, xPlacement: 0, yPlacement: 0, xAdvance: 120)),
            Lookup(8, 0x0008, positionChain.ToArray()),
            Lookup(1, 0, SinglePositioning(new[] { 3 }, xPlacement: -40, yPlacement: 25, xAdvance: -30)),
            Lookup(4, 0, MarkToBase(marks, bases))
        });
        return CreateFontFromCmap(CreateDistinctFormat12Cmap(scalars), glyphCount: 7, gsub: gsub, gpos: gpos,
            additionalTables: new[] { ("GDEF", GlyphClasses(7, marks, markAttachmentClasses: false, markSet: new[] { 6 })) },
            glyphAdvances: new Dictionary<int, ushort> { [6] = 0, [1] = 540, [2] = 520, [3] = 500, [4] = 510, [5] = 490 },
            strictHeaders: true);
    }

    private static byte[] Layout((string Tag, int[] Lookups)[] features, byte[][] lookups) {
        var data = new LayoutBuilder();
        data.U16(1); data.U16(0); int scripts = data.Reserve(), featureOffset = data.Reserve(), lookupOffset = data.Reserve();
        // Every script record shares one default language system.
        string[] scriptTags = { "DFLT", "hebr", "latn" };
        data.Offset(scripts, 0); int scriptList = data.Position; data.U16(scriptTags.Length);
        int[] scriptRecords = scriptTags.Select(tag => { data.Tag(tag); return data.Reserve(); }).ToArray();
        int scriptTable = data.Position;
        foreach (int record in scriptRecords) data.OffsetTo(record, scriptList, scriptTable);
        int defaultLanguage = data.Reserve(); data.U16(0);
        data.Offset(defaultLanguage, scriptTable); data.U16(0); data.U16(ushort.MaxValue); data.U16(features.Length);
        for (int index = 0; index < features.Length; index++) data.U16(index);
        data.Offset(featureOffset, 0); int featureList = data.Position; data.U16(features.Length);
        var featureTables = new int[features.Length];
        for (int index = 0; index < features.Length; index++) { data.Tag(features[index].Tag); featureTables[index] = data.Reserve(); }
        for (int index = 0; index < features.Length; index++) {
            data.Offset(featureTables[index], featureList); data.U16(0); data.U16(features[index].Lookups.Length);
            foreach (int lookup in features[index].Lookups) data.U16(lookup);
        }
        data.Offset(lookupOffset, 0); int lookupList = data.Position; data.U16(lookups.Length);
        int[] lookupTables = lookups.Select(_ => data.Reserve()).ToArray();
        for (int index = 0; index < lookups.Length; index++) { data.Offset(lookupTables[index], lookupList); data.Bytes(lookups[index]); }
        return data.ToArray();
    }

    private static byte[] Lookup(int type, ushort flags, byte[] subtable, int markFilteringSet = -1) {
        var data = new LayoutBuilder(); data.U16(type); data.U16(flags); data.U16(1);
        int offset = data.Reserve();
        if ((flags & 0x0010) != 0) data.U16(markFilteringSet);
        data.Offset(offset, 0); data.Bytes(subtable); return data.ToArray();
    }

    private static byte[] SingleSubstitution(int[] glyphs, int[] replacements) {
        var data = new LayoutBuilder(); data.U16(2); int coverage = data.Reserve(); data.U16(replacements.Length);
        foreach (int replacement in replacements) data.U16(replacement);
        data.Offset(coverage, 0); data.Coverage(glyphs); return data.ToArray();
    }

    private static byte[] SinglePositioning(int[] glyphs, int xPlacement, int yPlacement, int xAdvance) {
        var data = new LayoutBuilder(); data.U16(1); int coverage = data.Reserve(); data.U16(0x0007);
        data.U16(xPlacement); data.U16(yPlacement); data.U16(xAdvance);
        data.Offset(coverage, 0); data.Coverage(glyphs); return data.ToArray();
    }

    private static byte[] MarkToBase(int[] marks, int[] bases) {
        var data = new LayoutBuilder(); data.U16(1);
        int markCoverage = data.Reserve(), baseCoverage = data.Reserve(); data.U16(1);
        int markArray = data.Reserve(), baseArray = data.Reserve();
        data.Offset(markCoverage, 0); data.Coverage(marks);
        data.Offset(baseCoverage, 0); data.Coverage(bases);
        data.Offset(markArray, 0); int markTable = data.Position; data.U16(marks.Length);
        int[] markAnchors = marks.Select(_ => { data.U16(0); return data.Reserve(); }).ToArray();
        foreach (int anchor in markAnchors) { data.Offset(anchor, markTable); data.Anchor(-20, 600); }
        data.Offset(baseArray, 0); int baseTable = data.Position; data.U16(bases.Length);
        int[] baseAnchors = bases.Select(_ => data.Reserve()).ToArray();
        for (int index = 0; index < baseAnchors.Length; index++) { data.Offset(baseAnchors[index], baseTable); data.Anchor(230 + index * 10, 710); }
        return data.ToArray();
    }

    private static byte[] GlyphClasses(int glyphCount, int[] marks, bool markAttachmentClasses, int[]? markSet = null) {
        var data = new LayoutBuilder(); data.U16(1); data.U16(markSet == null ? 0 : 2);
        int classes = data.Reserve(); data.U16(0); data.U16(0); int attachment = data.Reserve();
        int sets = markSet == null ? -1 : data.Reserve();
        data.Offset(classes, 0); data.U16(1); data.U16(1); data.U16(glyphCount - 1);
        for (int glyph = 1; glyph < glyphCount; glyph++) data.U16(Array.IndexOf(marks, glyph) >= 0 ? 3 : 1);
        if (markAttachmentClasses) {
            data.Offset(attachment, 0); data.U16(1); data.U16(1); data.U16(glyphCount - 1);
            for (int glyph = 1; glyph < glyphCount; glyph++) data.U16(Array.IndexOf(marks, glyph) >= 0 ? 1 : 0);
        }
        if (markSet != null) { data.Offset(sets, 0); data.U16(1); data.U16(1); data.U32(8); data.Coverage(markSet); }
        return data.ToArray();
    }

    /// <summary>Big-endian OpenType layout writer with 16-bit offsets patched after their targets.</summary>
    private sealed class LayoutBuilder {
        private readonly List<byte> _bytes = new();
        internal int Position => _bytes.Count;
        internal void U16(int value) { _bytes.Add((byte)(value >> 8)); _bytes.Add((byte)value); }
        internal void U32(int value) { U16(value >> 16); U16(value); }
        internal int Reserve() { int offset = Position; U16(0); return offset; }
        internal void Offset(int offset, int origin) => OffsetTo(offset, origin, Position);
        internal void OffsetTo(int offset, int origin, int target) {
            int value = checked(target - origin);
            _bytes[offset] = checked((byte)(value >> 8)); _bytes[offset + 1] = (byte)value;
        }
        internal void Tag(string tag) { foreach (char character in tag) _bytes.Add((byte)character); }
        internal void Bytes(byte[] bytes) => _bytes.AddRange(bytes);
        internal void Coverage(int[] glyphs) { U16(1); U16(glyphs.Length); foreach (int glyph in glyphs) U16(glyph); }
        internal void Anchor(int x, int y) { U16(1); U16(x); U16(y); }
        internal byte[] ToArray() => _bytes.ToArray();
    }
}
