using OfficeIMO.Drawing;

namespace OfficeIMO.TestAssets;

internal static partial class ManagedTextShapingTestAssets {
    /// <summary>Units per em of <see cref="CreateBitmapColorFont"/>; its PNG strike is drawn at <see cref="BitmapColorFontPpem"/>.</summary>
    internal const int BitmapColorFontPpem = 100;

    /// <summary>The origin of the outline that <see cref="SbixVariant.OutlineOrigin"/> gives its bitmap glyph.</summary>
    internal static readonly (int Left, int Bottom) SbixOutlineOrigin = (100, 50);

    /// <summary>sbix layouts that <see cref="CreateSbixVariantFont"/> builds around the red PNG of glyph 1.</summary>
    internal enum SbixVariant {
        /// <summary>Glyph 2, mapped to U+1F601, is a <c>dupe</c> record of glyph 1.</summary>
        Duplicate,
        /// <summary>The largest of two strikes has no record for glyph 1; the smaller one has its PNG.</summary>
        MissingFromLargestStrike,
        /// <summary>Glyph 1 has an outline whose bounding box starts at <see cref="SbixOutlineOrigin"/>.</summary>
        OutlineOrigin,
        /// <summary>U+1F600 to U+1F602 map to glyphs 1 to 3: glyph 3 is a <c>dupe</c> of glyph 2, which is a <c>dupe</c> of glyph 1.</summary>
        DuplicateChain,
        /// <summary>
        /// A 10 ppem strike has glyph 1's PNG only (400 units); a 5 ppem strike has its own glyph 1 PNG (800 units)
        /// and glyph 2, mapped to U+1F601, as a <c>dupe</c> of glyph 1.
        /// </summary>
        DuplicateInSmallerStrike,
        /// <summary>
        /// Glyph 2, mapped to U+1F601, is a <c>dupe</c> of glyph 1; glyph 1's outline starts at <see cref="SbixOutlineOrigin"/>
        /// and glyph 2's at <see cref="SbixDuplicateOutlineOrigin"/>.
        /// </summary>
        DuplicateWithOwnOutline
    }

    /// <summary>The outline origin of the duplicate glyph in <see cref="SbixVariant.DuplicateWithOwnOutline"/>.</summary>
    internal static readonly (int Left, int Bottom) SbixDuplicateOutlineOrigin = (300, 20);

    /// <summary>
    /// One glyph whose only artwork is a solid red PNG, stored the way Noto Color Emoji (CBDT/CBLC, image format 17)
    /// or Apple Color Emoji (sbix, <c>png </c>) store it. The glyph outline is empty. The bitmap is
    /// <paramref name="pixels"/> square at <see cref="BitmapColorFontPpem"/> ppem, sitting on the baseline.
    /// <paramref name="sameGlyphScalars"/> map to that glyph too. With <paramref name="overlappingLetter"/>, that letter
    /// maps to a plain 400-unit box drawn 500 units after the emoji's origin, so it overlaps the 800-unit PNG.
    /// </summary>
    internal static byte[] CreateBitmapColorFont(int scalar, bool sbix, int pixels = 80,
        IReadOnlyList<int>? sameGlyphScalars = null, int? overlappingLetter = null) {
        byte[] png = CreateRedPng();
        if (overlappingLetter is int letter) {
            // Distinct cmap entries in scalar order: the letter is glyph 1 and the emoji glyph 2.
            int bitmapGlyph = letter < scalar ? 2 : 1;
            int[] ordered = letter < scalar ? new[] { letter, scalar } : new[] { scalar, letter };
            var tables = sbix
                ? new[] { ("sbix", PadToFourBytes(CreateSbixStrikes(3, (BitmapColorFontPpem * 4 / pixels, new Dictionary<int, (short, short, string, byte[])> { [bitmapGlyph] = (0, 0, "png ", png) })))) }
                : new[] { ("CBDT", PadToFourBytes(CreateCbdt(png, pixels))), ("CBLC", CreateCblc(png.Length, bitmapGlyph)) };
            return CreateFontFromCmap(CreateDistinctFormat12Cmap(ordered), includeTrailingMetric: true, glyphCount: 3,
                blankGlyphIds: new[] { bitmapGlyph }, additionalTables: tables);
        }
        var bitmapTables = sbix
            ? new[] { ("sbix", PadToFourBytes(CreateSbixStrikes(2, (BitmapColorFontPpem * 4 / pixels, new Dictionary<int, (short, short, string, byte[])> { [1] = (0, 0, "png ", png) })))) }
            : new[] { ("CBDT", PadToFourBytes(CreateCbdt(png, pixels))), ("CBLC", CreateCblc(png.Length, 1)) };
        return CreateFontFromCmap(CreateFormat12Cmap(new[] { scalar }.Concat(sameGlyphScalars ?? Array.Empty<int>()).ToArray()),
            includeTrailingMetric: true, glyphCount: 2, blankGlyphIds: new[] { 1 }, additionalTables: bitmapTables);
    }

    /// <summary>An sbix font for U+1F600 (glyph 1, an 80-pixel red PNG at 100 ppem) laid out as <paramref name="variant"/> describes.</summary>
    internal static byte[] CreateSbixVariantFont(SbixVariant variant) {
        byte[] png = CreateRedPng();
        const int ppem = BitmapColorFontPpem * 4 / 80;
        var glyph1 = new Dictionary<int, (short, short, string, byte[])> { [1] = (0, 0, "png ", png) };
        switch (variant) {
            case SbixVariant.Duplicate: {
                byte[] target = new byte[2];
                WriteUInt16(target, 0, 1);
                var glyphs = new Dictionary<int, (short, short, string, byte[])> { [1] = (0, 0, "png ", png), [2] = (0, 0, "dupe", target) };
                return CreateFontFromCmap(CreateDistinctFormat12Cmap(new[] { 0x1F600, 0x1F601 }), includeTrailingMetric: true, glyphCount: 3,
                    blankGlyphIds: new[] { 1, 2 }, additionalTables: new[] { ("sbix", PadToFourBytes(CreateSbixStrikes(3, (ppem, glyphs)))) });
            }
            case SbixVariant.DuplicateChain: {
                byte[] toFirst = new byte[2], toSecond = new byte[2];
                WriteUInt16(toFirst, 0, 1);
                WriteUInt16(toSecond, 0, 2);
                var glyphs = new Dictionary<int, (short, short, string, byte[])> {
                    [1] = (0, 0, "png ", png), [2] = (0, 0, "dupe", toFirst), [3] = (0, 0, "dupe", toSecond)
                };
                return CreateFontFromCmap(CreateDistinctFormat12Cmap(new[] { 0x1F600, 0x1F601, 0x1F602 }), includeTrailingMetric: true, glyphCount: 4,
                    blankGlyphIds: new[] { 1, 2, 3 }, additionalTables: new[] { ("sbix", PadToFourBytes(CreateSbixStrikes(4, (ppem, glyphs)))) });
            }
            case SbixVariant.DuplicateInSmallerStrike: {
                byte[] toFirst = new byte[2];
                WriteUInt16(toFirst, 0, 1);
                var small = new Dictionary<int, (short, short, string, byte[])> { [1] = (0, 0, "png ", png), [2] = (0, 0, "dupe", toFirst) };
                return CreateFontFromCmap(CreateDistinctFormat12Cmap(new[] { 0x1F600, 0x1F601 }), includeTrailingMetric: true, glyphCount: 3,
                    blankGlyphIds: new[] { 1, 2 }, additionalTables: new[] { ("sbix", PadToFourBytes(CreateSbixStrikes(3,
                        (ppem * 2, new Dictionary<int, (short, short, string, byte[])> { [1] = (0, 0, "png ", png) }), (ppem, small)))) });
            }
            case SbixVariant.DuplicateWithOwnOutline: {
                byte[] toFirst = new byte[2];
                WriteUInt16(toFirst, 0, 1);
                var glyphs = new Dictionary<int, (short, short, string, byte[])> { [1] = (0, 0, "png ", png), [2] = (0, 0, "dupe", toFirst) };
                return CreateFontFromCmap(CreateDistinctFormat12Cmap(new[] { 0x1F600, 0x1F601 }), includeTrailingMetric: true, glyphCount: 3,
                    glyphHeights: new Dictionary<int, int> { [1] = 700, [2] = 700 },
                    glyphLefts: new Dictionary<int, int> { [1] = SbixOutlineOrigin.Left, [2] = SbixDuplicateOutlineOrigin.Left },
                    glyphBottoms: new Dictionary<int, int> { [1] = SbixOutlineOrigin.Bottom, [2] = SbixDuplicateOutlineOrigin.Bottom },
                    additionalTables: new[] { ("sbix", PadToFourBytes(CreateSbixStrikes(3, (ppem, glyphs)))) });
            }
            case SbixVariant.MissingFromLargestStrike:
                return CreateFontFromCmap(CreateFormat12Cmap(new[] { 0x1F600 }), includeTrailingMetric: true, glyphCount: 2, blankGlyphIds: new[] { 1 },
                    additionalTables: new[] { ("sbix", PadToFourBytes(CreateSbixStrikes(2,
                        (ppem * 4, new Dictionary<int, (short, short, string, byte[])>()), (ppem, glyph1)))) });
            default:
                return CreateFontFromCmap(CreateFormat12Cmap(new[] { 0x1F600 }), includeTrailingMetric: true, glyphCount: 2,
                    glyphHeights: new Dictionary<int, int> { [1] = 700 },
                    glyphLefts: new Dictionary<int, int> { [1] = SbixOutlineOrigin.Left },
                    glyphBottoms: new Dictionary<int, int> { [1] = SbixOutlineOrigin.Bottom },
                    additionalTables: new[] { ("sbix", PadToFourBytes(CreateSbixStrikes(2, (ppem, glyph1)))) });
        }
    }

    /// <summary>
    /// An sbix font whose glyphs 1 to <paramref name="count"/> each have their own red PNG record, mapped from
    /// U+1F600 onwards.
    /// </summary>
    internal static byte[] CreateManyGlyphSbixFont(int count) {
        byte[] png = CreateRedPng();
        var glyphs = new Dictionary<int, (short, short, string, byte[])>();
        for (int glyph = 1; glyph <= count; glyph++) glyphs[glyph] = (0, 0, "png ", png);
        int[] scalars = Enumerable.Range(0x1F600, count).ToArray();
        return CreateFontFromCmap(CreateDistinctFormat12Cmap(scalars), includeTrailingMetric: true, glyphCount: count + 1,
            blankGlyphIds: Enumerable.Range(1, count).ToArray(),
            additionalTables: new[] { ("sbix", PadToFourBytes(CreateSbixStrikes(count + 1, (BitmapColorFontPpem * 4 / 80, glyphs)))) });
    }

    // A 4x4 solid red PNG: strikes scale it to the requested glyph size.
    private static byte[] CreateRedPng() =>
        OfficeRasterImageEncoder.Encode(new OfficeRasterImage(4, 4, OfficeColor.FromRgba(255, 0, 0, 255)), OfficeImageExportFormat.Png);

    // The font builder places tables back to back; readers require each to start on a four-byte boundary.
    private static byte[] PadToFourBytes(byte[] table) {
        if (table.Length % 4 == 0) return table;
        var padded = new byte[(table.Length + 3) / 4 * 4];
        Array.Copy(table, padded, table.Length);
        return padded;
    }

    // CBDT version 3.0, then one glyph in image format 17: small metrics, PNG length, PNG data.
    private static byte[] CreateCbdt(byte[] png, int pixels) {
        var data = new byte[4 + 5 + 4 + png.Length];
        WriteUInt16(data, 0, 3);
        data[4] = checked((byte)pixels); // height
        data[5] = checked((byte)pixels); // width
        data[6] = 0;                    // bearingX
        data[7] = checked((byte)pixels); // bearingY: the top edge, so the image sits on the baseline
        data[8] = checked((byte)pixels); // advance
        WriteUInt32(data, 9, checked((uint)png.Length));
        Array.Copy(png, 0, data, 13, png.Length);
        return data;
    }

    // CBLC version 3.0 with one strike covering one glyph through index format 1.
    private static byte[] CreateCblc(int pngLength, int glyph) {
        var data = new byte[8 + 48 + 8 + 8 + 8];
        WriteUInt16(data, 0, 3);
        WriteUInt32(data, 4, 1);
        int size = 8;
        WriteUInt32(data, size, 56);       // indexSubTableArrayOffset
        WriteUInt32(data, size + 4, 24);   // indexTablesSize
        WriteUInt32(data, size + 8, 1);    // numberOfIndexSubTables
        WriteUInt16(data, size + 40, checked((ushort)glyph)); // startGlyphIndex
        WriteUInt16(data, size + 42, checked((ushort)glyph)); // endGlyphIndex
        data[size + 44] = BitmapColorFontPpem;
        data[size + 45] = BitmapColorFontPpem;
        data[size + 46] = 32;
        data[size + 47] = 1;
        WriteUInt16(data, 56, checked((ushort)glyph)); // firstGlyphIndex
        WriteUInt16(data, 58, checked((ushort)glyph)); // lastGlyphIndex
        WriteUInt32(data, 60, 8);          // additionalOffsetToIndexSubtable
        WriteUInt16(data, 64, 1);          // indexFormat
        WriteUInt16(data, 66, 17);         // imageFormat
        WriteUInt32(data, 68, 4);          // imageDataOffset in CBDT
        WriteUInt32(data, 72, 0);
        WriteUInt32(data, 76, checked((uint)(5 + 4 + pngLength)));
        return data;
    }

    // sbix version 1 with the given strikes. Each glyph record holds origin offsets in strike pixels, a graphic
    // type and its data; glyphs without a record are empty in that strike.
    private static byte[] CreateSbixStrikes(int glyphCount, params (int Ppem, IReadOnlyDictionary<int, (short OriginX, short OriginY, string Type, byte[] Data)> Glyphs)[] strikes) {
        var table = new List<byte>();
        void U16(int value) { table.Add((byte)(value >> 8)); table.Add((byte)value); }
        void U32(int value) { table.Add((byte)(value >> 24)); table.Add((byte)(value >> 16)); table.Add((byte)(value >> 8)); table.Add((byte)value); }
        void Patch32(int at, int value) { table[at] = (byte)(value >> 24); table[at + 1] = (byte)(value >> 16); table[at + 2] = (byte)(value >> 8); table[at + 3] = (byte)value; }
        U16(1); U16(1); U32(strikes.Length);
        int strikeOffsets = table.Count;
        foreach (var _ in strikes) U32(0);
        for (int strikeIndex = 0; strikeIndex < strikes.Length; strikeIndex++) {
            int strike = table.Count;
            Patch32(strikeOffsets + strikeIndex * 4, strike);
            U16(strikes[strikeIndex].Ppem); U16(72);
            int offsets = table.Count;
            for (int index = 0; index <= glyphCount; index++) U32(0);
            for (int glyph = 0; glyph <= glyphCount; glyph++) {
                Patch32(offsets + glyph * 4, table.Count - strike);
                if (glyph == glyphCount || !strikes[strikeIndex].Glyphs.TryGetValue(glyph, out var record)) continue;
                U16(record.OriginX); U16(record.OriginY);
                foreach (char c in record.Type) table.Add((byte)c);
                table.AddRange(record.Data);
            }
        }
        return table.ToArray();
    }

    /// <summary>
    /// A COLR v0 glyph (red then blue layers) whose CPAL palette has <paramref name="paletteEntries"/> entries,
    /// as Segoe UI Emoji uses tens of thousands of entries in one palette. Either layer's palette entry can be translucent.
    /// </summary>
    internal static byte[] CreateColorFontWithPaletteSize(int scalar, int paletteEntries, byte redAlpha = 255, byte blueAlpha = 255) {
        var cpal = new byte[12 + 2 + paletteEntries * 4];
        WriteUInt16(cpal, 0, 0);
        WriteUInt16(cpal, 2, checked((ushort)paletteEntries));
        WriteUInt16(cpal, 4, 1);
        WriteUInt16(cpal, 6, checked((ushort)paletteEntries));
        WriteUInt32(cpal, 8, 14);
        WriteUInt16(cpal, 12, 0);
        for (int entry = 0; entry < paletteEntries; entry++) {
            if (entry == 0) WriteBgra(cpal, 14, 255, 0, 0, redAlpha);
            else if (entry == 1) WriteBgra(cpal, 18, 0, 0, 255, blueAlpha);
            else WriteBgra(cpal, 14 + entry * 4, 128, 128, 128, 255);
        }
        return CreateFontFromCmap(CreateFormat12Cmap(new[] { scalar }), glyphCount: 4, distinctSecondGlyph: true,
            colr: CreateColrV0(), cpal: cpal);
    }

    /// <summary>
    /// The <see cref="CreateColrV1Font"/> glyph with its first layer painted in the text color at half alpha and its
    /// gradient layer fully transparent, so the only ink is the half-transparent foreground layer.
    /// </summary>
    internal static byte[] CreateHalfAlphaForegroundColrFont(int scalar) {
        byte[] font = CreateColrV1Font(scalar);
        int tables = ReadUInt16(font, 4);
        for (int index = 0; index < tables; index++) {
            int record = 12 + index * 16;
            if (ReadUInt32(font, record) != 0x434F4C52) continue; // 'COLR'
            int colr = checked((int)ReadUInt32(font, record + 8));
            WriteUInt16(font, colr + 69, 0xFFFF); // PaintSolid paletteIndex: the text color
            WriteUInt16(font, colr + 71, 0x2000); // PaintSolid alpha 0.5
            WriteUInt16(font, colr + 109, 0);     // gradient stop alphas
            WriteUInt16(font, colr + 115, 0);
            return font;
        }
        throw new InvalidOperationException("The COLR v1 fixture has no COLR table.");
    }
}
