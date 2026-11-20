using OfficeIMO.Drawing;

namespace OfficeIMO.TestAssets;

internal static partial class ManagedTextShapingTestAssets {
    /// <summary>Units per em of <see cref="CreateBitmapColorFont"/>; its PNG strike is drawn at <see cref="BitmapColorFontPpem"/>.</summary>
    internal const int BitmapColorFontPpem = 100;

    /// <summary>
    /// One glyph whose only artwork is a solid red PNG, stored the way Noto Color Emoji (CBDT/CBLC, image format 17)
    /// or Apple Color Emoji (sbix, <c>png </c>) store it. The glyph outline is empty. The bitmap is
    /// <paramref name="pixels"/> square at <see cref="BitmapColorFontPpem"/> ppem, sitting on the baseline.
    /// </summary>
    internal static byte[] CreateBitmapColorFont(int scalar, bool sbix, int pixels = 80) {
        var image = new OfficeRasterImage(4, 4, OfficeColor.FromRgba(255, 0, 0, 255));
        byte[] png = OfficeRasterImageEncoder.Encode(image, OfficeImageExportFormat.Png);
        var tables = sbix ? new[] { ("sbix", PadToFourBytes(CreateSbix(png, pixels, glyphCount: 2))) }
            : new[] { ("CBDT", PadToFourBytes(CreateCbdt(png, pixels))), ("CBLC", CreateCblc(png.Length)) };
        return CreateFontFromCmap(CreateFormat12Cmap(new[] { scalar }), includeTrailingMetric: true, glyphCount: 2, blankGlyphIds: new[] { 1 },
            additionalTables: tables);
    }

    // The font builder places tables back to back; readers require each to start on a four-byte boundary.
    private static byte[] PadToFourBytes(byte[] table) {
        if (table.Length % 4 == 0) return table;
        var padded = new byte[(table.Length + 3) / 4 * 4];
        Array.Copy(table, padded, table.Length);
        return padded;
    }

    // CBDT version 3.0, then glyph 1 in image format 17: small metrics, PNG length, PNG data.
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

    // CBLC version 3.0 with one strike covering glyph 1 through index format 1.
    private static byte[] CreateCblc(int pngLength) {
        var data = new byte[8 + 48 + 8 + 8 + 8];
        WriteUInt16(data, 0, 3);
        WriteUInt32(data, 4, 1);
        int size = 8;
        WriteUInt32(data, size, 56);       // indexSubTableArrayOffset
        WriteUInt32(data, size + 4, 24);   // indexTablesSize
        WriteUInt32(data, size + 8, 1);    // numberOfIndexSubTables
        WriteUInt16(data, size + 40, 1);   // startGlyphIndex
        WriteUInt16(data, size + 42, 1);   // endGlyphIndex
        data[size + 44] = BitmapColorFontPpem;
        data[size + 45] = BitmapColorFontPpem;
        data[size + 46] = 32;
        data[size + 47] = 1;
        WriteUInt16(data, 56, 1);          // firstGlyphIndex
        WriteUInt16(data, 58, 1);          // lastGlyphIndex
        WriteUInt32(data, 60, 8);          // additionalOffsetToIndexSubtable
        WriteUInt16(data, 64, 1);          // indexFormat
        WriteUInt16(data, 66, 17);         // imageFormat
        WriteUInt32(data, 68, 4);          // imageDataOffset in CBDT
        WriteUInt32(data, 72, 0);
        WriteUInt32(data, 76, checked((uint)(5 + 4 + pngLength)));
        return data;
    }

    // sbix version 1 with one strike; glyph 1 is a PNG whose lower-left corner is at the glyph origin.
    private static byte[] CreateSbix(byte[] png, int pixels, int glyphCount) {
        // The fixture PNG is 4x4; sbix takes its size from the image, so scale the strike instead.
        int ppem = BitmapColorFontPpem * 4 / pixels;
        int strike = 12, offsets = strike + 4, glyphData = offsets + (glyphCount + 1) * 4;
        var data = new byte[glyphData + 8 + png.Length];
        WriteUInt16(data, 0, 1);
        WriteUInt16(data, 2, 1);
        WriteUInt32(data, 4, 1);
        WriteUInt32(data, 8, checked((uint)strike));
        WriteUInt16(data, strike, checked((ushort)ppem));
        WriteUInt16(data, strike + 2, 72);
        int glyphStart = glyphData - strike;
        WriteUInt32(data, offsets, checked((uint)glyphStart));      // glyph 0: empty
        WriteUInt32(data, offsets + 4, checked((uint)glyphStart));  // glyph 1 starts here
        WriteUInt32(data, offsets + 8, checked((uint)(glyphStart + 8 + png.Length)));
        WriteTag(data, glyphData + 4, "png ");
        Array.Copy(png, 0, data, glyphData + 8, png.Length);
        return data;
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
}
