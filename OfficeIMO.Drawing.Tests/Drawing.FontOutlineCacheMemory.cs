using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

[CollectionDefinition("Font outline cache memory", DisableParallelization = true)]
public sealed class FontOutlineCacheMemoryCollection { }

[Collection("Font outline cache memory")]
public sealed class FontOutlineCacheMemoryTests {
    [Fact]
    public void DecodedOutlinesRetainBoundedMemoryAcrossBoundedRequests() {
        // Each glyph is a triangle followed by coincident points, 65535 points in 532 bytes. Every
        // request stays within its point budget, so only a font-wide limit keeps the decoded
        // outlines (about 1.1 MB each) from accumulating while the font is alive.
        const int glyphCount = 192;
        const int pointCount = 65535;
        const int glyphLength = 532;
        var glyf = new byte[glyphCount * glyphLength];
        var loca = new byte[(glyphCount + 2) * 2];
        for (int glyph = 1; glyph <= glyphCount; glyph++) {
            int offset = (glyph - 1) * glyphLength;
            DrawingTests.WriteUInt16(loca, glyph * 2, offset / 2);
            DrawingTests.WriteUInt16(glyf, offset, 1);
            DrawingTests.WriteUInt16(glyf, offset + 10, pointCount - 1);
            int p = offset + 14;
            glyf[p++] = 0x31; // on curve at the origin
            glyf[p++] = 0x33; // x + 100
            glyf[p++] = 0x35; // y + 100
            for (int remaining = pointCount - 3; remaining > 0; remaining -= 256) {
                glyf[p++] = 0x39; // on curve, repeated, at the previous point
                glyf[p++] = (byte)(Math.Min(remaining, 256) - 1);
            }
            glyf[p++] = 100;
            glyf[p] = 100;
        }
        DrawingTests.WriteUInt16(loca, (glyphCount + 1) * 2, glyf.Length / 2);
        byte[] data = DrawingTests.CreateMinimalTrueTypeFont(CreateSequentialCmap(glyphCount),
            glyf, loca, glyphCount: (ushort)(glyphCount + 1));
        OfficeTrueTypeFont font = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(data));

        long before = GC.GetTotalMemory(forceFullCollection: true);
        for (int glyph = 0; glyph < glyphCount; glyph++) {
            Assert.NotEmpty(font.GetTextContoursBounded(((char)('A' + glyph)).ToString(), 0, 0, 12,
                pointCount, CancellationToken.None));
        }
        long retained = GC.GetTotalMemory(forceFullCollection: true) - before;
        GC.KeepAlive(font);

        Assert.True(retained < 64L * 1024 * 1024, $"Retained {retained} bytes of decoded outlines.");
    }

    // Format 4 cmap mapping 'A' and the following characters to glyphs 1 to glyphCount.
    private static byte[] CreateSequentialCmap(int glyphCount) {
        var data = new byte[12 + 32];
        DrawingTests.WriteUInt16(data, 2, 1);
        DrawingTests.WriteUInt16(data, 4, 3);
        DrawingTests.WriteUInt16(data, 6, 1);
        data[11] = 12;
        DrawingTests.WriteUInt16(data, 12, 4);
        DrawingTests.WriteUInt16(data, 14, 32);
        DrawingTests.WriteUInt16(data, 18, 4);
        DrawingTests.WriteUInt16(data, 20, 4);
        DrawingTests.WriteUInt16(data, 22, 1);
        DrawingTests.WriteUInt16(data, 26, 0x0041 + glyphCount - 1);
        DrawingTests.WriteUInt16(data, 28, 0xFFFF);
        DrawingTests.WriteUInt16(data, 32, 0x0041);
        DrawingTests.WriteUInt16(data, 34, 0xFFFF);
        DrawingTests.WriteUInt16(data, 36, unchecked((ushort)(1 - 0x0041)));
        DrawingTests.WriteUInt16(data, 38, 1);
        return data;
    }
}
