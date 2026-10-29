using System;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void CompositeOutlineWithEmptyLeavesConsumesVisitBudget() {
        const int levels = 9;
        const int componentsPerLevel = 32;
        const int glyphLength = 10 + componentsPerLevel * 6;
        var glyf = new byte[levels * glyphLength];
        var loca = new byte[(levels + 2) * 2];

        for (int glyph = 1; glyph <= levels; glyph++) {
            int offset = (glyph - 1) * glyphLength;
            WriteUInt16(loca, glyph * 2, offset / 2);
            WriteUInt16(glyf, offset, -1);
            for (int component = 0; component < componentsPerLevel; component++) {
                int record = offset + 10 + component * 6;
                WriteUInt16(glyf, record, component == componentsPerLevel - 1 ? 0x02 : 0x22);
                WriteUInt16(glyf, record + 2, glyph - 1);
            }
        }
        WriteUInt16(loca, (levels + 1) * 2, glyf.Length / 2);
        byte[] data = CreateMinimalTrueTypeFont(CreateFormat4Cmap(levels, useRangeOffset: false),
            glyf, loca, glyphCount: (ushort)(levels + 1));
        OfficeTrueTypeFont font = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(data));
        using var cancellation = new CancellationTokenSource(TimeSpan.FromSeconds(10));

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            font.GetTextContoursBounded("A", 0, 0, 12, 65536, cancellation.Token));

        Assert.Contains("composite visit budget", exception.Message);
    }
}
