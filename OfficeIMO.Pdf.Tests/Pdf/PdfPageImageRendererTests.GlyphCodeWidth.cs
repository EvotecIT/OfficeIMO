using System;
using System.Globalization;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Theory]
    [InlineData("Identity-H", 0x1784, 1)]
    [InlineData("Identity-H", 0x4184, 2)]
    [InlineData("Identity-V", 0x1784, 2)]
    [InlineData("Identity-V", 0x4184, 1)]
    public void DeclaredIdentityCodesRetainStationaryGlyphMappings(string encoding, int cid, int count) {
        byte[] source = BuildStationaryClusterPdf(encoding, cid, count);
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(source).Pages[0].GetTextSpans());

        Assert.Equal(string.Concat(Enumerable.Repeat("سَّ", count)), span.Text);
        Assert.Equal(0D, span.Advance);
        Assert.Equal(Enumerable.Repeat(3, count), span.GlyphCharacterLengths);
        Assert.Equal(Enumerable.Repeat(0D, count), span.GlyphPaintedAdvances);
        Assert.NotNull(span.GlyphBytes);
        Assert.Equal(count, span.GlyphBytes!.Count);
        Assert.All(span.GlyphBytes, code => Assert.Equal(new[] { (byte)(cid >> 8), (byte)cid }, code));
        Assert.NotNull(span.CharacterAdvances);
        Assert.All(span.CharacterAdvances!, advance => Assert.Equal(0D, advance));
    }

    [Fact]
    public void StationaryClusterReceivesTheEmbeddedPaintedGlyphAlias() {
        PdfReadPage page = PdfReadDocument.Open(BuildStationaryClusterPdf("Identity-H", 0x1784, 1)).Pages[0];
        AssertStationaryClusterPaint(page.ToDrawing());
    }

    [Fact]
    public void StationaryClusterUsesTheNestedFormsLocalIdentityFont() {
        PdfReadPage page = PdfReadDocument.Open(
            BuildStationaryClusterPdf("Identity-H", 0x4184, 1, nestedForm: true)).Pages[0];
        PdfTextSpan span = Assert.Single(page.GetTextSpans());

        Assert.Equal("سَّ", span.Text);
        Assert.NotNull(span.GlyphBytes);
        Assert.Equal(new byte[] { 0x41, 0x84 }, Assert.Single(span.GlyphBytes!));
        Assert.Equal(new[] { 3 }, span.GlyphCharacterLengths);
        Assert.Equal(0D, span.Advance);
        AssertStationaryClusterPaint(page.ToDrawing());
    }

    [Fact]
    public void ActualTextOverridesStationaryClusterExtractionWithoutChangingItsPaint() {
        const string logical = "Cluster replacement";
        PdfReadPage page = PdfReadDocument.Open(
            BuildStationaryClusterPdf("Identity-H", 0x1784, 1, actualText: logical)).Pages[0];

        PdfTextSpan span = Assert.Single(page.GetTextSpans());
        Assert.Equal(logical, span.Text);
        Assert.True(span.HasActualText);
        AssertStationaryClusterPaint(page.ToDrawing());
    }

    private static void AssertStationaryClusterPaint(OfficeDrawing drawing) {
        var retained = Assert.Single(ReplayText(drawing));
        OfficeDrawingText visual = retained.Text;
        Assert.Equal("سَّ", visual.Text);
        Assert.Single(visual.RasterText);
        Assert.InRange(visual.RasterText[0], '\uE000', '\uF8FF');
        OfficeFontFace face = Assert.Single(retained.Owner.Fonts.Faces);
        Assert.True(face.Program.TryGetGlyphMetrics(visual.RasterText[0], out int glyphId, out _));
        Assert.Equal(2, glyphId);
        Assert.Equal(0D, face.Program.Measure(visual.RasterText, 12D));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing);
        Assert.Contains(raster.GetPixels().Where((_, index) => (index & 3) == 3), alpha => alpha > 0);
    }

    [Fact]
    public void UnknownCodeWidthRetainsCompleteRunDecodingFallback() {
        PdfTextSpan span = Assert.Single(TextContentParser.Parse(
            "BT /F1 12 Tf 20 100 Td <1784> Tj ET",
            static (_, code) => code.Length == 2 ? "سَّ" : ((char)code[0]).ToString(),
            static (_, _) => 0D,
            useLogicalTextFilters: false));

        Assert.Equal("سَّ", span.Text);
        Assert.Equal(0D, span.Advance);
        Assert.Null(span.GlyphBytes);
        Assert.Null(span.GlyphCharacterLengths);
        Assert.Null(span.GlyphPaintedAdvances);
    }

    // Existing distinct outlines distinguish the painted second glyph from the first. The font's
    // shared horizontal advance is zero; the CID still names a visible outline for logical text.
    private static byte[] BuildStationaryClusterPdf(string encoding, int cid, int count,
        bool nestedForm = false, string? actualText = null) {
        byte[] program = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphOutlines('A', 'B');
        int tableCount = (program[4] << 8) | program[5];
        for (int index = 0; index < tableCount; index++) {
            int record = 12 + index * 16;
            if (System.Text.Encoding.ASCII.GetString(program, record, 4) != "hmtx") continue;
            int offset = (program[record + 8] << 24) | (program[record + 9] << 16) |
                (program[record + 10] << 8) | program[record + 11];
            program[offset] = program[offset + 1] = 0;
            break;
        }
        byte[] cidMap = new byte[(cid + 1) * 2];
        cidMap[cid * 2 + 1] = 2;
        string code = cid.ToString("X4", CultureInfo.InvariantCulture);
        string text = string.Concat(Enumerable.Repeat(code, count));
        const string resource = "<< /Font << /F1 5 0 R >> >>";
        string font = "5 0 obj\n<< /Type /Font /Subtype /Type0 /BaseFont /Stationary " +
            "/Encoding /" + encoding + " /DescendantFonts [7 0 R] /ToUnicode 6 0 R >>\nendobj";
        string cmap = "begincmap\n1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n" +
            "1 beginbfchar\n<" + code + "> <0633064E0651>\nendbfchar\nendcmap";
        const string descendant = "7 0 obj\n<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Stationary " +
            "/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> " +
            "/DW 0 /FontDescriptor 8 0 R /CIDToGIDMap 10 0 R >>\nendobj";
        const string descriptor = "8 0 obj\n<< /Type /FontDescriptor /FontName /Stationary /Flags 32 " +
            "/FontBBox [0 0 500 700] /Ascent 800 /Descent -200 /CapHeight 700 " +
            "/ItalicAngle 0 /StemV 80 /FontFile2 9 0 R >>\nendobj";
        string content = "BT /F1 12 Tf 20 100 Td <" + text + "> Tj ET";
        if (actualText != null) content = "/Span << /ActualText (" + actualText + ") >> BDC\n" + content + "\nEMC";
        var objects = new[] {
            font, BuildStreamObject(6, "<<", cmap), descendant, descriptor,
            BuildStreamObject(9, "<< /Filter /ASCIIHexDecode", BitConverter.ToString(program).Replace("-", "") + ">"),
            BuildStreamObject(10, "<< /Filter /ASCIIHexDecode", BitConverter.ToString(cidMap).Replace("-", "") + ">")
        };
        if (!nestedForm) return BuildSingleStreamPdf(content, resource, objects);
        string form = BuildStreamObject(11,
            "<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Resources " + resource, content);
        return BuildSingleStreamPdf("/Fm Do", "<< /XObject << /Fm 11 0 R >> >>",
            objects.Concat(new[] { form }).ToArray());
    }
}
