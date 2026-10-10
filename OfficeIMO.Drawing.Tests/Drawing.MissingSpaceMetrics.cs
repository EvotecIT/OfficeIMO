using System.Buffers.Binary;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingMissingSpaceMetricsTests {
    [Theory]
    [InlineData('\u2003', 16D, 250)]
    [InlineData('\u2003', 32D, 825)]
    [InlineData('\u2001', 16D, 825)]
    [InlineData('\u2001', 32D, 250)]
    public void MissingEmSeparatorUsesOneEmInsteadOfNotdefAdvance(char separator, double size, int notdefAdvance) {
        IOfficeFontProgram font = LoadFont(FontWithAdvance((ushort)notdefAdvance, 'A'));
        string space = separator.ToString();

        Assert.False(font.HasGlyphs(space));
        Assert.False(font.TryGetGlyphMetrics(separator, out _, out _));
        Assert.Equal(size, font.Measure(space, size), 6);
        Assert.Equal(size, Assert.Single(font.MeasureTextElements(new[] { space }, size)), 6);
        Assert.Empty(font.GetTextContours(space, 0, 0, size));
    }

    [Theory]
    [InlineData('\u2003', false)]
    [InlineData('\u2001', false)]
    [InlineData('\u2003', true)]
    [InlineData('\u2001', true)]
    public void MappedEmSeparatorKeepsItsFontAdvance(char separator, bool cff) {
        byte[] data = FontWithAdvance(725, 'A', separator, ' ', '\u00A0');
        IOfficeFontProgram font = LoadFont(cff ? AsCff(data) : data);

        Assert.True(font.HasGlyphs(separator.ToString()));
        Assert.Equal(14.5D, font.Measure(separator.ToString(), 20D), 6);
        Assert.Equal(14.5D, Assert.Single(font.MeasureTextElements(new[] { separator.ToString() }, 20D)), 6);
        Assert.Equal(14.5D, font.Measure(" ", 20D), 6);
        Assert.Equal(14.5D, font.Measure("\u00A0", 20D), 6);
    }

    [Theory]
    [InlineData('\u2003')]
    [InlineData('\u2001')]
    public void RegisteredFallbackGlyphWinsBeforeMissingSpaceSynthesis(char separator) {
        var fonts = new OfficeFontFaceCollection()
            .Add("Primary", FontWithAdvance(250, 'A'))
            .Add("Fallback", FontWithAdvance(725, separator))
            .AddFallbackFamily("Fallback");
        string space = separator.ToString();

        Assert.True(fonts.TryResolveFaceForText(space, "Primary, Fallback", OfficeFontStyle.Regular, out OfficeFontFace? styled));
        Assert.Equal("Fallback", styled!.FamilyName);
        Assert.True(fonts.TryResolveFaceForText(space, "Primary", OfficeFontFaceDescriptor.Regular, out OfficeFontFace? numeric));
        Assert.Equal("Fallback", numeric!.FamilyName);
        OfficeFontFallbackRun[] runs = fonts.PlanFallbackRuns("A" + space + "A", "Primary").ToArray();
        Assert.Equal(new[] { "Primary", "Fallback", "Primary" }, runs.Select(run => run.FamilyName));
        Assert.True(fonts.TryMeasureText("A" + space + "A", 20D, "Primary", OfficeFontStyle.Regular, out double width));
        Assert.Equal(24.5D, width, 6);
    }

    [Fact]
    public void CoveredWhitespaceKeepsTheCurrentFallbackFamily() {
        var fonts = new OfficeFontFaceCollection()
            .Add("Primary", FontWithAdvance(250, '\u2003'))
            .Add("Fallback", FontWithAdvance(725, 'A', '\u2003'))
            .AddFallbackFamily("Fallback");

        OfficeFontFallbackRun run = Assert.Single(fonts.PlanFallbackRuns("A\u2003A", "Primary"));

        Assert.Equal("Fallback", run.FamilyName);
        Assert.True(fonts.TryMeasureText(run.Text, 20D, run.FamilyName, OfficeFontStyle.Regular, out double width));
        Assert.Equal(43.5D, width, 6);
    }

    [Fact]
    public void ScopedMeasurementSynthesizesOnlyWhenNoRegisteredGlyphCoversTheSeparator() {
        var fonts = new OfficeFontFaceCollection().Add("Primary", FontWithAdvance(250, 'A'));

        Assert.True(fonts.TryMeasureText("A\u2003A", 20D, "Primary", OfficeFontStyle.Regular, out double width));
        Assert.Equal(30D, width, 6);
        Assert.True(fonts.TryResolveFaceForText("\u2003", "Primary", OfficeFontFaceDescriptor.Regular, out OfficeFontFace? face));
        Assert.Equal("Primary", face!.FamilyName);
        Assert.False(face.Program.HasGlyphs("\u2003"));
        // Existing ordinary whitespace/control handling does not acquire EM semantics.
        Assert.True(face.Program.HasGlyphs(" \u00A0\t"));
        Assert.Equal(15D, face.Program.Measure(" \u00A0\t", 20D), 6);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MissingEmSeparatorKeepsMeasurementAndPaintCursorInAgreement(bool cff) {
        byte[] data = FontWithAdvance(250, 'A');
        IOfficeFontProgram font = cff ? LoadFont(AsCff(data)) : LoadFont(data);
        const double size = 20D;
        double letterWidth = font.Measure("A", size);
        IReadOnlyList<double> elements = font.MeasureTextElements(new[] { "A", "\u2003", "A" }, size);
        List<List<OfficePoint>> single = font.GetTextContours("A", 0, 0, size);
        List<List<OfficePoint>> mixed = font.GetTextContours("A\u2003A", 0, 0, size);

        Assert.False(font.HasGlyphs("\u2003"));
        Assert.Equal(new[] { letterWidth, size, letterWidth }, elements);
        Assert.Equal(elements.Sum(), font.Measure("A\u2003A", size), 6);
        Assert.Equal(single.Count * 2, mixed.Count);
        double firstInkX = single.SelectMany(contour => contour).Min(point => point.X);
        double secondInkX = mixed.Skip(single.Count).SelectMany(contour => contour).Min(point => point.X);
        Assert.Equal(firstInkX + letterWidth + size, secondInkX, 6);
        Assert.Empty(font.GetTextContours("\u2003", 0, 0, size));
    }

    [Fact]
    public void DirectColorContoursKeepTheMissingEmSeparatorAdvance() {
        OfficeTrueTypeFont font = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(
            ManagedTextShapingTestAssets.CreateColorFont('A')));
        Assert.True(font.TryGetColorTextContours("A", 0, 0, 20D, null, OfficeColor.Black, int.MaxValue,
            default, out List<OfficeColorGlyphContours> single));
        Assert.True(font.TryGetColorTextContours("A\u2003A", 0, 0, 20D, null, OfficeColor.Black, int.MaxValue,
            default, out List<OfficeColorGlyphContours> mixed));

        Assert.Equal(single.Count * 2, mixed.Count);
        double firstInkX = single.SelectMany(layer => layer.Contours).SelectMany(contour => contour).Min(point => point.X);
        double secondInkX = mixed.Skip(single.Count).SelectMany(layer => layer.Contours).SelectMany(contour => contour).Min(point => point.X);
        Assert.Equal(firstInkX + font.Measure("A", 20D) + 20D, secondInkX, 6);
    }

    [Fact]
    public void SynthesizedSeparatorDoesNotInheritNotdefPairKerning() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'B', -125);
        int count = BinaryPrimitives.ReadUInt16BigEndian(data.AsSpan(4));
        bool patched = false;
        for (int index = 0; index < count; index++) {
            int record = 12 + index * 16;
            if (Encoding.ASCII.GetString(data, record, 4) != "kern") continue;
            int offset = checked((int)BinaryPrimitives.ReadUInt32BigEndian(data.AsSpan(record + 8)));
            // Existing format-zero pair A/B becomes A/.notdef. Missing EM
            // synthesis must not consume that unrelated glyph-zero adjustment.
            BinaryPrimitives.WriteUInt16BigEndian(data.AsSpan(offset + 20), 0);
            patched = true;
            break;
        }
        Assert.True(patched);
        IOfficeFontProgram font = LoadFont(data);

        Assert.Equal(30D, font.Measure("A\u2003", 20D), 6);
        Assert.Equal(new[] { 10D, 20D }, font.MeasureTextElements(new[] { "A", "\u2003" }, 20D));
    }

    [Fact]
    public void DirectRasterFontKeepsItsSyntheticBlankDuringMeasurementAndPainting() {
        OfficeTrueTypeFont font = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(FontWithAdvance(250, 'A')));
        var image = new OfficeRasterImage(60, 30, OfficeColor.White);
        var canvas = new OfficeRasterCanvas(image, font);

        Assert.Equal(30D, canvas.MeasureText("A\u2003A", 20D), 6);
        canvas.DrawTextLine("A\u2003A", 0, 0, 20D, OfficeColor.Black, alignment: OfficeTextAlignment.Left);

        Assert.Contains(Enumerable.Range(0, image.Height), y => image.GetPixel(27, y).R < 128);
        Assert.All(Enumerable.Range(12, 11), x =>
            Assert.All(Enumerable.Range(0, image.Height), y => Assert.Equal(OfficeColor.White, image.GetPixel(x, y))));
    }

    [Fact]
    public void BoundedBidiFallbackKeepsVisualOrderAroundMissingEmSpace() {
        IOfficeFontProgram font = LoadFont(FontWithAdvance(500, 0x05D0, 0x05D1));

        OfficeManagedTextFallback fallback = OfficeManagedTextShaper.Resolve("\u05D0\u2003\u05D1", font,
            OfficeTextDirection.RightToLeft);

        Assert.True(fallback.Used);
        Assert.False(fallback.Incomplete);
        Assert.Equal("\u05D1\u2003\u05D0", fallback.Text);
    }

    private static IOfficeFontProgram LoadFont(byte[] data) =>
        Assert.Single(new OfficeFontFaceCollection().Add("Synthetic", data).Faces).Program;

    private static byte[] FontWithAdvance(ushort advance, params int[] scalars) {
        byte[] data = ManagedTextShapingTestAssets.CreateFont(scalars);
        int count = BinaryPrimitives.ReadUInt16BigEndian(data.AsSpan(4));
        for (int index = 0; index < count; index++) {
            int record = 12 + index * 16;
            if (Encoding.ASCII.GetString(data, record, 4) != "hmtx") continue;
            int offset = checked((int)BinaryPrimitives.ReadUInt32BigEndian(data.AsSpan(record + 8)));
            BinaryPrimitives.WriteUInt16BigEndian(data.AsSpan(offset), advance);
            return data;
        }
        throw new InvalidOperationException("The synthetic font must contain horizontal metrics.");
    }

    private static byte[] AsCff(byte[] trueType) {
        // Reuse the existing synthetic font's cmap and metrics. Both .notdef and
        // A have ink, so missing EM must stay blank as well as move the cursor.
        byte[] cff = {
            1, 0, 4, 1,
            0, 1, 1, 1, 2, (byte)'A',
            0, 1, 1, 1, 3, 160, 17,
            0, 0, 0, 0,
            0, 2, 1, 1, 12, 23,
            139, 139, 21, 239, 139, 139, 239, 39, 39, 5, 14,
            139, 139, 21, 239, 139, 139, 239, 39, 39, 5, 14
        };
        var tables = new List<(string Tag, byte[] Data)>();
        int count = BinaryPrimitives.ReadUInt16BigEndian(trueType.AsSpan(4));
        for (int index = 0; index < count; index++) {
            int record = 12 + index * 16;
            string tag = Encoding.ASCII.GetString(trueType, record, 4);
            if (tag == "glyf" || tag == "loca") continue;
            int offset = checked((int)BinaryPrimitives.ReadUInt32BigEndian(trueType.AsSpan(record + 8)));
            int length = checked((int)BinaryPrimitives.ReadUInt32BigEndian(trueType.AsSpan(record + 12)));
            tables.Add((tag, trueType.AsSpan(offset, length).ToArray()));
        }
        tables.Add(("CFF ", cff));
        int dataOffset = 12 + tables.Count * 16;
        byte[] result = new byte[dataOffset + tables.Sum(table => (table.Data.Length + 3) & ~3)];
        BinaryPrimitives.WriteUInt32BigEndian(result.AsSpan(), 0x4F54544F);
        BinaryPrimitives.WriteUInt16BigEndian(result.AsSpan(4), checked((ushort)tables.Count));
        for (int index = 0; index < tables.Count; index++) {
            int record = 12 + index * 16;
            Encoding.ASCII.GetBytes(tables[index].Tag).CopyTo(result, record);
            BinaryPrimitives.WriteUInt32BigEndian(result.AsSpan(record + 8), checked((uint)dataOffset));
            BinaryPrimitives.WriteUInt32BigEndian(result.AsSpan(record + 12), checked((uint)tables[index].Data.Length));
            tables[index].Data.CopyTo(result, dataOffset);
            dataOffset += (tables[index].Data.Length + 3) & ~3;
        }
        return result;
    }
}
