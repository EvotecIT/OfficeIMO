using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingCffFontTests {
    [Fact]
    public void Cff1FontMeasuresAndProducesContoursWithoutExternalProvider() {
        byte[] data = ReadAsset("SourceSansPro-Regular.otf");
        var fonts = new OfficeFontFaceCollection().Add("Source Sans Pro CFF", data);

        OfficeFontFace face = Assert.Single(fonts.Faces);
        Assert.True(face.Program.IsOpenTypeCff);
        Assert.True(face.CanEmbedAsStaticPdfFont);
        Assert.True(face.Program.HasGlyphs("CFF office ffi"));
        Assert.True(face.Program.Measure("CFF office ffi", 24D) > 100D);
        Assert.NotEmpty(face.Program.GetTextContours("OfficeIMO", 0D, 0D, 24D));
    }

    [Fact]
    public void Cff2VariationAxesProduceDeterministicDistinctContours() {
        byte[] data = ReadAsset("AdobeVFPrototype-Subset.otf");
        var defaultFont = new OfficeFontFaceCollection();
        Assert.True(defaultFont.TryAddBounded(
            "Adobe Variable CFF2",
            data,
            OfficeFontStyle.Regular,
            OfficeFontUnicodeRangeSet.All,
            16 * 1024 * 1024,
            out _,
            out string? defaultError), defaultError);
        var selectedFont = new OfficeFontFaceCollection {
            FontVariationResolver = _ => new Dictionary<string, float> {
                ["wght"] = 700F,
                ["xxxx"] = 75F
            }
        };
        Assert.True(selectedFont.TryAddBounded(
            "Adobe Variable CFF2",
            data,
            OfficeFontStyle.Regular,
            OfficeFontUnicodeRangeSet.All,
            16 * 1024 * 1024,
            out _,
            out string? selectedError), selectedError);

        OfficeFontFace defaultFace = Assert.Single(defaultFont.Faces);
        OfficeFontFace selectedFace = Assert.Single(selectedFont.Faces);
        Assert.False(defaultFace.CanEmbedAsStaticPdfFont);
        Assert.False(selectedFace.CanEmbedAsStaticPdfFont);
        Assert.NotEqual(defaultFace.Program.Fingerprint, selectedFace.Program.Fingerprint);
        Assert.True(defaultFace.Program.TryGetGlyphMetrics('$', out int defaultGlyph, out int defaultAdvance));
        Assert.True(selectedFace.Program.TryGetGlyphMetrics('$', out int selectedGlyph, out int selectedAdvance));
        Assert.Equal(defaultGlyph, selectedGlyph);
        Assert.Equal(560, defaultAdvance);
        Assert.Equal(530, selectedAdvance);
        IReadOnlyList<List<OfficePoint>> defaultContours = defaultFace.Program.GetTextContours("$$$", 0D, 0D, 24D);
        IReadOnlyList<List<OfficePoint>> selectedContours = selectedFace.Program.GetTextContours("$$$", 0D, 0D, 24D);
        Assert.NotEmpty(defaultContours);
        Assert.NotEmpty(selectedContours);
        Assert.NotEqual(Serialize(defaultContours), Serialize(selectedContours));
        Assert.Equal(Serialize(selectedContours), Serialize(selectedFace.Program.GetTextContours("$$$", 0D, 0D, 24D)));
    }

    [Fact]
    public void Cff1SeacCompatibleEndCharComposesBaseAndAccent() {
        byte[] data = ReadAsset("SourceSansPro-Regular.otf");
        OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(data));
        OfficeCffFontData cff = OfficeCffFontData.Parse(reader, OfficeFontVariationModel.None);
        int baseGlyph = cff.ResolveStandardEncodingGlyph(65);  // A
        int accentGlyph = cff.ResolveStandardEncodingGlyph(194); // acute
        int targetGlyph = -1;
        OfficeCffFontData.CffSlice target = default;
        for (int glyph = 1; glyph < cff.GlyphCount; glyph++) {
            if (glyph == baseGlyph || glyph == accentGlyph) continue;
            OfficeCffFontData.CffSlice candidate = cff.GetCharString(glyph);
            if (candidate.Length < 6) continue;
            targetGlyph = glyph;
            target = candidate;
            break;
        }
        Assert.True(targetGlyph > 0);
        target.Data[target.Offset] = 139;     // adx = 0
        target.Data[target.Offset + 1] = 139; // ady = 0
        target.Data[target.Offset + 2] = 204; // bchar = StandardEncoding A (65)
        target.Data[target.Offset + 3] = 247; // achar = StandardEncoding acute (194)
        target.Data[target.Offset + 4] = 86;
        target.Data[target.Offset + 5] = 14;  // endchar
        var sink = new CountingCffSink();

        new OfficeType2CharStringInterpreter(cff, targetGlyph, sink, CancellationToken.None).Render(target);

        Assert.True(sink.MoveCount >= 2);
        Assert.True(sink.DrawingOperationCount > 2);
    }

    [Fact]
    public void Cff2VariationStoreAllowsNullSelectionsUntilReferenced() {
        byte[] source = ReadAsset("AdobeVFPrototype-Subset.otf");
        int storeLength = 16 + 2 * 6;
        byte[] data = new byte[source.Length + 2 + storeLength];
        Buffer.BlockCopy(source, 0, data, 0, source.Length);
        int outerStore = source.Length;
        int itemStore = outerStore + 2;
        WriteUInt16(data, outerStore, storeLength);
        WriteUInt16(data, itemStore, 1); // format
        WriteUInt32(data, itemStore + 2, 12); // VariationRegionList offset
        WriteUInt16(data, itemStore + 6, 1); // ItemVariationData count
        WriteUInt32(data, itemStore + 8, 0); // valid NULL ItemVariationData offset
        WriteUInt16(data, itemStore + 12, 2); // fixture axis count
        WriteUInt16(data, itemStore + 14, 1); // one neutral region

        OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(data));
        OfficeFontVariationModel model = OfficeFontVariationModel.Create(
            reader,
            new Dictionary<string, float> { ["wght"] = 700F, ["xxxx"] = 75F });
        OfficeCffVariationStore variationStore = OfficeCffVariationStore.Parse(
            reader,
            outerStore,
            data.Length,
            model);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => variationStore.GetScalars(0));
        Assert.Contains("null ItemVariationData", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Cff2BlendAppliesRegionMajorDeltasToEachValue() {
        var stack = new List<double> {
            10D, 20D,
            1D, 2D,
            3D, 4D
        };

        OfficeType2CharStringInterpreter.ApplyBlendDeltas(
            stack,
            start: 0,
            valueCount: 2,
            scalars: new[] { 0.5D, 0.25D });

        Assert.Equal(11.25D, stack[0], 6);
        Assert.Equal(22D, stack[1], 6);
    }

    private static byte[] ReadAsset(string name) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", name));

    private static void WriteUInt16(byte[] data, int offset, int value) {
        data[offset] = (byte)(value >> 8);
        data[offset + 1] = (byte)value;
    }

    private static void WriteUInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }

    private static string Serialize(IReadOnlyList<List<OfficePoint>> contours) {
        var text = new System.Text.StringBuilder();
        foreach (List<OfficePoint> contour in contours) {
            foreach (OfficePoint point in contour) text.Append(point.X.ToString("R", System.Globalization.CultureInfo.InvariantCulture)).Append(',').Append(point.Y.ToString("R", System.Globalization.CultureInfo.InvariantCulture)).Append(';');
            text.Append('|');
        }
        return text.ToString();
    }

    private sealed class CountingCffSink : IOfficeCffPathSink {
        internal int MoveCount { get; private set; }
        internal int DrawingOperationCount { get; private set; }

        public void MoveTo(double x, double y) {
            MoveCount++;
            DrawingOperationCount++;
        }

        public void LineTo(double x, double y) => DrawingOperationCount++;

        public void CurveTo(
            double control1X,
            double control1Y,
            double control2X,
            double control2Y,
            double x,
            double y) => DrawingOperationCount++;

        public void CloseContour() {
        }
    }
}
