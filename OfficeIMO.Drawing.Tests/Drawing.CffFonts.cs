using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingCffFontTests {
    [Fact]
    public void CffStructuralValidationRequiresNonOverlappingCharStringsIndexes() {
        byte[] validCff1 = {
            1, 0, 4, 1,
            0, 1, 1, 1, 2, (byte)'A',
            0, 1, 1, 1, 3, 160, 17,
            0, 0,
            0, 0,
            0, 1, 1, 1, 2, 14
        };
        byte[] validCff2 = {
            2, 0, 5, 0, 2,
            150, 17,
            0, 0, 0, 0,
            0, 0, 0, 1, 1, 1, 2, 14
        };
        byte[] overlappingCff2 = {
            2, 0, 5, 0, 2,
            148, 17,
            0, 0,
            0, 0, 0, 1, 1, 1, 2, 14
        };

        Assert.True(OfficeCffFontData.IsStructurallyValidProgram(validCff1, isCff2: false));
        Assert.True(OfficeCffFontData.IsStructurallyValidProgram(validCff2, isCff2: true));
        Assert.False(OfficeCffFontData.IsStructurallyValidProgram(overlappingCff2, isCff2: true));
    }

    [Fact]
    public void OfficeFontFaceCollection_CountsBuiltInCffReaderShapingAndFaceBuffers() {
        byte[] data = ReadAsset("SourceSansPro-Regular.otf");
        var fonts = new OfficeFontFaceCollection();

        Assert.False(fonts.TryAddBounded(
            "Bounded CFF",
            data,
            OfficeFontStyle.Regular,
            OfficeFontUnicodeRangeSet.All,
            maximumDecodedBytes: data.Length * 3 - 1,
            out int rejectedBytes,
            out string? rejectedError));
        Assert.Equal(0, rejectedBytes);
        Assert.Contains("limit", rejectedError, StringComparison.OrdinalIgnoreCase);

        Assert.True(fonts.TryAddBounded(
            "Bounded CFF",
            data,
            OfficeFontStyle.Regular,
            OfficeFontUnicodeRangeSet.All,
            maximumDecodedBytes: data.Length * 3,
            out int acceptedBytes,
            out string? acceptedError), acceptedError);
        Assert.Equal(data.Length * 3, acceptedBytes);
        Assert.Single(fonts.Faces);
    }

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
    public void CffScalarRenderingSkipsIgnorableShapingControls() {
        byte[] data = ReadAsset("SourceSansPro-Regular.otf");
        OfficeFontFace face = Assert.Single(new OfficeFontFaceCollection().Add("Source Sans Pro CFF", data).Faces);
        const string visible = "AAA";
        const string withControls = "A\u061C\u200D\uFE0FAA";

        Assert.Equal(face.Program.Measure(visible, 24D), face.Program.Measure(withControls, 24D), 6);
        Assert.Equal(
            Serialize(face.Program.GetTextContours(visible, 0D, 0D, 24D)),
            Serialize(face.Program.GetTextContours(withControls, 0D, 0D, 24D)));
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

    [Fact]
    public void CffTextRunSharesOneCharStringOperationBudgetAcrossGlyphs() {
        byte[] data = ReadAsset("SourceSansPro-Regular.otf");
        OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(data));
        OfficeCffFontData cff = OfficeCffFontData.Parse(reader, OfficeFontVariationModel.None);
        var budget = new OfficeCffOperationBudget(1);
        var sink = new CountingCffSink();
        var endChar = new OfficeCffFontData.CffSlice(new byte[] { 14 }, 0, 1);

        new OfficeType2CharStringInterpreter(cff, 0, sink, CancellationToken.None, budget).Render(endChar);
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            new OfficeType2CharStringInterpreter(cff, 0, sink, CancellationToken.None, budget).Render(endChar));

        Assert.Contains("operation budget", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void CffOutlinedRunsCanShareOneOperationBudgetAcrossCalls() {
        byte[] data = ReadAsset("SourceSansPro-Regular.otf");
        OfficeFontFace face = Assert.Single(new OfficeFontFaceCollection().Add("Source Sans", data).Faces);
        var cff = Assert.IsAssignableFrom<IOfficeCffBoundedFontProgram>(face.Program);
        var measuringBudget = new OfficeCffOperationBudget();

        Assert.NotEmpty(cff.GetTextContoursBounded(
            "A", 0D, 0D, 24D, 100_000, CancellationToken.None, measuringBudget));
        int operationsForOneRun = 1_000_000 - measuringBudget.RemainingOperations;
        Assert.True(operationsForOneRun > 0);
        var sharedBudget = new OfficeCffOperationBudget(operationsForOneRun);

        Assert.NotEmpty(cff.GetTextContoursBounded(
            "A", 0D, 0D, 24D, 100_000, CancellationToken.None, sharedBudget));
        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            cff.GetTextContoursBounded(
                "A", 0D, 0D, 24D, 100_000, CancellationToken.None, sharedBudget));

        Assert.Contains("operation budget", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void CffRandomOperatorProducesADeterministicSequenceWithinTheUnitRange() {
        byte[] data = ReadAsset("SourceSansPro-Regular.otf");
        OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(data));
        OfficeCffFontData cff = OfficeCffFontData.Parse(reader, OfficeFontVariationModel.None);
        var program = new OfficeCffFontData.CffSlice(new byte[] {
            12, 23, // random x
            12, 23, // random y
            21,     // rmoveto
            149, 139, 5, // 10 0 rlineto
            14
        }, 0, 9);
        var first = new CountingCffSink();
        var second = new CountingCffSink();

        new OfficeType2CharStringInterpreter(cff, 0, first, CancellationToken.None).Render(program);
        new OfficeType2CharStringInterpreter(cff, 0, second, CancellationToken.None).Render(program);

        Assert.InRange(first.LastMoveX, double.Epsilon, 1D);
        Assert.InRange(first.LastMoveY, double.Epsilon, 1D);
        Assert.NotEqual(first.LastMoveX, first.LastMoveY);
        Assert.Equal(first.LastMoveX, second.LastMoveX);
        Assert.Equal(first.LastMoveY, second.LastMoveY);
    }

    [Theory]
    [InlineData(true, 0D, 12D)]
    [InlineData(false, 12D, 0D)]
    public void CffFlex1AppliesTheFinalDeltaToTheMinorAxis(
        bool horizontallyDominant,
        double expectedX,
        double expectedY) {
        byte[] data = ReadAsset("SourceSansPro-Regular.otf");
        OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(data));
        OfficeCffFontData cff = OfficeCffFontData.Parse(reader, OfficeFontVariationModel.None);
        int major = 10;
        int minor = 1;
        var bytes = new List<byte>();
        for (int index = 0; index < 5; index++) {
            bytes.Add(checked((byte)(139 + (horizontallyDominant ? major : minor))));
            bytes.Add(checked((byte)(139 + (horizontallyDominant ? minor : major))));
        }
        bytes.Add(146); // Final minor-axis delta = 7.
        bytes.Add(12);
        bytes.Add(37); // flex1
        bytes.Add(14);
        var sink = new CountingCffSink();

        new OfficeType2CharStringInterpreter(cff, 0, sink, CancellationToken.None).Render(
            new OfficeCffFontData.CffSlice(bytes.ToArray(), 0, bytes.Count));

        Assert.Equal(expectedX, sink.LastCurveX, 6);
        Assert.Equal(expectedY, sink.LastCurveY, 6);
    }

    [Fact]
    public void CffRejectsCoordinateOverflowBeforeCallingThePathSink() {
        byte[] data = ReadAsset("SourceSansPro-Regular.otf");
        OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(data));
        OfficeCffFontData cff = OfficeCffFontData.Parse(reader, OfficeFontVariationModel.None);
        var bytes = new List<byte> { 139, 139, 21 }; // Consume width and open a contour at 0,0.
        AppendHugeFiniteOperand(bytes);
        bytes.Add(139);
        bytes.Add(5);
        AppendHugeFiniteOperand(bytes);
        bytes.Add(139);
        bytes.Add(5);
        bytes.Add(14);
        var sink = new CountingCffSink();
        var program = new OfficeCffFontData.CffSlice(bytes.ToArray(), 0, bytes.Count);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            new OfficeType2CharStringInterpreter(cff, 0, sink, CancellationToken.None).Render(program));

        Assert.Contains("not finite", exception.Message, StringComparison.Ordinal);
        Assert.Equal(2, sink.DrawingOperationCount);
    }

    [Fact]
    public void CffBoundedContoursRejectNonFiniteTransformedGeometry() {
        byte[] data = ReadAsset("SourceSansPro-Regular.otf");
        OfficeOpenTypeCffFont font = Assert.IsType<OfficeOpenTypeCffFont>(
            OfficeOpenTypeCffFont.TryLoad(data, null, out string? error));
        Assert.Null(error);

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            font.GetTextContoursBounded("A", double.MaxValue, 0D, double.MaxValue, 10_000, CancellationToken.None));

        Assert.Contains("not finite", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Cff2MvarAdjustsSelectedHorizontalLineMetrics() {
        byte[] original = ReadAsset("AdobeVFPrototype-Subset.otf");
        byte[] withMvar = ReplaceHvarWithMvar(original, CreateHorizontalMetricsMvar());
        var originalFonts = new OfficeFontFaceCollection {
            FontVariationResolver = _ => new Dictionary<string, float> { ["wght"] = 0F, ["xxxx"] = 0F }
        };
        var mvarFonts = new OfficeFontFaceCollection {
            FontVariationResolver = _ => new Dictionary<string, float> { ["wght"] = 0F, ["xxxx"] = 0F }
        };
        Assert.True(originalFonts.TryAddBounded(
            "Original CFF2", original, OfficeFontStyle.Regular, OfficeFontUnicodeRangeSet.All,
            16 * 1024 * 1024, out _, out string? originalError), originalError);
        Assert.True(mvarFonts.TryAddBounded(
            "MVAR CFF2", withMvar, OfficeFontStyle.Regular, OfficeFontUnicodeRangeSet.All,
            16 * 1024 * 1024, out _, out string? mvarError), mvarError);

        OfficeFontFace originalFace = Assert.Single(originalFonts.Faces);
        OfficeFontFace mvarFace = Assert.Single(mvarFonts.Faces);
        double scale = 24D / mvarFace.Program.UnitsPerEm;
        Assert.Equal(originalFace.Program.LineHeight(24D) + 120D * scale, mvarFace.Program.LineHeight(24D), 6);
        Assert.Equal(originalFace.Program.LineSpacingRatio + 150D / mvarFace.Program.UnitsPerEm, mvarFace.Program.LineSpacingRatio, 6);
    }

    [Fact]
    public void MvarRejectsUnsortedValueRecords() {
        byte[] source = ReadAsset("AdobeVFPrototype-Subset.otf");
        byte[] table = CreateHorizontalMetricsMvar();
        WriteUInt32(table, 12, 0x68647363); // hdsc before hasc is not sorted.
        WriteUInt32(table, 20, 0x68617363);
        byte[] data = ReplaceHvarWithMvar(source, table);
        OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(data));
        OfficeFontVariationModel model = OfficeFontVariationModel.Create(
            reader,
            new Dictionary<string, float> { ["wght"] = 0F, ["xxxx"] = 0F });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            OfficeOpenTypeMvarMetrics.TryParse(reader, model));

        Assert.Contains("strictly tag-sorted", exception.Message, StringComparison.Ordinal);
    }

    private static byte[] ReadAsset(string name) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", name));

    private static void AppendHugeFiniteOperand(List<byte> bytes) {
        bytes.AddRange(new byte[] { 28, 0x7F, 0xFF }); // 32767
        for (int index = 0; index < 6; index++) {
            bytes.AddRange(new byte[] { 12, 27, 12, 24 }); // dup, mul
        }
        for (int index = 0; index < 4; index++) {
            bytes.AddRange(new byte[] { 28, 0x7F, 0xFF, 12, 24 });
        }
        bytes.AddRange(new byte[] { 148, 12, 24 }); // multiply by 9; the result remains finite.
    }

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

    private static void WriteInt16(byte[] data, int offset, int value) => WriteUInt16(data, offset, unchecked((ushort)value));

    private static byte[] CreateHorizontalMetricsMvar() {
        const int recordCount = 3;
        const int itemStoreOffset = 12 + recordCount * 8;
        const int regionListOffset = 12;
        const int itemDataOffset = 28;
        var table = new byte[itemStoreOffset + 42];
        WriteUInt16(table, 0, 1); // major version
        WriteUInt16(table, 2, 0); // minor version
        WriteUInt16(table, 4, 0); // reserved
        WriteUInt16(table, 6, 8); // value record size
        WriteUInt16(table, 8, recordCount);
        WriteUInt16(table, 10, itemStoreOffset);
        WriteUInt32(table, 12, 0x68617363); // hasc
        WriteUInt16(table, 16, 0);
        WriteUInt16(table, 18, 0);
        WriteUInt32(table, 20, 0x68647363); // hdsc
        WriteUInt16(table, 24, 0);
        WriteUInt16(table, 26, 1);
        WriteUInt32(table, 28, 0x686C6770); // hlgp
        WriteUInt16(table, 32, 0);
        WriteUInt16(table, 34, 2);

        int store = itemStoreOffset;
        WriteUInt16(table, store, 1);
        WriteUInt32(table, store + 2, regionListOffset);
        WriteUInt16(table, store + 6, 1);
        WriteUInt32(table, store + 8, itemDataOffset);
        int region = store + regionListOffset;
        WriteUInt16(table, region, 2); // fixture axis count
        WriteUInt16(table, region + 2, 1);
        WriteInt16(table, region + 4, -0x4000);
        WriteInt16(table, region + 6, -0x4000);
        WriteInt16(table, region + 8, 0);
        WriteInt16(table, region + 10, 0);
        WriteInt16(table, region + 12, 0);
        WriteInt16(table, region + 14, 0);
        int itemData = store + itemDataOffset;
        WriteUInt16(table, itemData, 3);
        WriteUInt16(table, itemData + 2, 1);
        WriteUInt16(table, itemData + 4, 1);
        WriteUInt16(table, itemData + 6, 0);
        WriteInt16(table, itemData + 8, 100);
        WriteInt16(table, itemData + 10, -20);
        WriteInt16(table, itemData + 12, 30);
        return table;
    }

    private static byte[] ReplaceHvarWithMvar(byte[] source, byte[] table) {
        int tableOffset = checked((source.Length + 3) & ~3);
        var data = new byte[checked(tableOffset + table.Length)];
        Buffer.BlockCopy(source, 0, data, 0, source.Length);
        Buffer.BlockCopy(table, 0, data, tableOffset, table.Length);
        int tableCount = (data[4] << 8) | data[5];
        for (int index = 0; index < tableCount; index++) {
            int record = 12 + index * 16;
            if (data[record] != (byte)'H' || data[record + 1] != (byte)'V'
                || data[record + 2] != (byte)'A' || data[record + 3] != (byte)'R') continue;
            data[record] = (byte)'M';
            WriteUInt32(data, record + 8, checked((uint)tableOffset));
            WriteUInt32(data, record + 12, checked((uint)table.Length));
            return data;
        }
        throw new InvalidDataException("The CFF2 fixture does not contain an HVAR table record.");
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
        internal double LastMoveX { get; private set; }
        internal double LastMoveY { get; private set; }
        internal double LastCurveX { get; private set; }
        internal double LastCurveY { get; private set; }

        public void MoveTo(double x, double y) {
            MoveCount++;
            DrawingOperationCount++;
            LastMoveX = x;
            LastMoveY = y;
        }

        public void LineTo(double x, double y) => DrawingOperationCount++;

        public void CurveTo(
            double control1X,
            double control1Y,
            double control2X,
            double control2Y,
            double x,
            double y) {
            DrawingOperationCount++;
            LastCurveX = x;
            LastCurveY = y;
        }

        public void CloseContour() {
        }
    }
}
