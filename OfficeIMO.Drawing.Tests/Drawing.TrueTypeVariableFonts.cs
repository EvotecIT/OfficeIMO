using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingTrueTypeVariableFontTests {
    [Fact]
    public void TrueTypeVariationAxesProduceDistinctDeterministicOutlinesAndMetrics() {
        byte[] data = ReadAsset("RobotoFlex.ttf");
        OfficeFontFace light = Load(data, new Dictionary<string, float> {
            ["wght"] = 200F,
            ["wdth"] = 75F
        });
        OfficeFontFace black = Load(data, new Dictionary<string, float> {
            ["wght"] = 900F,
            ["wdth"] = 125F
        });

        Assert.False(light.CanEmbedAsStaticPdfFont);
        Assert.False(black.CanEmbedAsStaticPdfFont);
        Assert.NotEqual(light.Program.Fingerprint, black.Program.Fingerprint);
        string lightContours = Serialize(light.Program.GetTextContours("Variable OfficeIMO", 0D, 0D, 24D));
        string blackContours = Serialize(black.Program.GetTextContours("Variable OfficeIMO", 0D, 0D, 24D));
        Assert.NotEmpty(lightContours);
        Assert.NotEmpty(blackContours);
        Assert.NotEqual(lightContours, blackContours);
        Assert.Equal(lightContours, Serialize(light.Program.GetTextContours("Variable OfficeIMO", 0D, 0D, 24D)));
        Assert.NotEqual(light.Program.Measure("Variable OfficeIMO", 24D), black.Program.Measure("Variable OfficeIMO", 24D));
    }

    [Fact]
    public void VariableEmojiUsesFirstPartyGvarOutlines() {
        byte[] data = ReadAsset("NotoEmoji-VariableFont_wght.ttf");
        OfficeFontFace face = Load(data, new Dictionary<string, float> { ["wght"] = 700F }, "Noto Emoji");

        Assert.False(face.CanEmbedAsStaticPdfFont);
        Assert.True(face.Program.HasGlyphs("😀🚀🌍"));
        Assert.NotEmpty(face.Program.GetTextContours("😀🚀🌍", 0D, 0D, 24D));
    }

    [Fact]
    public void VariableAxisSelectionRejectsUnknownAndNonFiniteValues() {
        byte[] data = ReadAsset("RobotoFlex.ttf");
        var unknown = new OfficeFontFaceCollection {
            FontVariationResolver = _ => new Dictionary<string, float> { ["nope"] = 1F }
        };
        Assert.False(unknown.TryAddBounded(
            "Roboto Flex",
            data,
            OfficeFontStyle.Regular,
            OfficeFontUnicodeRangeSet.All,
            8 * 1024 * 1024,
            out _,
            out string? unknownError));
        Assert.Contains("not defined", unknownError, StringComparison.Ordinal);

        var nonFinite = new OfficeFontFaceCollection {
            FontVariationResolver = _ => new Dictionary<string, float> { ["wght"] = float.NaN }
        };
        Assert.False(nonFinite.TryAddBounded(
            "Roboto Flex",
            data,
            OfficeFontStyle.Regular,
            OfficeFontUnicodeRangeSet.All,
            8 * 1024 * 1024,
            out _,
            out string? nonFiniteError));
        Assert.Contains("finite", nonFiniteError, StringComparison.Ordinal);
    }

    [Fact]
    public void VariableAxisSelectionRejectsInvalidAvarSegmentMaps() {
        byte[] data = ReadAsset("RobotoFlex.ttf");
        int avarOffset = FindTableOffset(data, "avar");
        int firstMapCount = ReadUInt16(data, avarOffset + 8);
        Assert.True(firstMapCount >= 3);
        bool changedRequiredOrigin = false;
        for (int index = 0; index < firstMapCount; index++) {
            int entry = avarOffset + 10 + index * 4;
            if (ReadUInt16(data, entry) != 0 || ReadUInt16(data, entry + 2) != 0) continue;
            WriteUInt16(data, entry + 2, 1);
            changedRequiredOrigin = true;
            break;
        }
        Assert.True(changedRequiredOrigin);
        var fonts = new OfficeFontFaceCollection {
            FontVariationResolver = _ => new Dictionary<string, float> { ["wght"] = 700F }
        };

        Assert.False(fonts.TryAddBounded(
            "Roboto Flex",
            data,
            OfficeFontStyle.Regular,
            OfficeFontUnicodeRangeSet.All,
            8 * 1024 * 1024,
            out _,
            out string? error));
        Assert.Contains("avar segment map", error, StringComparison.Ordinal);
    }

    [Fact]
    public void ProviderReceivesTheResolvedVariableFontCoordinates() {
        byte[] data = ReadAsset("RobotoFlex.ttf");
        var provider = new CapturingFontProgramProvider();
        var requested = new Dictionary<string, float> { ["wght"] = 725F, ["wdth"] = 112F };
        var fonts = new OfficeFontFaceCollection {
            FontVariationResolver = _ => requested,
            FontProgramProvider = provider
        };

        Assert.True(fonts.TryAdd("Roboto Flex", data));
        Assert.NotNull(provider.LastRequest);
        Assert.Equal(725F, provider.LastRequest!.VariationCoordinates["wght"]);
        Assert.Equal(112F, provider.LastRequest.VariationCoordinates["wdth"]);
        requested["wght"] = 200F;
        Assert.Equal(725F, provider.LastRequest.VariationCoordinates["wght"]);
    }

    [Fact]
    public void TrueTypeScalarRenderingSkipsIgnorableShapingControls() {
        OfficeFontFace face = Load(
            ReadAsset("RobotoFlex.ttf"),
            new Dictionary<string, float> { ["wght"] = 700F });
        const string visible = "AA";
        const string withControls = "A\u061C\u200D\uFE0FAA";

        Assert.Equal(face.Program.Measure(visible + "A", 24D), face.Program.Measure(withControls, 24D), 6);
        Assert.Equal(
            Serialize(face.Program.GetTextContours(visible + "A", 0D, 0D, 24D)),
            Serialize(face.Program.GetTextContours(withControls, 0D, 0D, 24D)));
    }

    [Fact]
    public void GvarPhantomPointsSupplyAdvanceWidthsWhenHvarIsAbsent() {
        byte[] data = ReadAsset("RobotoFlex.ttf");
        RenameTable(data, "HVAR", "HVAX");
        OfficeFontFace selected = Load(data, new Dictionary<string, float> { ["wght"] = 900F });

        Assert.True(selected.Program.TryGetGlyphMetrics('A', out _, out int advance));
        Assert.Equal(1452, advance);
        Assert.Equal(advance * 24D / selected.Program.UnitsPerEm, selected.Program.Measure("A", 24D), 6);
    }

    [Fact]
    public void VariableFontRegistrationRejectsAnImplicitHvarMapOutsideTheVariationStore() {
        byte[] data = ReadAsset("RobotoFlex.ttf");
        int hvar = FindTableOffset(data, "HVAR");
        Assert.NotEqual(0U, ReadUInt32(data, hvar + 8));
        WriteUInt32(data, hvar + 8, 0U);
        var fonts = new OfficeFontFaceCollection();

        Assert.False(fonts.TryAdd("Malformed HVAR", data));
    }

    [Fact]
    public void VariableFontRegistrationRejectsGvarDataInsideTheOffsetDirectory() {
        byte[] data = ReadAsset("RobotoFlex.ttf");
        int gvar = FindTableOffset(data, "gvar");
        WriteUInt32(data, gvar + 16, 20U);
        var fonts = new OfficeFontFaceCollection();

        Assert.False(fonts.TryAdd("Malformed gvar", data));
    }

    [Fact]
    public void TrueTypeMvarAdjustsSelectedHorizontalLineMetrics() {
        byte[] original = ReadAsset("RobotoFlex.ttf");
        var coordinates = new Dictionary<string, float> { ["wght"] = 1000F };
        OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(original));
        OfficeFontVariationModel model = OfficeFontVariationModel.Create(reader, coordinates);
        byte[] withMvar = ReplaceHvarWithMvar(original, CreateHorizontalMetricsMvar(model));
        OfficeFontFace originalFace = Load(original, coordinates);
        OfficeFontFace mvarFace = Load(withMvar, coordinates);

        double scale = 24D / mvarFace.Program.UnitsPerEm;
        Assert.Equal(originalFace.Program.LineHeight(24D) + 120D * scale, mvarFace.Program.LineHeight(24D), 6);
        Assert.Equal(
            originalFace.Program.LineSpacingRatio + 150D / mvarFace.Program.UnitsPerEm,
            mvarFace.Program.LineSpacingRatio,
            6);
    }

    [Fact]
    public void CompositeUseMyMetricsSuppliesSelectedComponentAdvanceWhenHvarIsAbsent() {
        byte[] data = ReadAsset("RobotoFlex.ttf");
        RenameTable(data, "HVAR", "HVAX");
        OfficeFontFace selected = Load(data, new Dictionary<string, float> { ["wght"] = 900F });

        Assert.True(selected.Program.TryGetGlyphMetrics('A', out _, out int componentAdvance));
        Assert.True(selected.Program.TryGetGlyphMetrics('À', out _, out int compositeAdvance));
        Assert.Equal(componentAdvance, compositeAdvance);
        Assert.Equal(1452, compositeAdvance);
    }

    [Fact]
    public void ItemVariationStoreRejectsNullDataOffsets() {
        byte[] source = ReadAsset("RobotoFlex.ttf");
        byte[] data = new byte[source.Length + 16];
        Buffer.BlockCopy(source, 0, data, 0, source.Length);
        int store = source.Length;
        WriteUInt16(data, store, 1); // format
        WriteUInt32(data, store + 2, 12); // VariationRegionList offset
        WriteUInt16(data, store + 6, 1); // ItemVariationData count
        WriteUInt32(data, store + 8, 0); // ItemVariationData offsets are not nullable.

        OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(data));
        OfficeFontVariationModel model = OfficeFontVariationModel.Create(
            reader,
            new Dictionary<string, float> { ["wght"] = 900F });
        WriteUInt16(data, store + 12, model.AxisCount);
        WriteUInt16(data, store + 14, 0); // no regions are required by an absent selection

        Assert.Throws<InvalidDataException>(() => OfficeOpenTypeItemVariationStore.Parse(
            reader,
            store,
            data.Length,
            model));
    }

    [Theory]
    [InlineData(1, 0)]
    [InlineData(0, 3)]
    public void MvarRejectsOutOfRangeVariationStoreIndices(int outerIndex, int innerIndex) {
        byte[] original = ReadAsset("RobotoFlex.ttf");
        var coordinates = new Dictionary<string, float> { ["wght"] = 1000F };
        OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(original));
        OfficeFontVariationModel model = OfficeFontVariationModel.Create(reader, coordinates);
        byte[] mvar = CreateHorizontalMetricsMvar(model);
        WriteUInt16(mvar, 16, outerIndex);
        WriteUInt16(mvar, 18, innerIndex);
        byte[] malformed = ReplaceHvarWithMvar(original, mvar);
        var fonts = new OfficeFontFaceCollection { FontVariationResolver = _ => coordinates };

        Assert.False(fonts.TryAddBounded(
            "Invalid MVAR",
            malformed,
            OfficeFontStyle.Regular,
            OfficeFontUnicodeRangeSet.All,
            8 * 1024 * 1024,
            out _,
            out string? error));
        Assert.Contains("variation-store", error, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0.5D, 0.25D, 0.75D)]
    [InlineData(-0.5D, 0.75D, 0.5D)]
    [InlineData(-0.5D, 0.25D, 0.5D)]
    public void NonParticipatingVariationAxesHaveNeutralScalar(double start, double peak, double end) {
        Assert.Equal(1D, OfficeOpenTypeVariationRegion.CalculateScalar(0.4D, start, peak, end));
    }

    [Fact]
    public void ParticipatingVariationAxesStillSuppressCoordinatesOutsideTheRegion() {
        Assert.Equal(0D, OfficeOpenTypeVariationRegion.CalculateScalar(-0.5D, 0D, 0.5D, 1D));
    }

    [Fact]
    public void IntermediateGvarTupleIgnoresZeroPeakAxes() {
        Assert.Equal(
            1D,
            OfficeOpenTypeVariationRegion.CalculateTupleScalar(
                coordinate: 0.75D,
                peak: 0D,
                intermediateStart: 0D,
                intermediateEnd: 0D));
        Assert.Equal(
            0.5D,
            OfficeOpenTypeVariationRegion.CalculateTupleScalar(
                coordinate: 0.25D,
                peak: 0.5D,
                intermediateStart: 0D,
                intermediateEnd: 1D));
    }

    [Theory]
    [InlineData(1D, 0.5D)]
    [InlineData(-1D, -0.5D)]
    public void NonIntermediateGvarTupleClampsSameSignCoordinatesBeyondThePeak(double coordinate, double peak) {
        Assert.Equal(
            1D,
            OfficeOpenTypeVariationRegion.CalculateTupleScalar(
                coordinate,
                peak,
                intermediateStart: null,
                intermediateEnd: null));
    }

    [Fact]
    public void OpenTypeReaderRejectsTablesInsideMisalignedOrOverlappingTheDirectory() {
        byte[] original = ReadAsset("RobotoFlex.ttf");
        int firstRecord = FindNonEmptyTableRecords(original).First();
        int secondRecord = FindNonEmptyTableRecords(original).Skip(1).First();
        uint firstOffset = ReadUInt32(original, firstRecord + 8);

        byte[] insideDirectory = (byte[])original.Clone();
        WriteUInt32(insideDirectory, firstRecord + 8, 12U);
        Assert.Null(OfficeOpenTypeReader.TryCreate(insideDirectory));

        byte[] misaligned = (byte[])original.Clone();
        WriteUInt32(misaligned, firstRecord + 8, checked(firstOffset + 1U));
        Assert.Null(OfficeOpenTypeReader.TryCreate(misaligned));

        byte[] overlapping = (byte[])original.Clone();
        WriteUInt32(overlapping, secondRecord + 8, firstOffset);
        Assert.Null(OfficeOpenTypeReader.TryCreate(overlapping));
    }

    [Theory]
    [InlineData(0.5D, 0.25D, 0.75D)]
    [InlineData(-0.5D, 0.75D, 0.5D)]
    [InlineData(-0.5D, 0.25D, 0.5D)]
    public void IntermediateGvarTupleIgnoresInvalidAxisRegions(double start, double peak, double end) {
        Assert.Equal(
            1D,
            OfficeOpenTypeVariationRegion.CalculateTupleScalar(
                coordinate: 0.4D,
                peak,
                intermediateStart: start,
            intermediateEnd: end));
    }

    [Fact]
    public void FontCollectionFaceCanShapeButIsNotDirectlyEmbeddable() {
        int[] probeScalars = "OfficeIMO 0123456789".Select(character => (int)character).ToArray();
        byte[] collection = ManagedTextShapingTestAssets.CreateFontCollection(probeScalars);
        var fonts = new OfficeFontFaceCollection();

        Assert.True(fonts.TryAdd("Collection", collection));

        OfficeFontFace face = Assert.Single(fonts.Faces);
        Assert.True(face.Program.HasGlyphs("OfficeIMO 0123456789"));
        Assert.False(face.CanEmbedAsStaticPdfFont);
    }

    [Fact]
    public void VariableAxisSelectionRejectsFontCollectionsExplicitly() {
        byte[] collection = ManagedTextShapingTestAssets.CreateFontCollection('B');
        var fonts = new OfficeFontFaceCollection {
            FontVariationResolver = _ => new Dictionary<string, float> { ["wght"] = 700F }
        };

        Assert.False(fonts.TryAddBounded(
            "Collection",
            collection,
            OfficeFontStyle.Regular,
            OfficeFontUnicodeRangeSet.All,
            8 * 1024 * 1024,
            out _,
            out string? error));
        Assert.Contains("font collection", error, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Extract", error, StringComparison.Ordinal);
    }

    private static OfficeFontFace Load(
        byte[] data,
        IReadOnlyDictionary<string, float> axes,
        string family = "Roboto Flex") {
        var fonts = new OfficeFontFaceCollection { FontVariationResolver = _ => axes };
        Assert.True(fonts.TryAddBounded(
            family,
            data,
            OfficeFontStyle.Regular,
            OfficeFontUnicodeRangeSet.All,
            8 * 1024 * 1024,
            out _,
            out string? error), error);
        return Assert.Single(fonts.Faces);
    }

    private sealed class CapturingFontProgramProvider : IOfficeFontProgramProvider {
        internal OfficeFontProgramLoadRequest? LastRequest { get; private set; }

        public OfficeFontProgramLoadResult? TryLoad(OfficeFontProgramLoadRequest request) {
            LastRequest = request;
            return null;
        }
    }

    private static byte[] ReadAsset(string name) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", name));

    private static void RenameTable(byte[] data, string oldTag, string newTag) {
        int tableCount = (data[4] << 8) | data[5];
        for (int table = 0; table < tableCount; table++) {
            int offset = 12 + table * 16;
            if (data[offset] != oldTag[0] || data[offset + 1] != oldTag[1]
                || data[offset + 2] != oldTag[2] || data[offset + 3] != oldTag[3]) continue;
            for (int index = 0; index < 4; index++) data[offset + index] = (byte)newTag[index];
            return;
        }
        throw new InvalidOperationException("The test font does not contain table " + oldTag + ".");
    }

    private static int FindTableOffset(byte[] data, string tag) {
        int tableCount = ReadUInt16(data, 4);
        for (int table = 0; table < tableCount; table++) {
            int record = 12 + table * 16;
            if (data[record] != tag[0] || data[record + 1] != tag[1]
                || data[record + 2] != tag[2] || data[record + 3] != tag[3]) continue;
            return checked((int)ReadUInt32(data, record + 8));
        }
        throw new InvalidOperationException("The test font does not contain table " + tag + ".");
    }

    private static IEnumerable<int> FindNonEmptyTableRecords(byte[] data) {
        int tableCount = ReadUInt16(data, 4);
        for (int table = 0; table < tableCount; table++) {
            int record = 12 + table * 16;
            if (ReadUInt32(data, record + 12) > 0) yield return record;
        }
    }

    private static int ReadUInt16(byte[] data, int offset) => (data[offset] << 8) | data[offset + 1];

    private static uint ReadUInt32(byte[] data, int offset) =>
        ((uint)data[offset] << 24) |
        ((uint)data[offset + 1] << 16) |
        ((uint)data[offset + 2] << 8) |
        data[offset + 3];

    private static byte[] CreateHorizontalMetricsMvar(OfficeFontVariationModel model) {
        const int recordCount = 3;
        const int itemStoreOffset = 12 + recordCount * 8;
        const int regionListOffset = 12;
        int itemDataOffset = checked(regionListOffset + 4 + model.AxisCount * 6);
        var table = new byte[checked(itemStoreOffset + itemDataOffset + 14)];
        WriteUInt16(table, 0, 1);
        WriteUInt16(table, 6, 8);
        WriteUInt16(table, 8, recordCount);
        WriteUInt16(table, 10, itemStoreOffset);
        WriteUInt32(table, 12, 0x68617363); // hasc
        WriteUInt16(table, 18, 0);
        WriteUInt32(table, 20, 0x68647363); // hdsc
        WriteUInt16(table, 26, 1);
        WriteUInt32(table, 28, 0x686C6770); // hlgp
        WriteUInt16(table, 34, 2);

        int store = itemStoreOffset;
        WriteUInt16(table, store, 1);
        WriteUInt32(table, store + 2, regionListOffset);
        WriteUInt16(table, store + 6, 1);
        WriteUInt32(table, store + 8, checked((uint)itemDataOffset));
        int region = store + regionListOffset;
        WriteUInt16(table, region, model.AxisCount);
        WriteUInt16(table, region + 2, 1);
        for (int axis = 0; axis < model.AxisCount; axis++) {
            double coordinate = model.NormalizedCoordinates[axis];
            int axisOffset = region + 4 + axis * 6;
            WriteInt16(table, axisOffset, coordinate < 0D ? -0x4000 : 0);
            WriteInt16(table, axisOffset + 2, ToF2Dot14(coordinate));
            WriteInt16(table, axisOffset + 4, coordinate > 0D ? 0x4000 : 0);
        }

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
        throw new InvalidDataException("The TrueType fixture does not contain an HVAR table record.");
    }

    private static int ToF2Dot14(double value) => checked((int)Math.Round(value * 16384D, MidpointRounding.AwayFromZero));

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

    private static string Serialize(IReadOnlyList<List<OfficePoint>> contours) {
        var text = new StringBuilder();
        foreach (List<OfficePoint> contour in contours) {
            foreach (OfficePoint point in contour) text.Append(point.X.ToString("R", CultureInfo.InvariantCulture)).Append(',').Append(point.Y.ToString("R", CultureInfo.InvariantCulture)).Append(';');
            text.Append('|');
        }
        return text.ToString();
    }
}
