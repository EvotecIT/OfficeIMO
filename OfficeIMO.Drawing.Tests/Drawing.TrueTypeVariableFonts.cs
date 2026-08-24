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
    public void ItemVariationStoreAllowsNullDataOffsetsAsAbsentSelections() {
        byte[] source = ReadAsset("RobotoFlex.ttf");
        byte[] data = new byte[source.Length + 16];
        Buffer.BlockCopy(source, 0, data, 0, source.Length);
        int store = source.Length;
        WriteUInt16(data, store, 1); // format
        WriteUInt32(data, store + 2, 12); // VariationRegionList offset
        WriteUInt16(data, store + 6, 1); // ItemVariationData count
        WriteUInt32(data, store + 8, 0); // valid NULL ItemVariationData offset

        OfficeOpenTypeReader reader = Assert.IsType<OfficeOpenTypeReader>(OfficeOpenTypeReader.TryCreate(data));
        OfficeFontVariationModel model = OfficeFontVariationModel.Create(
            reader,
            new Dictionary<string, float> { ["wght"] = 900F });
        WriteUInt16(data, store + 12, model.AxisCount);
        WriteUInt16(data, store + 14, 0); // no regions are required by an absent selection

        OfficeOpenTypeItemVariationStore variationStore = OfficeOpenTypeItemVariationStore.Parse(
            reader,
            store,
            data.Length,
            model);

        Assert.Equal(0, variationStore.Evaluate(0, 0));
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
        var text = new StringBuilder();
        foreach (List<OfficePoint> contour in contours) {
            foreach (OfficePoint point in contour) text.Append(point.X.ToString("R", CultureInfo.InvariantCulture)).Append(',').Append(point.Y.ToString("R", CultureInfo.InvariantCulture)).Append(';');
            text.Append('|');
        }
        return text.ToString();
    }
}
