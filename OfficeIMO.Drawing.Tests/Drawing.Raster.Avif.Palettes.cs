using System.Text.Json;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Native component streams protect cache boundaries, raw colors, filter gates and every reachable map context.</summary>
public sealed class DrawingAv1PaletteTests {
    [Theory]
    [InlineData(8)]
    [InlineData(10)]
    public void NativePaletteStreamsMatchColorsFiltersAndPaddedMapsAcrossIndependentTiles(int bitDepth) {
        using var json = OpenFixture(bitDepth);
        foreach (var c in json.RootElement.GetProperty("cases").EnumerateArray()) {
            int scenario = c.GetProperty("scenario").GetInt32();
            int width = c.GetProperty("states")[0].GetProperty("block")[2].GetInt32();
            int height = c.GetProperty("states")[0].GetProperty("block")[3].GetInt32();
            int extent = scenario % 20 == 18 ? 18 : width == 128 || width == 64 || height == 64 || scenario % 20 == 19 ? 32 : 16;
            int units = width == 128 ? 32 : 16;
            var contexts = Enumerable.Range(0, 2).Select(tile => new OfficeAv1PaletteReader(
                new OfficeAv1StillFrame { BitDepth = bitDepth, MiRows = extent + tile * units, MiCols = extent + tile * units, AllowScreenContentTools = scenario != 21 },
                new OfficeAv1StillSequence { BitDepth = bitDepth, Use128Superblock = width == 128, FilterIntra = scenario != 22, Monochrome = scenario == 20 },
                new OfficeAv1Tile(0, 1, tile * units, tile * units + extent, tile * units, tile * units + extent), new OfficeRasterDecodeOptions())).ToArray();
            byte[] bytes = Hex(c.GetProperty("hex").GetString()!);
            var streams = Enumerable.Range(0, 2).Select(_ => new OfficeAv1SymbolReader(bytes, 0, bytes.Length, 1000000, !c.GetProperty("updates").GetBoolean())).ToArray();
            foreach (var state in c.GetProperty("states").EnumerateArray()) {
                int[] b = state.GetProperty("block").EnumerateArray().Select(v => v.GetInt32()).ToArray();
                for (int tile = 0; tile < 2; tile++) {
                    var block = new OfficeAv1BlockRegion(b[0] + tile * units, b[1] + tile * units, b[2], b[3]);
                    AssertPalette(state, contexts[tile].Read(streams[tile], block, Modes(state.GetProperty("modes"))));
                    contexts[tile].CompleteBlock();
                }
            }
            foreach (var stream in streams) stream.Finish(); // Complete isolated component grammar, not a real AV1 tile.
        }
    }

    [Fact]
    public void FrozenFirstLeavesReachTransformBoundaryWithNativePaletteAndFilterResults() {
        using var json = OpenFixture();
        foreach (var c in json.RootElement.GetProperty("framePrefixes").EnumerateArray()) {
            byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Avif", c.GetProperty("name").GetString()! + ".avif"));
            var options = new OfficeRasterDecodeOptions();
            Assert.True(OfficeAvifContainerReader.TryRead(bytes, options, out var container));
            var item = c.GetProperty("alpha").GetBoolean() ? container!.Alpha! : container!.Color;
            Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes, item, options, out var sequence));
            Assert.True(OfficeAv1StillFrameReader.TryRead(bytes, item, sequence!, options, out var frame));
            Assert.Equal(c.GetProperty("screen").GetBoolean(), frame!.AllowScreenContentTools);
            Assert.True(sequence!.FilterIntra); Assert.False(frame.AllowIntraBlockCopy);
            var tile = frame.Tiles[0];
            var symbols = new OfficeAv1SymbolReader(bytes, tile.Offset, tile.Length, 100000, frame.DisableCdfUpdate);
            var partitions = new OfficeAv1PartitionContext(frame, tile, sequence.Use128Superblock ? 128 : 64);
            int pixels = sequence.Use128Superblock ? 128 : 64;
            OfficeAv1PartitionLayout layout;
            while ((layout = partitions.Read(symbols, 0, 0, pixels)).Kind == OfficeAv1Partition.Split) pixels /= 2;
            Assert.Equal(OfficeAv1Partition.None, layout.Kind); Assert.Equal(c.GetProperty("pixels").GetInt32(), pixels);
            var block = layout.Child(0);
            var preludes = new OfficeAv1BlockPreludeReader(frame, sequence, tile, options); preludes.BeginSuperblock(0, 0);
            var prelude = preludes.Read(symbols, block);
            Assert.Equal(c.GetProperty("preludeQ").GetInt32(), prelude.CurrentQIndex);
            var modes = new OfficeAv1IntraModeReader(frame, sequence, tile, options).Read(symbols, block, prelude);
            AssertPalette(c, new OfficeAv1PaletteReader(frame, sequence, tile, options).Read(symbols, block, modes));
            // Transform sizes/residuals remain: no tile termination or neighbor publication here.
        }
    }

    [Fact]
    public void PaletteLifetimeGeometryAndResourceLimitsRejectWithoutPublishingOrConsuming() {
        var frame = new OfficeAv1StillFrame { MiRows = 16, MiCols = 16, AllowScreenContentTools = true };
        var sequence = new OfficeAv1StillSequence { FilterIntra = true }; var tile = new OfficeAv1Tile(0, 1, 0, 16, 0, 16);
        var context = new OfficeAv1PaletteReader(frame, sequence, tile, new OfficeRasterDecodeOptions());
        Assert.Throws<InvalidOperationException>(() => context.CompleteBlock());
        var untouched = new OfficeAv1SymbolReader(Hex("20"), 0, 1, 1, true);
        var dc = new OfficeAv1IntraModes(false, true, true, OfficeAv1IntraMode.Dc, OfficeAv1IntraMode.Dc, 0, 0, 0, 0);
        Assert.Throws<FormatException>(() => context.Read(untouched, new OfficeAv1BlockRegion(0, 15, 8, 8), dc));
        Assert.False(untouched.ReadBool()); untouched.Finish();
        Assert.Throws<FormatException>(() => new OfficeAv1PaletteReader(frame, sequence, tile, new OfficeRasterDecodeOptions { RetainedManagedBytes = long.MaxValue }));
        Assert.Throws<FormatException>(() => new OfficeAv1PaletteReader(frame, sequence, tile, new OfficeRasterDecodeOptions { MaximumDecodedPixels = 1 }));
        using var json = OpenFixture(); var c = json.RootElement.GetProperty("cases")[0];
        byte[] bytes = Hex(c.GetProperty("hex").GetString()!);
        var symbols = new OfficeAv1SymbolReader(bytes, 0, bytes.Length, 100000, true);
        var block = new OfficeAv1BlockRegion(0, 0, 8, 8);
        var result = context.Read(symbols, block, dc);
        Assert.Throws<InvalidOperationException>(() => context.Read(symbols, block, dc));
        context.CompleteBlock(); AssertPalette(c.GetProperty("states")[0], result);
    }

    [Theory]
    [InlineData(8)]
    [InlineData(10)]
    public void FailedAndCancelledPaletteContextsCannotPublishPartialCaches(int bitDepth) {
        using var json = OpenFixture(bitDepth); byte[] bytes = Hex(json.RootElement.GetProperty("cases")[0].GetProperty("hex").GetString()!);
        var frame = new OfficeAv1StillFrame { BitDepth = bitDepth, MiRows = 16, MiCols = 16, AllowScreenContentTools = true };
        var sequence = new OfficeAv1StillSequence { BitDepth = bitDepth, FilterIntra = true }; var tile = new OfficeAv1Tile(0, 1, 0, 16, 0, 16);
        var block = new OfficeAv1BlockRegion(0, 0, 8, 8);
        var modes = new OfficeAv1IntraModes(false, true, true, OfficeAv1IntraMode.Dc, OfficeAv1IntraMode.Dc, 0, 0, 0, 0);
        var context = new OfficeAv1PaletteReader(frame, sequence, tile, new OfficeRasterDecodeOptions());
        var exhausted = new OfficeAv1SymbolReader(bytes, 0, bytes.Length, 1, true);
        Assert.Throws<FormatException>(() => context.Read(exhausted, block, modes));
        Assert.Throws<FormatException>(() => context.CompleteBlock());
        Assert.Throws<FormatException>(() => context.Read(exhausted, block, modes));
        using var cancellation = new CancellationTokenSource();
        var cancelled = new OfficeAv1PaletteReader(frame, sequence, tile, new OfficeRasterDecodeOptions { CancellationToken = cancellation.Token });
        var untouched = new OfficeAv1SymbolReader(Hex("20"), 0, 1, 1, true); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => cancelled.Read(untouched, block, modes));
        Assert.Throws<OperationCanceledException>(() => cancelled.CompleteBlock());
        Assert.False(untouched.ReadBool()); untouched.Finish();
    }

    [Fact]
    public void PaletteDepthIsValidatedAndSnapshottedBeforeReadingSymbols() {
        var frame = new OfficeAv1StillFrame { BitDepth = 10, MiRows = 16, MiCols = 16, AllowScreenContentTools = true };
        var sequence = new OfficeAv1StillSequence { BitDepth = 8, FilterIntra = true };
        var tile = new OfficeAv1Tile(0, 1, 0, 16, 0, 16);
        var options = new OfficeRasterDecodeOptions();
        Assert.Throws<FormatException>(() => new OfficeAv1PaletteReader(frame, sequence, tile, options));
        frame.BitDepth = sequence.BitDepth = 12;
        Assert.Throws<FormatException>(() => new OfficeAv1PaletteReader(frame, sequence, tile, options));
        frame.BitDepth = sequence.BitDepth = 10;
        var context = new OfficeAv1PaletteReader(frame, sequence, tile, options);
        // Parsed metadata can be reused by its caller; the tile keeps its chosen depth.
        frame.BitDepth = sequence.BitDepth = 8;
        using var fixture = OpenFixture(10);
        var c = fixture.RootElement.GetProperty("cases")[0];
        byte[] bytes = Hex(c.GetProperty("hex").GetString()!);
        var state = c.GetProperty("states")[0];
        var symbols = new OfficeAv1SymbolReader(bytes, 0, bytes.Length, 100000, true);
        AssertPalette(state, context.Read(symbols, new OfficeAv1BlockRegion(0, 0, 8, 8), Modes(state.GetProperty("modes"))));
        context.CompleteBlock();
    }

    private static void AssertPalette(JsonElement expected, OfficeAv1Palette actual) {
        Assert.Equal(expected.GetProperty("filter").GetInt32(), actual.FilterMode);
        var colors = new[] { Colors(expected.GetProperty("y")), Colors(expected.GetProperty("u")), Colors(expected.GetProperty("v")) };
        Assert.Equal(colors[0].Length, actual.SizeY); Assert.Equal(colors[1].Length, actual.SizeUv);
        for (int plane = 0; plane < 3; plane++) Assert.Equal(colors[plane], Enumerable.Range(0, colors[plane].Length).Select(i => actual.Color(plane, i)).ToArray());
        foreach (bool chroma in new[] { false, true }) {
            byte[] map = Hex(expected.GetProperty(chroma ? "mapUv" : "mapY").GetString()!);
            int width = chroma ? actual.ChromaWidth : actual.Width;
            Assert.Equal(map, Enumerable.Range(0, map.Length).Select(i => actual.Index(chroma, i / width, i % width)).ToArray());
        }
    }
    private static OfficeAv1IntraModes Modes(JsonElement element) {
        int[] v = element.EnumerateArray().Select(x => x.GetInt32()).ToArray();
        return new OfficeAv1IntraModes(v[0] != 0, v[1] != 0, v[2] != 0, (OfficeAv1IntraMode)v[3], (OfficeAv1IntraMode)v[4], v[5], v[6], v[7], v[8]);
    }
    private static ushort[] Colors(JsonElement colors) => colors.ValueKind == JsonValueKind.String
        ? Hex(colors.GetString()!).Select(x => (ushort)x).ToArray()
        : colors.EnumerateArray().Select(x => x.GetUInt16()).ToArray();
    private static JsonDocument OpenFixture(int bitDepth = 8) => JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Avif", bitDepth == 8 ? "palette-reference.json" : "palette-main10-reference.json")));
    private static byte[] Hex(string text) => Enumerable.Range(0, text.Length / 2).Select(i => Convert.ToByte(text.Substring(i * 2, 2), 16)).ToArray();
}
