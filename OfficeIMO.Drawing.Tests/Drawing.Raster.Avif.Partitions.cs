using System.Text.Json;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Native reference streams protect context selection, edge probability collapse and child decode order.</summary>
public sealed class DrawingAv1PartitionTests {
    [Fact]
    public void NativeStreamsMatchEverySizeContextAndMixedAdaptiveEdgeSequence() {
        using var json = JsonDocument.Parse(File.ReadAllText(FixturePath()));
        foreach (var c in json.RootElement.GetProperty("cases").EnumerateArray()) {
            int pixels = c.GetProperty("pixels").GetInt32(), expectedContext = c.GetProperty("context").GetInt32();
            int operations = c.GetProperty("operations").GetInt32();
            int end = pixels == 8 ? 96 : 64 + pixels / 8;
            var frame = new OfficeAv1StillFrame { MiRows = end, MiCols = end };
            var tile = new OfficeAv1Tile(0, 1, (expectedContext & 1) != 0 ? 0 : 32, end,
                (expectedContext & 2) != 0 ? 0 : 32, end);
            var context = new OfficeAv1PartitionContext(frame, tile, 128);
            byte[] payload = Hex(c.GetProperty("hex").GetString()!);
            var reader = new OfficeAv1SymbolReader(payload, 0, payload.Length, operations, !c.GetProperty("updates").GetBoolean());
            for (int i = 0; i < operations; i++) {
                int mode = pixels == 8 ? 0 : i % 3;
                int row = mode == 1 ? 64 : 32, col = mode == 2 ? 64 : 32;
                SetNeighbors(context, tile, row, col, pixels, expectedContext);
                int symbols = pixels == 8 ? 4 : pixels == 128 ? 8 : 10;
                int symbol = mode == 0 ? (i * 7 + i / 5) % symbols : (i / 3) & 1;
                var expected = mode == 0 ? (OfficeAv1Partition)symbol : symbol != 0 ? OfficeAv1Partition.Split
                    : mode == 1 ? OfficeAv1Partition.Horizontal : OfficeAv1Partition.Vertical;
                var partition = context.Read(reader, row, col, pixels);
                Assert.True(expected == partition.Kind, $"size={pixels}, ctx={expectedContext}, op={i}, mode={mode}");
            }
            // Temporary edge adaptation must not corrupt the full CDF used by the next interior symbol.
            reader.Finish();
        }
    }

    [Fact]
    public void FrozenAvifDescentReachesTheSameFirstLeafAsNativeDecoder() {
        using var json = JsonDocument.Parse(File.ReadAllText(FixturePath()));
        foreach (var c in json.RootElement.GetProperty("framePrefixes").EnumerateArray()) {
            byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Avif", c.GetProperty("name").GetString()! + ".avif"));
            var options = new OfficeRasterDecodeOptions();
            Assert.True(OfficeAvifContainerReader.TryRead(bytes, options, out var container));
            var item = c.GetProperty("alpha").GetBoolean() ? container!.Alpha! : container!.Color;
            Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes, item, options, out var sequence));
            Assert.True(OfficeAv1StillFrameReader.TryRead(bytes, item, sequence!, options, out var frame));
            Assert.Equal(c.GetProperty("miRows").GetInt32(), frame!.MiRows);
            Assert.Equal(c.GetProperty("miCols").GetInt32(), frame.MiCols);
            int superblock = sequence!.Use128Superblock ? 128 : 64;
            Assert.Equal(c.GetProperty("superblockPixels").GetInt32(), superblock);
            var tile = frame.Tiles[0];
            Assert.Equal(c.GetProperty("offset").GetInt32(), tile.Offset);
            Assert.Equal(c.GetProperty("length").GetInt32(), tile.Length);
            var context = new OfficeAv1PartitionContext(frame, tile, superblock);
            var reader = new OfficeAv1SymbolReader(bytes, tile.Offset, tile.Length, 6, frame.DisableCdfUpdate);
            foreach (var expected in c.GetProperty("partitions").EnumerateArray()) {
                var partition = context.Read(reader, tile.MiRowStart, tile.MiColStart, expected.GetProperty("pixels").GetInt32());
                Assert.Equal(expected.GetProperty("kind").GetInt32(), (int)partition.Kind);
            }
            Assert.NotEqual(3, c.GetProperty("partitions").EnumerateArray().Last().GetProperty("kind").GetInt32());
            // Stop before leaf modes/coefficients. Partial tile syntax cannot be passed to Finish().
        }
    }

    [Fact]
    public void ForcedEdgeSplitAndFourPixelLeafConsumeNoEntropy() {
        var frame = new OfficeAv1StillFrame { MiRows = 2, MiCols = 2 };
        var context = new OfficeAv1PartitionContext(frame, new OfficeAv1Tile(0, 1, 0, 2, 0, 2), 64);
        var reader = new OfficeAv1SymbolReader(Hex("20"), 0, 1, 1, true);
        Assert.Equal(OfficeAv1Partition.Split, context.Read(reader, 0, 0, 64).Kind);
        Assert.Equal(OfficeAv1Partition.Split, context.Read(reader, 0, 0, 16).Kind);
        Assert.Equal(OfficeAv1Partition.None, context.Read(reader, 0, 0, 4).Kind);
        Assert.False(reader.ReadBool());
        reader.Finish();
    }

    [Fact]
    public void AdaptivePartitionProbabilitiesAreNotSharedBetweenTiles() {
        using var json = JsonDocument.Parse(File.ReadAllText(FixturePath()));
        var c = json.RootElement.GetProperty("cases").EnumerateArray().First(v =>
            v.GetProperty("pixels").GetInt32() == 64 && v.GetProperty("context").GetInt32() == 0 && v.GetProperty("updates").GetBoolean());
        byte[] payload = Hex(c.GetProperty("hex").GetString()!);
        var frame = new OfficeAv1StillFrame { MiRows = 72, MiCols = 72 };
        var tile = new OfficeAv1Tile(0, 1, 32, 72, 32, 72);
        var contexts = new[] { new OfficeAv1PartitionContext(frame, tile, 128), new OfficeAv1PartitionContext(frame, tile, 128) };
        var readers = new[] { new OfficeAv1SymbolReader(payload, 0, payload.Length, 48, false), new OfficeAv1SymbolReader(payload, 0, payload.Length, 48, false) };
        for (int i = 0; i < 48; i++) {
            int mode = i % 3, row = mode == 1 ? 64 : 32, col = mode == 2 ? 64 : 32;
            var expected = mode == 0 ? (OfficeAv1Partition)((i * 7 + i / 5) % 10)
                : ((i / 3) & 1) != 0 ? OfficeAv1Partition.Split : mode == 1 ? OfficeAv1Partition.Horizontal : OfficeAv1Partition.Vertical;
            for (int j = 0; j < 2; j++) {
                SetNeighbors(contexts[j], tile, row, col, 64, 0);
                Assert.Equal(expected, contexts[j].Read(readers[j], row, col, 64).Kind);
            }
        }
        foreach (var reader in readers) reader.Finish();
    }

    [Fact]
    public void PartitionChildrenFollowDecodeOrderAndTileIndependentDimensions() {
        int[][][] expected = {
            new[] { new[] {0,0,64,64} },
            new[] { new[] {0,0,64,32}, new[] {8,0,64,32} },
            new[] { new[] {0,0,32,64}, new[] {0,8,32,64} },
            new[] { new[] {0,0,32,32}, new[] {0,8,32,32}, new[] {8,0,32,32}, new[] {8,8,32,32} },
            new[] { new[] {0,0,32,32}, new[] {0,8,32,32}, new[] {8,0,64,32} },
            new[] { new[] {0,0,64,32}, new[] {8,0,32,32}, new[] {8,8,32,32} },
            new[] { new[] {0,0,32,32}, new[] {8,0,32,32}, new[] {0,8,32,64} },
            new[] { new[] {0,0,32,64}, new[] {0,8,32,32}, new[] {8,8,32,32} },
            new[] { new[] {0,0,64,16}, new[] {4,0,64,16}, new[] {8,0,64,16}, new[] {12,0,64,16} },
            new[] { new[] {0,0,16,64}, new[] {0,4,16,64}, new[] {0,8,16,64}, new[] {0,12,16,64} }
        };
        for (int kind = 0; kind < expected.Length; kind++) {
            var layout = new OfficeAv1PartitionLayout((OfficeAv1Partition)kind, 32, 64, 64);
            Assert.Equal(expected[kind].Length, layout.Count);
            for (int i = 0; i < layout.Count; i++) {
                var child = layout.Child(i); var e = expected[kind][i];
                Assert.Equal(new[] {32 + e[0], 64 + e[1], e[2], e[3]}, new[] {child.MiRow, child.MiCol, child.Width, child.Height});
            }
        }
        Assert.Throws<FormatException>(() => new OfficeAv1PartitionLayout(OfficeAv1Partition.HorizontalA, 0, 0, 8));
        Assert.Throws<FormatException>(() => new OfficeAv1PartitionLayout(OfficeAv1Partition.Horizontal4, 0, 0, 128));
    }

    [Fact]
    public void NeighborDimensionsCannotCrossAnInteriorTileOrCancellationBoundary() {
        var frame = new OfficeAv1StillFrame { MiRows = 64, MiCols = 64 };
        var tile = new OfficeAv1Tile(0, 1, 0, 32, 0, 32);
        using var cancellation = new CancellationTokenSource();
        var context = new OfficeAv1PartitionContext(frame, tile, 128, cancellation.Token);
        Assert.Throws<FormatException>(() => context.RecordBlock(new OfficeAv1BlockRegion(32, 0, 4, 4)));
        Assert.Throws<FormatException>(() => context.RecordBlock(new OfficeAv1BlockRegion(0, 31, 8, 8)));
        Assert.Throws<FormatException>(() => context.RecordBlock(new OfficeAv1BlockRegion(0, 0, 128, 32)));
        Assert.Throws<FormatException>(() => new OfficeAv1PartitionContext(frame, new OfficeAv1Tile(0, 1, 1, 32, 0, 32), 128));
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => context.RecordBlock(new OfficeAv1BlockRegion(0, 0, 4, 4)));
        var reader = new OfficeAv1SymbolReader(Hex("20"), 0, 1, 1, true);
        Assert.Throws<OperationCanceledException>(() => context.Read(reader, 0, 0, 4));
        Assert.False(reader.ReadBool()); // A canceled partition read did not touch entropy state.
    }

    private static void SetNeighbors(OfficeAv1PartitionContext context, OfficeAv1Tile tile, int row, int col, int pixels, int flags) {
        if (row > tile.MiRowStart) {
            // Width selects the above context; the rectangular neighbor's height must not substitute for it.
            int width = (flags & 1) != 0 ? pixels / 2 : pixels;
            context.RecordBlock(new OfficeAv1BlockRegion(row - pixels / 4, col, width, pixels));
        }
        if (col > tile.MiColStart) {
            // Height selects the left context, independently of the neighbor's width.
            int height = (flags & 2) != 0 ? pixels / 2 : pixels;
            context.RecordBlock(new OfficeAv1BlockRegion(row, col - pixels / 4, pixels, height));
        }
    }
    private static string FixturePath() => Path.Combine(AppContext.BaseDirectory, "TestAssets", "Avif", "partition-reference.json");
    private static byte[] Hex(string hex) {
        var result = new byte[hex.Length / 2];
        for (int i = 0; i < result.Length; i++) result[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
        return result;
    }
}
