using System.Text.Json;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Independent native component streams protect leaf syntax order, neighbor contexts and running state.</summary>
public sealed class DrawingAv1PreludeTests {
    [Fact]
    public void NativeStreamsMatchSegmentsSkipsCdefAndRunningDeltasAcrossSuperblocksAndTiles() {
        using var json = OpenFixture();
        foreach (var c in json.RootElement.GetProperty("cases").EnumerateArray()) {
            int sb = c.GetProperty("superblock").GetInt32(), units = sb / 4;
            int scenario = c.GetProperty("scenario").GetInt32();
            var sequence = new OfficeAv1StillSequence { Use128Superblock = sb == 128, Cdef = true, Monochrome = scenario == 1 };
            var frame = CreateFrame(scenario, units * 3);
            var options = new OfficeRasterDecodeOptions();
            var contexts = new[] {
                new OfficeAv1BlockPreludeReader(frame, sequence, new OfficeAv1Tile(0,1,0,units*2,0,units*2), options),
                new OfficeAv1BlockPreludeReader(frame, sequence, new OfficeAv1Tile(0,1,units,units*3,units,units*3), options)
            };
            byte[] payload = Hex(c.GetProperty("hex").GetString()!);
            var readers = Enumerable.Range(0,2).Select(_ => new OfficeAv1SymbolReader(payload,0,payload.Length,10000,
                !c.GetProperty("updates").GetBoolean())).ToArray();
            int sr = -1, sc = -1;
            foreach (var state in c.GetProperty("states").EnumerateArray()) {
                int[] expected = state.EnumerateArray().Select(n => n.GetInt32()).ToArray();
                int row = expected[0], col = expected[1], nextRow = row / units * units, nextCol = col / units * units;
                for (int tile = 0; tile < 2; tile++) {
                    int offset = tile * units;
                    if (nextRow != sr || nextCol != sc) contexts[tile].BeginSuperblock(nextRow + offset, nextCol + offset);
                    var block = new OfficeAv1BlockRegion(row+offset,col+offset,expected[2],expected[3]);
                    var actual = contexts[tile].Read(readers[tile],block);
                    Assert.True(expected.Skip(4).SequenceEqual(State(actual)), $"sb={sb}, scenario={scenario}, leaf={row},{col}, tile={tile}");
                    contexts[tile].CompleteBlock();
                }
                sr = nextRow; sc = nextCol;
            }
            // Complete component streams, not actual AV1 tiles with omitted prediction/residual syntax.
            foreach (var reader in readers) reader.Finish();
        }
    }

    [Fact]
    public void FrozenAvifFirstLeafPreludeMatchesNativePrefixAfterPartitionDescent() {
        using var json = OpenFixture();
        foreach (var c in json.RootElement.GetProperty("framePrefixes").EnumerateArray()) {
            byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif",c.GetProperty("name").GetString()!+".avif"));
            var options = new OfficeRasterDecodeOptions();
            Assert.True(OfficeAvifContainerReader.TryRead(bytes,options,out var container));
            var item = c.GetProperty("alpha").GetBoolean() ? container!.Alpha! : container!.Color;
            Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes,item,options,out var sequence));
            Assert.True(OfficeAv1StillFrameReader.TryRead(bytes,item,sequence!,options,out var frame));
            Assert.Equal(c.GetProperty("baseQ").GetInt32(),frame!.BaseQIndex);
            Assert.Equal(c.GetProperty("deltaQ").GetBoolean(),frame.DeltaQPresent);
            Assert.Equal(c.GetProperty("cdef").GetBoolean(),sequence!.Cdef);
            var tile = frame.Tiles[0];
            Assert.Equal(c.GetProperty("offset").GetInt32(),tile.Offset);
            Assert.Equal(c.GetProperty("length").GetInt32(),tile.Length);
            var reader = new OfficeAv1SymbolReader(bytes,tile.Offset,tile.Length,100,frame.DisableCdfUpdate);
            var partitions = new OfficeAv1PartitionContext(frame,tile,sequence.Use128Superblock?128:64);
            int pixels = sequence.Use128Superblock?128:64;
            OfficeAv1PartitionLayout layout;
            while ((layout=partitions.Read(reader,0,0,pixels)).Kind==OfficeAv1Partition.Split) pixels/=2;
            Assert.Equal(OfficeAv1Partition.None,layout.Kind);
            Assert.Equal(c.GetProperty("pixels").GetInt32(),pixels);
            var prelude = new OfficeAv1BlockPreludeReader(frame,sequence,tile,options);
            prelude.BeginSuperblock(0,0);
            var actual = prelude.Read(reader,layout.Child(0));
            Assert.Equal(c.GetProperty("skip").GetInt32()!=0,actual.Skip);
            Assert.Equal(c.GetProperty("segment").GetInt32(),actual.SegmentId);
            Assert.Equal(c.GetProperty("cdefIndex").GetInt32(),actual.CdefIndex);
            Assert.Equal(c.GetProperty("q").GetInt32(),actual.CurrentQIndex);
            Assert.Equal(new[] {0,0,0,0},new[] {actual.Filter0,actual.Filter1,actual.Filter2,actual.Filter3});
            // Stop at intra mode syntax. Do not finish the tile or publish this incomplete leaf's neighbors.
        }
    }

    [Fact]
    public void LeafLifetimeGeometryAndMemoryBoundsProtectEntropyAndNeighborState() {
        var frame = CreateFrame(0,32);
        var sequence = new OfficeAv1StillSequence { Cdef = true };
        var tile = new OfficeAv1Tile(0,1,0,32,0,32);
        var context = new OfficeAv1BlockPreludeReader(frame,sequence,tile,new OfficeRasterDecodeOptions());
        var symbols = new OfficeAv1SymbolReader(Hex("20"),0,1,1,true);
        Assert.Throws<FormatException>(()=>context.Read(symbols,new OfficeAv1BlockRegion(0,0,32,32)));
        Assert.Throws<InvalidOperationException>(()=>context.CompleteBlock());
        Assert.Throws<FormatException>(()=>context.BeginSuperblock(1,0));
        context.BeginSuperblock(0,0);
        Assert.Throws<FormatException>(()=>context.Read(symbols,new OfficeAv1BlockRegion(0,15,8,8)));
        Assert.False(symbols.ReadBool()); symbols.Finish(); // Invalid geometry consumed no entropy.
        Assert.Throws<FormatException>(()=>new OfficeAv1BlockPreludeReader(frame,sequence,tile,
            new OfficeRasterDecodeOptions { RetainedManagedBytes = long.MaxValue }));
        Assert.Throws<FormatException>(()=>new OfficeAv1BlockPreludeReader(frame,sequence,tile,
            new OfficeRasterDecodeOptions { MaximumDecodedPixels = 1 }));
        frame.MiCols = frame.MiRows = 16384;
        Assert.Throws<FormatException>(()=>new OfficeAv1BlockPreludeReader(frame,sequence,new OfficeAv1Tile(0,1,0,16384,0,16384),new OfficeRasterDecodeOptions()));

        using var json = OpenFixture();
        var vector = json.RootElement.GetProperty("cases").EnumerateArray().First(c=>c.GetProperty("scenario").GetInt32()==0);
        byte[] bytes = Hex(vector.GetProperty("hex").GetString()!);
        var pending = new OfficeAv1BlockPreludeReader(CreateFrame(0,32),sequence,tile,new OfficeRasterDecodeOptions());
        pending.BeginSuperblock(0,0);
        var stream = new OfficeAv1SymbolReader(bytes,0,bytes.Length,10000,true);
        pending.Read(stream,new OfficeAv1BlockRegion(0,0,32,32));
        Assert.Throws<InvalidOperationException>(()=>pending.Read(stream,new OfficeAv1BlockRegion(0,8,32,32)));
        Assert.Throws<FormatException>(()=>pending.BeginSuperblock(0,16));
        pending.CompleteBlock();
        Assert.Equal(vector.GetProperty("states")[1].EnumerateArray().Skip(4).Select(n=>n.GetInt32()),
            State(pending.Read(stream,new OfficeAv1BlockRegion(0,8,32,32))));
    }

    [Fact]
    public void SegmentSymbolsOutsideTheFramesActiveRangeAreRejectedBeforePublishingNeighbors() {
        using var json = OpenFixture();
        var vector = json.RootElement.GetProperty("cases").EnumerateArray().First(c=>c.GetProperty("scenario").GetInt32()==3);
        var frame = CreateFrame(3,32);
        frame.LastActiveSegmentId = 0;
        var context = new OfficeAv1BlockPreludeReader(frame,new OfficeAv1StillSequence {Cdef=true},
            new OfficeAv1Tile(0,1,0,32,0,32),new OfficeRasterDecodeOptions());
        byte[] bytes = Hex(vector.GetProperty("hex").GetString()!);
        var reader = new OfficeAv1SymbolReader(bytes,0,bytes.Length,10000,true);
        context.BeginSuperblock(0,0);
        Assert.Equal(0,context.Read(reader,new OfficeAv1BlockRegion(0,0,32,32)).SegmentId);
        context.CompleteBlock();
        Assert.Throws<FormatException>(()=>context.Read(reader,new OfficeAv1BlockRegion(0,8,32,32)));
        Assert.Throws<FormatException>(()=>context.CompleteBlock());
    }

    [Fact]
    public void FailedOrCanceledLeafCannotPublishNeighborsOrResumeSyntax() {
        var frame = CreateFrame(0,32);
        using var cancellation = new CancellationTokenSource();
        var context = new OfficeAv1BlockPreludeReader(frame,new OfficeAv1StillSequence {Cdef=true},new OfficeAv1Tile(0,1,0,32,0,32),
            new OfficeRasterDecodeOptions { CancellationToken = cancellation.Token });
        context.BeginSuperblock(0,0);
        var reader = new OfficeAv1SymbolReader(Hex("20"),0,1,1,true);
        Assert.Throws<FormatException>(()=>context.Read(reader,new OfficeAv1BlockRegion(0,0,32,32)));
        Assert.Throws<FormatException>(()=>context.CompleteBlock());
        Assert.Throws<FormatException>(()=>context.BeginSuperblock(0,16));
        var canceled = new OfficeAv1BlockPreludeReader(frame,new OfficeAv1StillSequence(),new OfficeAv1Tile(0,1,0,32,0,32),
            new OfficeRasterDecodeOptions {CancellationToken=cancellation.Token});
        canceled.BeginSuperblock(0,0);
        var untouched = new OfficeAv1SymbolReader(Hex("20"),0,1,1,true);
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(()=>canceled.Read(untouched,new OfficeAv1BlockRegion(0,0,32,32)));
        Assert.False(untouched.ReadBool()); untouched.Finish();
    }

    private static OfficeAv1StillFrame CreateFrame(int scenario, int extent) {
        var f = new OfficeAv1StillFrame { MiRows=extent,MiCols=extent,BaseQIndex=scenario==5?0:120,
            DeltaQPresent=scenario!=5,DeltaQResolution=2,DeltaLoopFilterPresent=scenario!=5&&scenario!=6,DeltaLoopFilterResolution=1,
            DeltaLoopFilterMulti=scenario!=2,CdefBits=2,CodedLossless=scenario==5,AllowIntraBlockCopy=scenario==6,
            SegmentationEnabled=scenario==3||scenario==4,SegmentIdPreSkip=scenario==4,LastActiveSegmentId=7 };
        f.SegmentFeatures[1,6]=true;
        if (f.SegmentationEnabled) f.LosslessSegments[2]=true;
        if (scenario==5) Array.Fill(f.LosslessSegments,true);
        return f;
    }
    private static int[] State(OfficeAv1BlockPrelude b) => new[] {b.Skip?1:0,b.SegmentId,b.Lossless?1:0,b.CdefIndex,b.CurrentQIndex,b.Filter0,b.Filter1,b.Filter2,b.Filter3};
    private static JsonDocument OpenFixture() => JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif","prelude-reference.json")));
    private static byte[] Hex(string hex) {
        var b=new byte[hex.Length/2];for(int i=0;i<b.Length;i++) b[i]=Convert.ToByte(hex.Substring(i*2,2),16);return b;
    }
}
