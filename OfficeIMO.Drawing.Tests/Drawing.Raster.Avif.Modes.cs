using System.Text.Json;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Native entropy streams exercise completed luma neighbors, chroma gating and shared adaptive angle/alpha CDFs.</summary>
public sealed class DrawingAv1ModeTests {
    [Fact]
    public void NativeModeStreamsMatchCompletedNeighborsChromaEligibilityAndSharedAdaptation() {
        using var json = OpenFixture();
        foreach (var c in json.RootElement.GetProperty("cases").EnumerateArray()) {
            int scenario=c.GetProperty("scenario").GetInt32(), units=scenario==9?32:16, extent=scenario==9||scenario==3?32:16;
            var frame=new OfficeAv1StillFrame {MiRows=extent+units,MiCols=extent+units,AllowIntraBlockCopy=scenario==8};
            var sequence=new OfficeAv1StillSequence {Use128Superblock=scenario==9,Monochrome=scenario==1};
            var contexts=new[] {
                new OfficeAv1IntraModeReader(frame,sequence,new OfficeAv1Tile(0,1,0,extent,0,extent),new OfficeRasterDecodeOptions()),
                new OfficeAv1IntraModeReader(frame,sequence,new OfficeAv1Tile(0,1,units,extent+units,units,extent+units),new OfficeRasterDecodeOptions())
            };
            byte[] bytes=Hex(c.GetProperty("hex").GetString()!);
            var streams=Enumerable.Range(0,2).Select(_=>new OfficeAv1SymbolReader(bytes,0,bytes.Length,100000,
                !c.GetProperty("updates").GetBoolean())).ToArray();
            foreach (var state in c.GetProperty("states").EnumerateArray()) {
                int[] expected=state.EnumerateArray().Select(v=>v.GetInt32()).ToArray();
                for (int tile=0;tile<2;tile++) {
                    var block=new OfficeAv1BlockRegion(expected[0]+tile*units,expected[1]+tile*units,expected[2],expected[3]);
                    var actual=contexts[tile].Read(streams[tile],block,Prelude(scenario==2||scenario==3||scenario==10));
                    Assert.True(expected.Skip(4).SequenceEqual(State(actual)), $"scenario={scenario}, leaf={expected[0]},{expected[1]}, tile={tile}");
                    contexts[tile].CompleteBlock();
                }
            }
            foreach (var stream in streams) stream.Finish(); // Entire isolated mode grammar; not a full AV1 tile.
        }
    }

    [Fact]
    public void FrozenFirstLeafModesMatchNativeAfterPartitionAndPreludeSyntax() {
        using var json=OpenFixture();
        foreach (var c in json.RootElement.GetProperty("framePrefixes").EnumerateArray()) {
            byte[] bytes=File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif",c.GetProperty("name").GetString()!+".avif"));
            var options=new OfficeRasterDecodeOptions();
            Assert.True(OfficeAvifContainerReader.TryRead(bytes,options,out var container));
            var item=c.GetProperty("alpha").GetBoolean()?container!.Alpha!:container!.Color;
            Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes,item,options,out var sequence));
            Assert.True(OfficeAv1StillFrameReader.TryRead(bytes,item,sequence!,options,out var frame));
            var tile=frame!.Tiles[0];
            Assert.False(frame.AllowIntraBlockCopy);
            Assert.Equal(c.GetProperty("alpha").GetBoolean(),sequence!.Monochrome);
            Assert.Equal(c.GetProperty("offset").GetInt32(),tile.Offset);
            Assert.Equal(c.GetProperty("length").GetInt32(),tile.Length);
            var stream=new OfficeAv1SymbolReader(bytes,tile.Offset,tile.Length,1000,frame.DisableCdfUpdate);
            var partitions=new OfficeAv1PartitionContext(frame,tile,sequence.Use128Superblock?128:64);
            int pixels=sequence.Use128Superblock?128:64;
            OfficeAv1PartitionLayout layout;
            while ((layout=partitions.Read(stream,0,0,pixels)).Kind==OfficeAv1Partition.Split) pixels/=2;
            Assert.Equal(OfficeAv1Partition.None,layout.Kind);
            Assert.Equal(c.GetProperty("pixels").GetInt32(),pixels);
            var block=layout.Child(0);
            var preludes=new OfficeAv1BlockPreludeReader(frame,sequence,tile,options);
            preludes.BeginSuperblock(0,0);
            var prelude=preludes.Read(stream,block);
            Assert.Equal(c.GetProperty("preludeQ").GetInt32(),prelude.CurrentQIndex);
            var modes=new OfficeAv1IntraModeReader(frame,sequence,tile,options);
            Assert.Equal(c.GetProperty("modes").EnumerateArray().Select(v=>v.GetInt32()),State(modes.Read(stream,block,prelude)));
            // Stop before palette/filter-intra/motion and residual syntax; do not finish or publish this incomplete leaf.
        }
    }

    [Fact]
    public void ModeLifetimeAndResourceRejectionDoNotConsumeEntropy() {
        var frame=new OfficeAv1StillFrame {MiRows=16,MiCols=16};
        var sequence=new OfficeAv1StillSequence();var tile=new OfficeAv1Tile(0,1,0,16,0,16);
        var context=new OfficeAv1IntraModeReader(frame,sequence,tile,new OfficeRasterDecodeOptions());
        Assert.Throws<InvalidOperationException>(()=>context.CompleteBlock());
        var untouched=new OfficeAv1SymbolReader(Hex("20"),0,1,1,true);
        Assert.Throws<FormatException>(()=>context.Read(untouched,new OfficeAv1BlockRegion(0,15,8,8),Prelude(false)));
        Assert.False(untouched.ReadBool());untouched.Finish();
        Assert.Throws<FormatException>(()=>new OfficeAv1IntraModeReader(frame,sequence,tile,
            new OfficeRasterDecodeOptions {RetainedManagedBytes=long.MaxValue}));
        Assert.Throws<FormatException>(()=>new OfficeAv1IntraModeReader(frame,sequence,tile,
            new OfficeRasterDecodeOptions {MaximumDecodedPixels=1}));
        using var json=OpenFixture();
        var c=json.RootElement.GetProperty("cases").EnumerateArray().First(v=>v.GetProperty("scenario").GetInt32()==0);
        byte[] bytes=Hex(c.GetProperty("hex").GetString()!);
        var stream=new OfficeAv1SymbolReader(bytes,0,bytes.Length,100000,true);
        var block=new OfficeAv1BlockRegion(0,0,8,8);
        context.Read(stream,block,Prelude(false));
        Assert.Throws<InvalidOperationException>(()=>context.Read(stream,new OfficeAv1BlockRegion(0,2,8,8),Prelude(false)));
        context.CompleteBlock();
        Assert.Equal(c.GetProperty("states")[1].EnumerateArray().Skip(4).Select(v=>v.GetInt32()),
            State(context.Read(stream,new OfficeAv1BlockRegion(0,2,8,8),Prelude(false))));
    }

    [Fact]
    public void FailedOrCanceledModesCannotPublishOrResumeLeafState() {
        using var cancellation=new CancellationTokenSource();
        var frame=new OfficeAv1StillFrame {MiRows=16,MiCols=16};
        var options=new OfficeRasterDecodeOptions {CancellationToken=cancellation.Token};
        var tile=new OfficeAv1Tile(0,1,0,16,0,16);
        var context=new OfficeAv1IntraModeReader(frame,new OfficeAv1StillSequence(),tile,options);
        var stream=new OfficeAv1SymbolReader(Hex("20"),0,1,1,true);
        Assert.Throws<FormatException>(()=>context.Read(stream,new OfficeAv1BlockRegion(0,0,8,8),Prelude(false)));
        Assert.Throws<FormatException>(()=>context.CompleteBlock());
        Assert.Throws<FormatException>(()=>context.Read(stream,new OfficeAv1BlockRegion(0,0,8,8),Prelude(false)));
        var canceled=new OfficeAv1IntraModeReader(frame,new OfficeAv1StillSequence(),tile,options);
        var untouched=new OfficeAv1SymbolReader(Hex("20"),0,1,1,true);
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(()=>canceled.Read(untouched,new OfficeAv1BlockRegion(0,0,8,8),Prelude(false)));
        Assert.False(untouched.ReadBool());untouched.Finish();
    }

    private static OfficeAv1BlockPrelude Prelude(bool lossless) => new OfficeAv1BlockPrelude(false,0,lossless,-1,lossless?0:120,new int[4]);
    private static int[] State(OfficeAv1IntraModes m) => new[] {m.UseIntraBlockCopy?1:0,m.HasChroma?1:0,m.CflAllowed?1:0,
        (int)m.YMode,(int)m.UvMode,m.AngleDeltaY,m.AngleDeltaUv,m.CflAlphaU,m.CflAlphaV};
    private static JsonDocument OpenFixture() => JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif","mode-reference.json")));
    private static byte[] Hex(string hex) {
        var b=new byte[hex.Length/2];for(int i=0;i<b.Length;i++)b[i]=Convert.ToByte(hex.Substring(i*2,2),16);return b;
    }
}
