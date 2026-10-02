using System.IO.Compression;
using System.Text.Json;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Native entropy streams protect copy predictor ordering, integer grammar and tile-source safety.</summary>
public sealed class DrawingAv1CopyMotionTests {
    [Fact]
    public void NativeCopyStreamsMatchCompletedSpatialPredictorsAndIndependentTileAdaptation() {
        using var json=OpenFixture();
        foreach (var c in json.RootElement.GetProperty("cases").EnumerateArray()) {
            int rows=c.GetProperty("rows").GetInt32(),cols=c.GetProperty("cols").GetInt32(),start=c.GetProperty("start").GetInt32();
            var frame=new OfficeAv1StillFrame {MiRows=rows,MiCols=cols,AllowIntraBlockCopy=true};
            var sequence=new OfficeAv1StillSequence {Use128Superblock=c.GetProperty("sb").GetInt32()==128,Monochrome=c.GetProperty("variant").GetInt32()==2};
            var tile=new OfficeAv1Tile(0,1,start,rows,start,cols);
            var contexts=Enumerable.Range(0,2).Select(_=>new OfficeAv1CopyMotionReader(frame,sequence,tile,new OfficeRasterDecodeOptions())).ToArray();
            byte[] bytes=Hex(c.GetProperty("hex").GetString()!);
            var streams=Enumerable.Range(0,2).Select(_=>new OfficeAv1SymbolReader(bytes,0,bytes.Length,1000000,!c.GetProperty("updates").GetBoolean())).ToArray();
            foreach (var state in c.GetProperty("states").EnumerateArray()) {
                int[] s=state.EnumerateArray().Select(v=>v.GetInt32()).ToArray();
                var block=new OfficeAv1BlockRegion(s[0],s[1],s[2],s[3]);
                for (int i=0;i<2;i++) {
                    var actual=contexts[i].Read(streams[i],block,Modes(s[4]!=0,s[5]!=0));
                    Assert.True(actual.Row==s[8] && actual.Col==s[9],$"sb={c.GetProperty("sb")}, variant={c.GetProperty("variant")}, block={s[0]},{s[1]}:{s[2]}x{s[3]}, predictor={s[6]},{s[7]}");
                    contexts[i].CompleteBlock();
                }
            }
            foreach (var stream in streams) stream.Finish(); // Entire isolated copy grammar, not full AV1 tile syntax.
        }
    }

    [Fact]
    public void AllIntegerDifferenceClassesRejectUnavailableCopySourcesWithoutPublishing() {
        using var json=OpenFixture();
        var frame=new OfficeAv1StillFrame {MiRows=32,MiCols=96,AllowIntraBlockCopy=true};
        var tile=new OfficeAv1Tile(0,1,0,32,0,96);
        foreach (var c in json.RootElement.GetProperty("rejections").EnumerateArray()) {
            byte[] bytes=Hex(c.GetProperty("hex").GetString()!);
            var context=new OfficeAv1CopyMotionReader(frame,new OfficeAv1StillSequence(),tile,new OfficeRasterDecodeOptions());
            var stream=new OfficeAv1SymbolReader(bytes,0,bytes.Length,1000,false);
            var error=Assert.Throws<FormatException>(()=>context.Read(stream,new OfficeAv1BlockRegion(0,0,8,8),Modes(true,true)));
            Assert.Contains("source displacement",error.Message); // Grammar reached source validation rather than exhausting the stream.
            Assert.Throws<FormatException>(()=>context.CompleteBlock());
            Assert.Throws<FormatException>(()=>context.Read(stream,new OfficeAv1BlockRegion(0,2,8,8),Modes(false,true)));
        }
    }

    [Fact]
    public void SourceProbesDistinguishBothPipelineDelaysTileEdgesAndSharedChromaFootprints() {
        using var json=OpenFixture();
        foreach (var c in json.RootElement.GetProperty("sourceProbes").EnumerateArray()) {
            int rows=c.GetProperty("rows").GetInt32(),cols=c.GetProperty("cols").GetInt32();
            bool mono=c.GetProperty("mono").GetBoolean();
            var frame=new OfficeAv1StillFrame {MiRows=rows,MiCols=cols,AllowIntraBlockCopy=true};
            var context=new OfficeAv1CopyMotionReader(frame,new OfficeAv1StillSequence {Monochrome=mono},
                new OfficeAv1Tile(0,1,0,rows,0,cols),new OfficeRasterDecodeOptions());
            byte[] bytes=Hex(c.GetProperty("hex").GetString()!);
            var stream=new OfficeAv1SymbolReader(bytes,0,bytes.Length,1000,false);
            foreach (var v in c.GetProperty("prefix").EnumerateArray()) {
                int[] s=v.EnumerateArray().Select(x=>x.GetInt32()).ToArray();
                context.Read(stream,new OfficeAv1BlockRegion(s[0],s[1],s[2],s[3]),Modes(false,Chroma(s,mono)));
                context.CompleteBlock();
            }
            int[] b=c.GetProperty("block").EnumerateArray().Select(v=>v.GetInt32()).ToArray();
            var block=new OfficeAv1BlockRegion(b[0],b[1],b[2],b[3]);
            if (c.GetProperty("valid").GetBoolean()) {
                var actual=context.Read(stream,block,Modes(true,Chroma(b,mono)));
                Assert.Equal(c.GetProperty("motion")[0].GetInt32(),actual.Row);
                Assert.Equal(c.GetProperty("motion")[1].GetInt32(),actual.Col);
                context.CompleteBlock();stream.Finish();
            } else {
                var error=Assert.Throws<FormatException>(()=>context.Read(stream,block,Modes(true,Chroma(b,mono))));
                Assert.Contains("source displacement",error.Message);
                Assert.Throws<FormatException>(()=>context.CompleteBlock());
            }
        }
    }

    [Fact]
    public void CopyLifetimeAndResourceChecksPreserveUntouchedEntropy() {
        var frame=new OfficeAv1StillFrame {MiRows=32,MiCols=96};
        var tile=new OfficeAv1Tile(0,1,0,32,0,96);var sequence=new OfficeAv1StillSequence();
        var context=new OfficeAv1CopyMotionReader(frame,sequence,tile,new OfficeRasterDecodeOptions());
        var stream=new OfficeAv1SymbolReader(Hex("20"),0,1,1,true);var block=new OfficeAv1BlockRegion(0,0,8,8);
        Assert.Throws<InvalidOperationException>(()=>context.CompleteBlock());
        Assert.Throws<FormatException>(()=>context.Read(stream,block,Modes(true,true)));
        Assert.Throws<FormatException>(()=>context.Read(stream,new OfficeAv1BlockRegion(0,1,8,8),Modes(false,true)));
        var motion=context.Read(stream,block,Modes(false,true));Assert.Equal(0,motion.Row);Assert.Equal(0,motion.Col);
        Assert.Throws<InvalidOperationException>(()=>context.Read(stream,new OfficeAv1BlockRegion(0,2,8,8),Modes(false,true)));
        context.CompleteBlock();
        Assert.Throws<FormatException>(()=>context.Read(stream,block,Modes(false,true)));
        Assert.False(stream.ReadBool());stream.Finish();
        Assert.Throws<FormatException>(()=>new OfficeAv1CopyMotionReader(frame,sequence,tile,new OfficeRasterDecodeOptions {MaximumDecodedPixels=1}));
        Assert.Throws<FormatException>(()=>new OfficeAv1CopyMotionReader(frame,sequence,tile,new OfficeRasterDecodeOptions {RetainedManagedBytes=long.MaxValue}));
    }

    [Fact]
    public void CanceledCopyLeafCannotPublishNeighborState() {
        using var cancellation=new CancellationTokenSource();
        var frame=new OfficeAv1StillFrame {MiRows=32,MiCols=96};
        var context=new OfficeAv1CopyMotionReader(frame,new OfficeAv1StillSequence(),new OfficeAv1Tile(0,1,0,32,0,96),
            new OfficeRasterDecodeOptions {CancellationToken=cancellation.Token});
        var stream=new OfficeAv1SymbolReader(Hex("20"),0,1,1,true);
        context.Read(stream,new OfficeAv1BlockRegion(0,0,8,8),Modes(false,true));
        cancellation.Cancel();Assert.Throws<OperationCanceledException>(()=>context.CompleteBlock());
        Assert.Throws<OperationCanceledException>(()=>context.Read(stream,new OfficeAv1BlockRegion(0,2,8,8),Modes(false,true)));
        Assert.False(stream.ReadBool());stream.Finish();
    }

    private static OfficeAv1IntraModes Modes(bool copy,bool chroma) => new OfficeAv1IntraModes(copy,chroma,false,
        OfficeAv1IntraMode.Dc,OfficeAv1IntraMode.Dc,0,0,0,0);
    private static bool Chroma(int[] b,bool mono) => !mono && !(b[3]==4&&(b[0]&1)==0) && !(b[2]==4&&(b[1]&1)==0);
    private static JsonDocument OpenFixture() {
        using var file=File.OpenRead(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif","copy-reference.json.gz"));
        using var gzip=new GZipStream(file,CompressionMode.Decompress);return JsonDocument.Parse(gzip);
    }
    private static byte[] Hex(string hex) => Convert.FromHexString(hex);
}
