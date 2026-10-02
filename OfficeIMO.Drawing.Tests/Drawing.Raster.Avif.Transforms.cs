using System.Text.Json;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Full-grid native oracle protects size CDF contexts, variable trees, residual order and clipped frame edges.</summary>
public sealed class DrawingAv1TransformTests {
    [Fact]
    public void NativeTransformStreamsMatchGridsAndResidualOrderAcrossIndependentTiles() {
        using var json=OpenFixture();
        foreach (var c in json.RootElement.GetProperty("cases").EnumerateArray()) {
            int extent=c.GetProperty("extent").GetInt32(), mode=c.GetProperty("mode").GetInt32();
            bool large=c.GetProperty("states")[0].GetProperty("block")[2].GetInt32()>64 ||
                c.GetProperty("states")[0].GetProperty("block")[3].GetInt32()>64;
            int units=large?32:16;
            var contexts=Enumerable.Range(0,2).Select(tile=>new OfficeAv1TransformReader(
                new OfficeAv1StillFrame { MiRows=extent+tile*units,MiCols=extent+tile*units,TransformMode=mode },
                new OfficeAv1StillSequence { Use128Superblock=large },
                new OfficeAv1Tile(0,1,tile*units,tile*units+extent,tile*units,tile*units+extent),new OfficeRasterDecodeOptions())).ToArray();
            byte[] bytes=Hex(c.GetProperty("hex").GetString()!);
            var streams=Enumerable.Range(0,2).Select(_=>new OfficeAv1SymbolReader(bytes,0,bytes.Length,1000000,!c.GetProperty("updates").GetBoolean())).ToArray();
            foreach (var expected in c.GetProperty("states").EnumerateArray()) {
                int[] b=expected.GetProperty("block").EnumerateArray().Select(v=>v.GetInt32()).ToArray();
                var prelude=Prelude(expected.GetProperty("skip").GetBoolean(),expected.GetProperty("lossless").GetBoolean());
                var modes=Modes(expected.GetProperty("inter").GetBoolean(),expected.GetProperty("chroma").GetBoolean());
                for (int tile=0;tile<2;tile++) {
                    var block=new OfficeAv1BlockRegion(b[0]+tile*units,b[1]+tile*units,b[2],b[3]);
                    AssertLayout(expected,contexts[tile].Read(streams[tile],block,prelude,modes),tile*units);
                    contexts[tile].CompleteBlock();
                }
            }
            foreach (var stream in streams) stream.Finish(); // Isolated transform component, not a complete AV1 tile.
        }
    }

    [Fact]
    public void FrozenFirstLeavesReachCoefficientBoundaryWithNativeTransformGeometry() {
        using var json=OpenFixture();
        foreach (var c in json.RootElement.GetProperty("framePrefixes").EnumerateArray()) {
            byte[] bytes=File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif",c.GetProperty("name").GetString()!+".avif"));
            var options=new OfficeRasterDecodeOptions();
            Assert.True(OfficeAvifContainerReader.TryRead(bytes,options,out var container));
            var item=c.GetProperty("alpha").GetBoolean()?container!.Alpha!:container!.Color;
            Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes,item,options,out var sequence));
            Assert.True(OfficeAv1StillFrameReader.TryRead(bytes,item,sequence!,options,out var frame));
            Assert.Equal(c.GetProperty("mode").GetInt32(),frame!.TransformMode);
            var tile=frame.Tiles[0];var symbols=new OfficeAv1SymbolReader(bytes,tile.Offset,tile.Length,100000,frame.DisableCdfUpdate);
            var partitions=new OfficeAv1PartitionContext(frame,tile,sequence!.Use128Superblock?128:64);
            int pixels=sequence.Use128Superblock?128:64;OfficeAv1PartitionLayout layout;
            while ((layout=partitions.Read(symbols,0,0,pixels)).Kind==OfficeAv1Partition.Split) pixels/=2;
            Assert.Equal(OfficeAv1Partition.None,layout.Kind);Assert.Equal(c.GetProperty("pixels").GetInt32(),pixels);
            var block=layout.Child(0);var preludes=new OfficeAv1BlockPreludeReader(frame,sequence,tile,options);preludes.BeginSuperblock(0,0);
            var prelude=preludes.Read(symbols,block);Assert.False(prelude.Skip);Assert.False(prelude.Lossless);
            Assert.Equal(c.GetProperty("preludeQ").GetInt32(),prelude.CurrentQIndex);
            var modes=new OfficeAv1IntraModeReader(frame,sequence,tile,options).Read(symbols,block,prelude);
            Assert.False(modes.UseIntraBlockCopy);
            new OfficeAv1PaletteReader(frame,sequence,tile,options).Read(symbols,block,modes);
            var actual=new OfficeAv1TransformReader(frame,sequence,tile,options).Read(symbols,block,prelude,modes);
            AssertLayout(c,actual,0,pixels);
            // No coefficient/type syntax, reconstructed pixels, tile termination or neighbor publication is claimed.
        }
    }

    [Fact]
    public void TransformGeometryResourcesAndLeafLifetimeRejectWithoutConsumingAndKeepSnapshots() {
        var frame=new OfficeAv1StillFrame { MiRows=16,MiCols=16,TransformMode=2 };
        var sequence=new OfficeAv1StillSequence();var tile=new OfficeAv1Tile(0,1,0,16,0,16);
        var context=new OfficeAv1TransformReader(frame,sequence,tile,new OfficeRasterDecodeOptions());
        var untouched=new OfficeAv1SymbolReader(Hex("20"),0,1,1,true);
        Assert.Throws<InvalidOperationException>(()=>context.CompleteBlock());
        Assert.Throws<FormatException>(()=>context.Read(untouched,new OfficeAv1BlockRegion(0,15,8,8),Prelude(false,false),Modes(false,true)));
        Assert.False(untouched.ReadBool());untouched.Finish();
        Assert.Throws<FormatException>(()=>new OfficeAv1TransformReader(frame,sequence,tile,new OfficeRasterDecodeOptions { MaximumDecodedPixels=1 }));
        Assert.Throws<FormatException>(()=>new OfficeAv1TransformReader(frame,sequence,tile,new OfficeRasterDecodeOptions { RetainedManagedBytes=long.MaxValue }));
        using var json=OpenFixture();var c=json.RootElement.GetProperty("cases").EnumerateArray().First(x=>x.GetProperty("scenario").GetInt32()==3);
        byte[] bytes=Hex(c.GetProperty("hex").GetString()!);var symbols=new OfficeAv1SymbolReader(bytes,0,bytes.Length,10000,!c.GetProperty("updates").GetBoolean());
        var first=c.GetProperty("states")[0];var block=new OfficeAv1BlockRegion(0,0,8,8);
        var saved=context.Read(symbols,block,Prelude(false,false),Modes(false,true));
        Assert.Throws<InvalidOperationException>(()=>context.Read(symbols,block,Prelude(false,false),Modes(false,true)));
        context.CompleteBlock();context.Read(symbols,new OfficeAv1BlockRegion(0,2,8,8),Prelude(false,false),Modes(false,true));
        context.CompleteBlock();AssertLayout(first,saved,0);
    }

    [Fact]
    public void FailedAndCancelledTransformContextsCannotPublishPartialTrees() {
        using var json=OpenFixture();var c=json.RootElement.GetProperty("cases").EnumerateArray().First(x=>x.GetProperty("scenario").GetInt32()==37);
        byte[] bytes=Hex(c.GetProperty("hex").GetString()!);
        var frame=new OfficeAv1StillFrame { MiRows=32,MiCols=32,TransformMode=2 };
        var sequence=new OfficeAv1StillSequence { Use128Superblock=true };var tile=new OfficeAv1Tile(0,1,0,32,0,32);
        var block=new OfficeAv1BlockRegion(0,0,128,128);var modes=Modes(true,true);
        var context=new OfficeAv1TransformReader(frame,sequence,tile,new OfficeRasterDecodeOptions());
        var exhausted=new OfficeAv1SymbolReader(bytes,0,bytes.Length,1,true);
        Assert.Throws<FormatException>(()=>context.Read(exhausted,block,Prelude(false,false),modes));
        Assert.Throws<FormatException>(()=>context.CompleteBlock());
        Assert.Throws<FormatException>(()=>context.Read(exhausted,block,Prelude(false,false),modes));
        using var cancellation=new CancellationTokenSource();
        var cancelled=new OfficeAv1TransformReader(frame,sequence,tile,new OfficeRasterDecodeOptions { CancellationToken=cancellation.Token });
        var untouched=new OfficeAv1SymbolReader(Hex("20"),0,1,1,true);cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(()=>cancelled.Read(untouched,block,Prelude(false,true),modes));
        Assert.Throws<OperationCanceledException>(()=>cancelled.CompleteBlock());
        Assert.False(untouched.ReadBool());untouched.Finish();
    }

    private static void AssertLayout(JsonElement expected,OfficeAv1TransformLayout actual,int tileUnits,int pixels=0) {
        Assert.Equal(expected.GetProperty("last").GetInt32(),actual.LastSize);
        int width=pixels==0?expected.GetProperty("block")[2].GetInt32():pixels;
        byte[] grid=Hex(expected.GetProperty("grid").GetString()!);
        Assert.Equal(grid,Enumerable.Range(0,grid.Length).Select(i=>(byte)actual.SizeAt(i/(width/4),i%(width/4))).ToArray());
        var residuals=expected.GetProperty("residuals");Assert.Equal(residuals.GetArrayLength(),actual.Count);
        for(int i=0;i<actual.Count;i++) {
            var b=actual.Block(i);int sub=b.Plane==0?0:1;
            Assert.Equal(residuals[i].EnumerateArray().Select(v=>v.GetInt32()).ToArray(),new[] {b.Plane,b.X-(tileUnits*4>>sub),b.Y-(tileUnits*4>>sub),b.Size});
            Assert.InRange(b.Width,4,64);Assert.InRange(b.Height,4,64);
        }
    }
    private static OfficeAv1BlockPrelude Prelude(bool skip,bool lossless)=>new OfficeAv1BlockPrelude(skip,0,lossless,0,32,new int[4]);
    private static OfficeAv1IntraModes Modes(bool inter,bool chroma)=>new OfficeAv1IntraModes(inter,chroma,false,OfficeAv1IntraMode.Dc,OfficeAv1IntraMode.Dc,0,0,0,0);
    private static JsonDocument OpenFixture()=>JsonDocument.Parse(File.ReadAllText(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif","transform-reference.json")));
    private static byte[] Hex(string text)=>Enumerable.Range(0,text.Length/2).Select(i=>Convert.ToByte(text.Substring(i*2,2),16)).ToArray();
}
