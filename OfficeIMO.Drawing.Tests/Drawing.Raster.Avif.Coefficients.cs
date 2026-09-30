using System.Text.Json;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Native arithmetic streams, native scans and independent full-grid state protect coefficient contracts.</summary>
public sealed class DrawingAv1CoefficientTests {
    [Fact]
    public void NativeCoefficientStreamsMatchSignedValuesAndTypesAcrossIndependentTiles() {
        using var fixture=Open();
        foreach(var c in fixture.RootElement.GetProperty("cases").EnumerateArray()) {
            int extent=c.GetProperty("extent").GetInt32(),q=c.GetProperty("q").GetInt32();
            int scenario=c.GetProperty("scenario").GetInt32();
            bool large=c.GetProperty("states")[0].GetProperty("block")[2].GetInt32()>64 || c.GetProperty("states")[0].GetProperty("block")[3].GetInt32()>64;
            int units=large?32:16;
            var sequence=new OfficeAv1StillSequence { Use128Superblock=large };
            var frames=Enumerable.Range(0,2).Select(t=>new OfficeAv1StillFrame {
                MiRows=extent+t*units,MiCols=extent+t*units,BaseQIndex=q,TransformMode=c.GetProperty("mode").GetInt32(),ReducedTransformSet=scenario/22%3==2,SegmentationEnabled=true }).ToArray();
            foreach(var f in frames) {f.SegmentFeatures[1,0]=true;f.SegmentData[1,0]=-q;f.SegmentFeatures[2,0]=true;f.SegmentData[2,0]=255-q;}
            var tiles=Enumerable.Range(0,2).Select(t=>new OfficeAv1Tile(0,1,t*units,t*units+extent,t*units,t*units+extent)).ToArray();
            var readers=Enumerable.Range(0,2).Select(t=>new OfficeAv1CoefficientReader(frames[t],sequence,tiles[t],new OfficeRasterDecodeOptions())).ToArray();
            // The coefficient owner snapshots mutable frame probability/type eligibility inputs.
            foreach(var f in frames) {f.BaseQIndex=127;f.ReducedTransformSet=!f.ReducedTransformSet;f.SegmentData[1,0]=255;f.SegmentData[2,0]=-255;}
            var transforms=Enumerable.Range(0,2).Select(t=>new OfficeAv1TransformReader(frames[t],sequence,tiles[t],new OfficeRasterDecodeOptions())).ToArray();
            byte[] bytes=Hex(c.GetProperty("hex").GetString()!);
            var symbols=Enumerable.Range(0,2).Select(_=>new OfficeAv1SymbolReader(bytes,0,bytes.Length,10000000,!c.GetProperty("updates").GetBoolean())).ToArray();
            foreach(var expected in c.GetProperty("states").EnumerateArray()) {
                int[] b=Ints(expected.GetProperty("block"));
                var prelude=new OfficeAv1BlockPrelude(expected.GetProperty("skip").GetBoolean(),expected.GetProperty("segment").GetInt32(),expected.GetProperty("lossless").GetBoolean(),0,q,new int[4]);
                var modes=new OfficeAv1IntraModes(expected.GetProperty("copy").GetBoolean(),expected.GetProperty("chroma").GetBoolean(),false,
                    (OfficeAv1IntraMode)expected.GetProperty("y").GetInt32(),(OfficeAv1IntraMode)expected.GetProperty("uv").GetInt32(),0,0,0,0);
                for(int t=0;t<2;t++) {
                    var block=new OfficeAv1BlockRegion(b[0]+t*units,b[1]+t*units,b[2],b[3]);
                    var layout=transforms[t].Read(symbols[t],block,prelude,modes);
                    readers[t].BeginBlock(block,prelude,modes,Palette(block,expected.GetProperty("filter").GetInt32()),layout);
                    int i=0;
                    foreach(var result in expected.GetProperty("coefficients").EnumerateArray()) {
                        var actual=readers[t].Read(symbols[t]);AssertCoefficients(result,actual,t*units);
                        Assert.Equal(layout.Block(i++).Size,actual.Block.Size);
                    }
                    Assert.Equal(layout.Count,i);readers[t].CompleteBlock();transforms[t].CompleteBlock();
                }
            }
            foreach(var stream in symbols) stream.Finish(); // This is an isolated coefficient component stream, not an AV1 tile.
        }
    }

    [Fact]
    public void OwnedScansMatchEveryNormalizedNativeShapeAndClass() {
        using var fixture=Open();
        foreach(var scan in fixture.RootElement.GetProperty("scans").EnumerateObject()) {
            string[] parts=scan.Name.Split('_');int[] shape=parts[2].Split('x').Select(int.Parse).ToArray();
            int cls=parts[0]=="mrow"?2:parts[0]=="mcol"?1:0;
            Assert.Equal(Ints(scan.Value),OfficeAv1CoefficientReader.CreateScan(shape[0],shape[1],cls));
        }
    }

    [Fact]
    public void FrozenAvifFirstTransformsMatchNativeCoefficientValuesBeforeReconstruction() {
        using var fixture=Open();
        foreach(var c in fixture.RootElement.GetProperty("framePrefixes").EnumerateArray()) {
            byte[] bytes=File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif",c.GetProperty("name").GetString()!+".avif"));
            var options=new OfficeRasterDecodeOptions();Assert.True(OfficeAvifContainerReader.TryRead(bytes,options,out var container));
            var item=c.GetProperty("alpha").GetBoolean()?container!.Alpha!:container!.Color;
            Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes,item,options,out var sequence));
            Assert.True(OfficeAv1StillFrameReader.TryRead(bytes,item,sequence!,options,out var frame));
            var tile=frame!.Tiles[0];var symbols=new OfficeAv1SymbolReader(bytes,tile.Offset,tile.Length,100000,frame.DisableCdfUpdate);
            var partitions=new OfficeAv1PartitionContext(frame,tile,sequence!.Use128Superblock?128:64);
            int pixels=sequence.Use128Superblock?128:64;OfficeAv1PartitionLayout partition;
            while((partition=partitions.Read(symbols,0,0,pixels)).Kind==OfficeAv1Partition.Split)pixels/=2;
            Assert.Equal(OfficeAv1Partition.None,partition.Kind);var block=partition.Child(0);
            var preludes=new OfficeAv1BlockPreludeReader(frame,sequence,tile,options);preludes.BeginSuperblock(0,0);var prelude=preludes.Read(symbols,block);
            var modes=new OfficeAv1IntraModeReader(frame,sequence,tile,options).Read(symbols,block,prelude);
            var palette=new OfficeAv1PaletteReader(frame,sequence,tile,options).Read(symbols,block,modes);
            var layout=new OfficeAv1TransformReader(frame,sequence,tile,options).Read(symbols,block,prelude,modes);
            var reader=new OfficeAv1CoefficientReader(frame,sequence,tile,options);reader.BeginBlock(block,prelude,modes,palette,layout);
            AssertCoefficients(c.GetProperty("coefficient"),reader.Read(symbols),0);
            // Later coefficients, complete tile traversal, motion, dequantization and pixels are not qualified here.
        }
    }

    [Fact]
    public void LeafFailureCancellationAndResourceBoundsDoNotPublishPartialCoefficientState() {
        using var fixture=Open();var c=fixture.RootElement.GetProperty("cases").EnumerateArray().First(v=>v.GetProperty("scenario").GetInt32()==3 && v.GetProperty("q").GetInt32()==60);
        byte[] bytes=Hex(c.GetProperty("hex").GetString()!);var frame=new OfficeAv1StillFrame {MiRows=8,MiCols=8,BaseQIndex=60,TransformMode=1};
        var sequence=new OfficeAv1StillSequence();var tile=new OfficeAv1Tile(0,1,0,8,0,8);var options=new OfficeRasterDecodeOptions();
        var reader=new OfficeAv1CoefficientReader(frame,sequence,tile,options);var block=new OfficeAv1BlockRegion(0,0,8,8);
        var prelude=Prelude(false,false,60);var modes=new OfficeAv1IntraModes(false,true,false,OfficeAv1IntraMode.Dc,OfficeAv1IntraMode.Dc,0,0,0,0);
        var transformSymbols=new OfficeAv1SymbolReader(bytes,0,bytes.Length,10000,false);
        var layout=new OfficeAv1TransformReader(frame,sequence,tile,options).Read(transformSymbols,block,prelude,modes);
        Assert.Throws<InvalidOperationException>(()=>reader.CompleteBlock());
        reader.BeginBlock(block,prelude,modes,Palette(block,-1),layout);
        Assert.Throws<InvalidOperationException>(()=>reader.CompleteBlock());
        var exhausted=new OfficeAv1SymbolReader(bytes,0,bytes.Length,1,false);
        Assert.Throws<FormatException>(()=>reader.Read(exhausted));Assert.Throws<FormatException>(()=>reader.CompleteBlock());
        Assert.Throws<FormatException>(()=>reader.Read(transformSymbols));
        Assert.Throws<FormatException>(()=>new OfficeAv1CoefficientReader(frame,sequence,tile,new OfficeRasterDecodeOptions {MaximumDecodedPixels=1}));
        Assert.Throws<FormatException>(()=>new OfficeAv1CoefficientReader(frame,sequence,tile,new OfficeRasterDecodeOptions {RetainedManagedBytes=long.MaxValue}));
        using var cancellation=new CancellationTokenSource();var cancelled=new OfficeAv1CoefficientReader(frame,sequence,tile,new OfficeRasterDecodeOptions {CancellationToken=cancellation.Token});
        cancelled.BeginBlock(block,prelude,modes,Palette(block,-1),layout);cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(()=>cancelled.Read(transformSymbols));Assert.Throws<OperationCanceledException>(()=>cancelled.CompleteBlock());
    }

    [Fact]
    public void GolombLengthAndMaskedDcCategoryKeepTheBoundedSignContract() {
        using var fixture=Open();
        foreach(bool malformed in new[] {false,true}) {
            byte[] bytes=Hex(fixture.RootElement.GetProperty(malformed?"malformedGolombHex":"maskedDcHex").GetString()!);
            var frame=new OfficeAv1StillFrame {MiRows=2,MiCols=2,BaseQIndex=20,TransformMode=1};
            var sequence=new OfficeAv1StillSequence();var tile=new OfficeAv1Tile(0,1,0,2,0,2);var options=new OfficeRasterDecodeOptions();
            var reader=new OfficeAv1CoefficientReader(frame,sequence,tile,options);var transforms=new OfficeAv1TransformReader(frame,sequence,tile,options);
            var symbols=new OfficeAv1SymbolReader(bytes,0,bytes.Length,1000,true);
            var prelude=Prelude(false,false,20);var modes=new OfficeAv1IntraModes(false,false,false,OfficeAv1IntraMode.Dc,OfficeAv1IntraMode.Dc,0,0,0,0);
            var block=new OfficeAv1BlockRegion(0,0,4,4);var layout=transforms.Read(symbols,block,prelude,modes);
            reader.BeginBlock(block,prelude,modes,Palette(block,-1),layout);
            if(malformed) {
                Assert.Throws<FormatException>(()=>reader.Read(symbols));Assert.Throws<FormatException>(()=>reader.CompleteBlock());
            } else {
                var saved=reader.Read(symbols);Assert.Equal(1,saved.EndOfBlock);Assert.All(Enumerable.Range(0,saved.Count),i=>Assert.Equal(0,saved.Value(i)));
                reader.CompleteBlock();transforms.CompleteBlock();
                block=new OfficeAv1BlockRegion(0,1,4,4);layout=transforms.Read(symbols,block,prelude,modes);
                reader.BeginBlock(block,prelude,modes,Palette(block,-1),layout);var following=reader.Read(symbols);
                Assert.Equal(-1,following.Value(0));reader.CompleteBlock();transforms.CompleteBlock();symbols.Finish();
                Assert.All(Enumerable.Range(0,saved.Count),i=>Assert.Equal(0,saved.Value(i)));
            }
        }
    }

    private static void AssertCoefficients(JsonElement expected,OfficeAv1Coefficients actual,int units) {
        var b=actual.Block;int sub=b.Plane==0?0:1;
        Assert.Equal(Ints(expected.GetProperty("block")),new[] {b.Plane,b.X-(units*4>>sub),b.Y-(units*4>>sub),b.Size});
        Assert.Equal(expected.GetProperty("type").GetInt32(),(int)actual.Type);Assert.Equal(expected.GetProperty("eob").GetInt32(),actual.EndOfBlock);
        var values=new int[actual.Width*actual.Height];foreach(var pair in expected.GetProperty("values").EnumerateArray())values[pair[0].GetInt32()]=pair[1].GetInt32();
        Assert.Equal(values,Enumerable.Range(0,actual.Count).Select(actual.Value).ToArray());
    }
    private static OfficeAv1Palette Palette(OfficeAv1BlockRegion b,int filter)=>new OfficeAv1Palette(Array.Empty<byte>(),Array.Empty<byte>(),Array.Empty<byte>(),Array.Empty<byte>(),Array.Empty<byte>(),b.Width,b.Height,filter);
    private static OfficeAv1BlockPrelude Prelude(bool skip,bool lossless,int q)=>new OfficeAv1BlockPrelude(skip,0,lossless,0,q,new int[4]);
    private static int[] Ints(JsonElement value)=>value.EnumerateArray().Select(v=>v.GetInt32()).ToArray();
    private static byte[] Hex(string value)=>Enumerable.Range(0,value.Length/2).Select(i=>Convert.ToByte(value.Substring(i*2,2),16)).ToArray();
    private static JsonDocument Open() {
        using var file=File.OpenRead(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif","coefficient-reference.json.gz"));
        using var gzip=new System.IO.Compression.GZipStream(file,System.IO.Compression.CompressionMode.Decompress);
        return JsonDocument.Parse(gzip);
    }
}
