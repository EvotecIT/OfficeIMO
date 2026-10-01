using System.IO.Compression;
using System.Text.Json;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Native frame pixels and numeric kernels protect CDEF ordering, padded edges and signed arithmetic.</summary>
public sealed class DrawingAv1CdefTests {
    [Theory]
    [InlineData("avif-opaque",false)]
    [InlineData("avif-alpha",false)]
    [InlineData("avif-alpha",true)]
    [InlineData("multitile",false)]
    [InlineData("lossless-copy",false)]
    [InlineData("odd-color",false)]
    [InlineData("odd-mono",false)]
    [InlineData("odd-tiled-color",false)]
    [InlineData("svt-tiled-mixed-skip",false)]
    public void CompleteFramesMatchNativeDeblockedAndCdefPixels(string name,bool alpha) {
        using var fixture=Open();var c=fixture.RootElement.GetProperty("cases").EnumerateArray().Single(v=>
            v.GetProperty("name").GetString()==name && v.GetProperty("alpha").GetBoolean()==alpha);
        var (bytes,sequence,frame)=Read(c,name,alpha);
        if(name=="odd-tiled-color" || name=="svt-tiled-mixed-skip") {
            var native=c.GetProperty("unfiltered");
            Assert.Equal(native.GetProperty("tileCols").GetInt32()*native.GetProperty("tileRows").GetInt32(),frame.Tiles.Length);
            Assert.True(frame.Tiles.Length>1);
            if(name=="svt-tiled-mixed-skip") {
                int skip=native.GetProperty("skipLeaves").GetInt32();
                Assert.True(skip>0 && skip<native.GetProperty("leaves").GetInt32());
            }
        }
        foreach(var stage in new[] {OfficeAv1ReconstructionStage.Deblocked,OfficeAv1ReconstructionStage.Cdef}) {
            var result=OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions(),stage);
            var planes=c.GetProperty(stage==OfficeAv1ReconstructionStage.Cdef?"cdef":"deblocked").GetProperty("planes");
            Assert.Equal(result.PlaneCount,planes.GetArrayLength());
            for(int p=0;p<result.PlaneCount;p++) {
                byte[] pixels=Convert.FromBase64String(planes[p].GetString()!);int sub=p==0?0:1,w=frame.MiCols*4>>sub,h=frame.MiRows*4>>sub;
                Assert.Equal(w*h,pixels.Length);
                for(int y=0;y<h;y++)for(int x=0;x<w;x++)
                    Assert.True(pixels[y*w+x]==result.Value(p,x,y),$"{name}/{alpha} {stage}, plane {p} ({x},{y}): native {pixels[y*w+x]}, managed {result.Value(p,x,y)}");
            }
        }
    }
    [Fact]
    public void DirectionVarianceAndFiltersMatchNativeIncludingUnavailableEdges() {
        using var fixture=Open();var partial=new int[120];var cost=new int[8];int changed=0,edgeChanged=0;
        foreach(var c in fixture.RootElement.GetProperty("directions").EnumerateArray()) {
            int direction=OfficeAv1CdefFilter.FindDirection(Hex(c.GetProperty("input").GetString()!),12,0,partial,cost,out int variance);
            Assert.Equal(c.GetProperty("direction").GetInt32(),direction);Assert.Equal(c.GetProperty("variance").GetInt32(),variance);
        }
        foreach(var c in fixture.RootElement.GetProperty("kernels").EnumerateArray()) {
            byte[] input=Hex(c.GetProperty("input").GetString()!),output=(byte[])input.Clone();int x=c.GetProperty("x").GetInt32();
            OfficeAv1CdefFilter.Apply(input,output,12,x,c.GetProperty("y").GetInt32(),c.GetProperty("size").GetInt32(),
                c.GetProperty("width").GetInt32(),c.GetProperty("height").GetInt32(),c.GetProperty("primary").GetInt32(),
                c.GetProperty("secondary").GetInt32(),c.GetProperty("damping").GetInt32(),c.GetProperty("direction").GetInt32());
            Assert.True(output.SequenceEqual(Hex(c.GetProperty("output").GetString()!)),
                $"CDEF size {c.GetProperty("size")}, dir {c.GetProperty("direction")}, primary {c.GetProperty("primary")}, secondary {c.GetProperty("secondary")}, damping {c.GetProperty("damping")}, x {x}");
            if(c.GetProperty("input").GetString()!=c.GetProperty("output").GetString()) {changed++;if(x==0)edgeChanged++;}
        }
        Assert.True(changed>0 && edgeChanged>0);
    }
    [Fact]
    public void ImmutableFilterInputSharesTheAggregateRetainedBudget() {
        using var fixture=Open();var c=fixture.RootElement.GetProperty("cases").EnumerateArray().Single(v=>v.GetProperty("name").GetString()=="odd-color");
        var (bytes,sequence,frame)=Read(c,"odd-color",false);int sb=sequence.Use128Superblock?128:64;
        long storage=(long)((frame.Width+sb-1)/sb*sb)*((frame.Height+sb-1)/sb*sb)*3/2+(long)frame.MiRows*frame.MiCols*2;
        long reserve=storage+131072+OfficeAv1IntraPredictor.ContextBytes+OfficeAv1ResidualTransform.ContextBytes+
            OfficeAv1Deblocker.ContextBytes(frame,false)+frame.Tiles.Max(t=>OfficeAv1TileReader.ContextBytes(frame,t))+bytes.LongLength;
        var options=new OfficeRasterDecodeOptions {RetainedManagedBytes=OfficeRasterGuards.MaximumDecodedBytes-reserve-8192};
        Assert.NotNull(OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,options,OfficeAv1ReconstructionStage.Deblocked));
        Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,options,OfficeAv1ReconstructionStage.Cdef));
    }
    private static JsonDocument Open() {
        using var input=File.OpenRead(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif", "cdef-reference.json.gz"));
        using var gzip=new GZipStream(input,CompressionMode.Decompress);return JsonDocument.Parse(gzip);
    }
    private static byte[] Hex(string value)=>Enumerable.Range(0,value.Length/2).Select(i=>Convert.ToByte(value.Substring(i*2,2),16)).ToArray();
    private static (byte[],OfficeAv1StillSequence,OfficeAv1StillFrame) Read(JsonElement c,string name,bool alpha) {
        byte[] bytes;OfficeAvifImageItem item;var options=new OfficeRasterDecodeOptions();
        if(c.TryGetProperty("inputBase64",out var input)) {
            bytes=Convert.FromBase64String(input.GetString()!);var native=c.GetProperty("unfiltered");
            item=new OfficeAvifImageItem(1,native.GetProperty("width").GetInt32(),native.GetProperty("height").GetInt32(),0,bytes.Length,
                new byte[] {0x81,(byte)native.GetProperty("level").GetInt32(),(byte)(c.GetProperty("monochrome").GetBoolean()?28:12),0},c.GetProperty("monochrome").GetBoolean(),null);
        } else {
            bytes=File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif",name+".avif"));
            Assert.True(OfficeAvifContainerReader.TryRead(bytes,options,out var container));item=alpha?container!.Alpha!:container!.Color;
        }
        Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes,item,options,out var sequence));
        Assert.True(OfficeAv1StillFrameReader.TryRead(bytes,item,sequence!,options,out var frame));return(bytes,sequence!,frame!);
    }
}
