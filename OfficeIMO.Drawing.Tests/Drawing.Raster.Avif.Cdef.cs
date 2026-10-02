using System.IO.Compression;
using System.Text.Json;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Native frame pixels and numeric kernels protect CDEF ordering, padded edges and signed arithmetic.</summary>
public sealed class DrawingAv1CdefTests {
    [Theory]
    [InlineData("avif-opaque",false,8)]
    [InlineData("avif-alpha",false,8)]
    [InlineData("avif-alpha",true,8)]
    [InlineData("multitile",false,8)]
    [InlineData("lossless-copy",false,8)]
    [InlineData("odd-color",false,8)]
    [InlineData("odd-mono",false,8)]
    [InlineData("odd-tiled-color",false,8)]
    [InlineData("svt-tiled-mixed-skip",false,8)]
    [InlineData("avif-main10-420-full",false,10)]
    [InlineData("avif-main10-420-limited",false,10)]
    [InlineData("avif-main10-420-full-alpha",false,10)]
    [InlineData("avif-main10-420-full-alpha",true,10)]
    [InlineData("avif-main10-420-limited-alpha",false,10)]
    [InlineData("avif-main10-420-limited-alpha",true,10)]
    [InlineData("avif-main10-mono-full",false,10)]
    [InlineData("avif-main10-mono-limited",false,10)]
    [InlineData("avif-main10-mono-full-alpha",false,10)]
    [InlineData("avif-main10-mono-full-alpha",true,10)]
    [InlineData("avif-main10-mono-limited-alpha",false,10)]
    [InlineData("avif-main10-mono-limited-alpha",true,10)]
    [InlineData("lossless-copy",false,10)]
    [InlineData("odd-color",false,10)]
    [InlineData("odd-mono",false,10)]
    [InlineData("odd-tiled-color",false,10)]
    public void CompleteFramesMatchNativeDeblockedAndCdefPixels(string name,bool alpha,int depth) {
        using var fixture=Open(depth);var c=fixture.RootElement.GetProperty("cases").EnumerateArray().Single(v=>
            v.GetProperty("name").GetString()==name && v.GetProperty("alpha").GetBoolean()==alpha);
        var (bytes,sequence,frame)=Read(c,name,alpha,depth);
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
            Assert.Equal(depth,result.BitDepth);
            var planes=c.GetProperty(stage==OfficeAv1ReconstructionStage.Cdef?"cdef":"deblocked").GetProperty("planes");
            Assert.Equal(result.PlaneCount,planes.GetArrayLength());
            for(int p=0;p<result.PlaneCount;p++) {
                ushort[] pixels=Samples(Convert.FromBase64String(planes[p].GetString()!),depth);int sub=p==0?0:1,w=frame.MiCols*4>>sub,h=frame.MiRows*4>>sub;
                Assert.Equal(w*h,pixels.Length);
                for(int y=0;y<h;y++)for(int x=0;x<w;x++)
                    Assert.True(pixels[y*w+x]==result.Value(p,x,y),$"{name}/{alpha}/{depth} {stage}, plane {p} ({x},{y}): native {pixels[y*w+x]}, managed {result.Value(p,x,y)}");
            }
        }
    }
    [Theory]
    [InlineData(8)]
    [InlineData(10)]
    public void DirectionVarianceAndFiltersMatchNativeIncludingUnavailableEdges(int depth) {
        using var fixture=Open(depth);var partial=new int[120];var cost=new int[8];int changed=0,edgeChanged=0;
        foreach(var c in fixture.RootElement.GetProperty("directions").EnumerateArray()) {
            int direction=OfficeAv1CdefFilter.FindDirection(Hex(c.GetProperty("input").GetString()!,depth),12,0,depth,partial,cost,out int variance);
            Assert.Equal(c.GetProperty("direction").GetInt32(),direction);Assert.Equal(c.GetProperty("variance").GetInt32(),variance);
        }
        foreach(var c in fixture.RootElement.GetProperty("kernels").EnumerateArray()) {
            ushort[] input=Hex(c.GetProperty("input").GetString()!,depth),output=(ushort[])input.Clone();int x=c.GetProperty("x").GetInt32();
            int shift=depth-8;
            OfficeAv1CdefFilter.Apply(input,output,12,x,c.GetProperty("y").GetInt32(),c.GetProperty("size").GetInt32(),
                c.GetProperty("width").GetInt32(),c.GetProperty("height").GetInt32(),c.GetProperty("primary").GetInt32()<<shift,
                c.GetProperty("secondary").GetInt32()<<shift,c.GetProperty("damping").GetInt32()+shift,c.GetProperty("direction").GetInt32(),depth);
            Assert.True(output.SequenceEqual(Hex(c.GetProperty("output").GetString()!,depth)),
                $"CDEF size {c.GetProperty("size")}, dir {c.GetProperty("direction")}, primary {c.GetProperty("primary")}, secondary {c.GetProperty("secondary")}, damping {c.GetProperty("damping")}, x {x}");
            if(c.GetProperty("input").GetString()!=c.GetProperty("output").GetString()) {changed++;if(x==0)edgeChanged++;}
        }
        Assert.True(changed>0 && edgeChanged>0);
    }
    [Theory]
    [InlineData(8)]
    [InlineData(10)]
    public void ImmutableFilterInputSharesTheAggregateRetainedBudget(int depth) {
        using var fixture=Open(depth);var c=fixture.RootElement.GetProperty("cases").EnumerateArray().Single(v=>v.GetProperty("name").GetString()=="odd-color");
        var (bytes,sequence,frame)=Read(c,"odd-color",false,depth);int sb=sequence.Use128Superblock?128:64;
        long storage=(long)((frame.Width+sb-1)/sb*sb)*((frame.Height+sb-1)/sb*sb)*3+(long)frame.MiRows*frame.MiCols*2;
        long reserve=storage+131072+OfficeAv1IntraPredictor.ContextBytes+OfficeAv1ResidualTransform.ContextBytes+
            OfficeAv1Deblocker.ContextBytes(frame,false)+frame.Tiles.Max(t=>OfficeAv1TileReader.ContextBytes(frame,t))+bytes.LongLength;
        var options=new OfficeRasterDecodeOptions {RetainedManagedBytes=OfficeRasterGuards.MaximumDecodedBytes-reserve-8192};
        Assert.NotNull(OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,options,OfficeAv1ReconstructionStage.Deblocked));
        Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,options,OfficeAv1ReconstructionStage.Cdef));
    }
    private static JsonDocument Open(int depth=8) {
        using var input=File.OpenRead(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif",depth==10?"cdef-main10-reference.json.gz":"cdef-reference.json.gz"));
        using var gzip=new GZipStream(input,CompressionMode.Decompress);return JsonDocument.Parse(gzip);
    }
    private static ushort[] Hex(string value,int depth)=>Samples(Enumerable.Range(0,value.Length/2).Select(i=>Convert.ToByte(value.Substring(i*2,2),16)).ToArray(),depth);
    private static ushort[] Samples(byte[] bytes,int depth)=>Enumerable.Range(0,bytes.Length/(depth==10?2:1))
        .Select(i=>(ushort)(depth==10?bytes[i*2]+bytes[i*2+1]*256:bytes[i])).ToArray();
    private static (byte[],OfficeAv1StillSequence,OfficeAv1StillFrame) Read(JsonElement c,string name,bool alpha,int depth) {
        byte[] bytes;OfficeAvifImageItem item;var options=new OfficeRasterDecodeOptions();
        if(c.TryGetProperty("inputBase64",out var input)) {
            bytes=Convert.FromBase64String(input.GetString()!);var native=c.GetProperty("unfiltered");
            item=new OfficeAvifImageItem(1,native.GetProperty("width").GetInt32(),native.GetProperty("height").GetInt32(),0,bytes.Length,
                new byte[] {0x81,(byte)native.GetProperty("level").GetInt32(),(byte)((depth==10?76:12)+(c.GetProperty("monochrome").GetBoolean()?16:0)),0},false,null);
        } else {
            bytes=File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif",name+".avif"));
            Assert.True(OfficeAvifContainerReader.TryRead(bytes,options,out var container));item=alpha?container!.Alpha!:container!.Color;
        }
        Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes,item,options,out var sequence));
        Assert.True(OfficeAv1StillFrameReader.TryRead(bytes,item,sequence!,options,out var frame));return(bytes,sequence!,frame!);
    }
}
