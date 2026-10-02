using System.IO.Compression;
using System.Text.Json;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Actual native frame and kernel pixels protect deblocking geometry and fixed-point arithmetic.</summary>
public sealed class DrawingAv1DeblockingTests {
    [Theory]
    [InlineData("avif-opaque",false,8)]
    [InlineData("avif-alpha",false,8)]
    [InlineData("avif-alpha",true,8)]
    [InlineData("multitile",false,8)]
    [InlineData("lossless-copy",false,8)]
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
    public void CompleteFrameMatchesNativePostDeblockPlanes(string name,bool alpha,int depth) {
        using var fixture=Open(depth);var c=fixture.RootElement.GetProperty("cases").EnumerateArray().Single(v=>
            v.GetProperty("name").GetString()==name && v.GetProperty("alpha").GetBoolean()==alpha);
        var (bytes,sequence,frame)=c.TryGetProperty("inputBase64",out var input)?ReadControl(c,input,depth):Read(name,alpha);
        var result=OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions(),OfficeAv1ReconstructionStage.Deblocked);
        Assert.Equal(depth,result.BitDepth);
        var planes=c.GetProperty("deblocked").GetProperty("planes");Assert.Equal(result.PlaneCount,planes.GetArrayLength());
        int changed=0;
        for(int p=0;p<result.PlaneCount;p++) {
            ushort[] pixels=Samples(Convert.FromBase64String(planes[p].GetString()!),depth);
            ushort[] unfiltered=Samples(Convert.FromBase64String(c.GetProperty("unfiltered").GetProperty("planes")[p].GetString()!),depth);
            int sub=p==0?0:1,w=frame.MiCols*4>>sub,h=frame.MiRows*4>>sub;
            Assert.Equal(w*h,pixels.Length);
            changed+=pixels.Zip(unfiltered,(a,b)=>a!=b?1:0).Sum();
            for(int y=0;y<h;y++) for(int x=0;x<w;x++)
                Assert.True(pixels[y*w+x]==result.Value(p,x,y),$"{name}/{alpha}/{depth}, plane {p} ({x},{y}): native {pixels[y*w+x]}, managed {result.Value(p,x,y)}");
        }
        if(name.StartsWith("odd-",StringComparison.Ordinal)) {
            Assert.True(changed>0);
            Assert.Contains(frame.LoopFilterLevels,v=>v>0);
            if(name=="odd-tiled-color") Assert.True(frame.Tiles.Length>1);
        }
    }
    [Theory]
    [InlineData(8)]
    [InlineData(10)]
    public void NarrowAndWideKernelsMatchNativeMasksRoundingAndUntouchedPixels(int depth) {
        using var fixture=Open(depth);var scratch=new int[32];int changed=0,reversedChanged=0;
        foreach(var c in fixture.RootElement.GetProperty("kernels").EnumerateArray()) {
            ushort[] pixels=Samples(Hex(c.GetProperty("input").GetString()!),depth);
            ushort[] expected=Samples(Hex(c.GetProperty("output").GetString()!),depth);
            bool horizontal=c.GetProperty("pass").GetInt32()!=0;int step=horizontal?16:1;
            for(int i=0;i<4;i++) OfficeAv1DeblockFilter.Apply(pixels,8*16+8+i*(horizontal?1:16),step,
                c.GetProperty("size").GetInt32(),c.GetProperty("chroma").GetBoolean(),c.GetProperty("limit").GetInt32(),
                c.GetProperty("blimit").GetInt32(),c.GetProperty("threshold").GetInt32(),depth,scratch);
            Assert.True(expected.SequenceEqual(pixels),$"Kernel size {c.GetProperty("size")}, chroma {c.GetProperty("chroma")}, pass {c.GetProperty("pass")}, level {c.GetProperty("level")}, sharpness {c.GetProperty("sharpness")}, pattern {c.GetProperty("pattern")}");
            if(c.GetProperty("input").GetString()!=c.GetProperty("output").GetString()) {
                changed++;
                if(c.GetProperty("pattern").GetInt32()>=9) reversedChanged++;
            }
        }
        Assert.True(changed>0);
        Assert.True(reversedChanged>0);
    }
    [Theory]
    [InlineData(8)]
    [InlineData(10)]
    public void FilterMetadataSharesTheAggregateRetainedBudget(int depth) {
        using var fixture=Open(depth);
        var control=fixture.RootElement.GetProperty("cases").EnumerateArray().Single(v=>v.GetProperty("name").GetString()==(depth==10?"odd-tiled-color":"multitile"));
        var (bytes,sequence,frame)=depth==10?ReadControl(control,control.GetProperty("inputBase64"),depth):Read("multitile",false);
        int sb=sequence.Use128Superblock?128:64;
        long storage=(long)((frame.Width+sb-1)/sb*sb)*((frame.Height+sb-1)/sb*sb)*(sequence.Monochrome?2:3)+
            (long)frame.MiRows*frame.MiCols*2;
        long reserve=storage+131072+OfficeAv1IntraPredictor.ContextBytes+OfficeAv1ResidualTransform.ContextBytes+
            frame.Tiles.Max(t=>OfficeAv1TileReader.ContextBytes(frame,t))+bytes.LongLength;
        var options=new OfficeRasterDecodeOptions {RetainedManagedBytes=OfficeRasterGuards.MaximumDecodedBytes-reserve-8192};
        Assert.NotNull(OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,options));
        Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,options,OfficeAv1ReconstructionStage.Deblocked));
    }
    private static JsonDocument Open(int depth=8) {
        using var input=File.OpenRead(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif",depth==10?"deblock-main10-reference.json.gz":"deblock-reference.json.gz"));
        using var gzip=new GZipStream(input,CompressionMode.Decompress);return JsonDocument.Parse(gzip);
    }
    private static byte[] Hex(string value)=>Enumerable.Range(0,value.Length/2).Select(i=>Convert.ToByte(value.Substring(i*2,2),16)).ToArray();
    private static ushort[] Samples(byte[] bytes,int depth)=>Enumerable.Range(0,bytes.Length/(depth==10?2:1))
        .Select(i=>(ushort)(depth==10?bytes[i*2]+bytes[i*2+1]*256:bytes[i])).ToArray();
    private static (byte[],OfficeAv1StillSequence,OfficeAv1StillFrame) ReadControl(JsonElement c,JsonElement input,int depth) {
        byte[] bytes=Convert.FromBase64String(input.GetString()!);var native=c.GetProperty("unfiltered");
        var item=new OfficeAvifImageItem(1,native.GetProperty("width").GetInt32(),native.GetProperty("height").GetInt32(),0,bytes.Length,
            new byte[] {0x81,(byte)native.GetProperty("level").GetInt32(),(byte)((depth==10?76:12)+(native.GetProperty("planes").GetArrayLength()==1?16:0)),0},false,null);
        var options=new OfficeRasterDecodeOptions();
        Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes,item,options,out var sequence));
        Assert.True(OfficeAv1StillFrameReader.TryRead(bytes,item,sequence!,options,out var frame));return (bytes,sequence!,frame!);
    }
    private static (byte[],OfficeAv1StillSequence,OfficeAv1StillFrame) Read(string name,bool alpha) {
        var bytes=File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif",name+".avif"));var options=new OfficeRasterDecodeOptions();
        Assert.True(OfficeAvifContainerReader.TryRead(bytes,options,out var container));var item=alpha?container!.Alpha!:container!.Color;
        Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes,item,options,out var sequence));
        Assert.True(OfficeAv1StillFrameReader.TryRead(bytes,item,sequence!,options,out var frame));return (bytes,sequence!,frame!);
    }
}
