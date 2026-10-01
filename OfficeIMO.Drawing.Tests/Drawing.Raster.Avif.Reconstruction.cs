using System.IO.Compression;
using System.Text.Json;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Actual independent frame pixels protect the integration of entropy order, edges and inverse residuals.</summary>
public sealed class DrawingAv1ReconstructionTests {
    [Theory]
    [InlineData("avif-opaque",false,8)]
    [InlineData("avif-alpha",false,8)]
    [InlineData("avif-alpha",true,8)]
    [InlineData("multitile",false,8)]
    [InlineData("lossless-copy",false,8)]
    [InlineData("avif-main10-420-full",false,10)]
    [InlineData("avif-main10-420-limited",false,10)]
    [InlineData("avif-main10-420-full-alpha",false,10)]
    [InlineData("avif-main10-420-limited-alpha",false,10)]
    [InlineData("avif-main10-mono-full",false,10)]
    [InlineData("avif-main10-mono-limited",false,10)]
    [InlineData("avif-main10-mono-full-alpha",false,10)]
    [InlineData("avif-main10-mono-limited-alpha",false,10)]
    [InlineData("lossless-copy",false,10)]
    [InlineData("avif-main10-420-full-alpha",true,10)]
    [InlineData("avif-main10-420-limited-alpha",true,10)]
    [InlineData("avif-main10-mono-full-alpha",true,10)]
    [InlineData("avif-main10-mono-limited-alpha",true,10)]
    public void FrozenFramesMatchEveryNativePreFilterSample(string name,bool alpha,int depth) {
        using var fixture=Open(depth);
        var c=fixture.RootElement.GetProperty("cases").EnumerateArray().Single(v=>v.GetProperty("name").GetString()==name && v.GetProperty("alpha").GetBoolean()==alpha);
        var (bytes,sequence,frame)=c.TryGetProperty("inputBase64",out var input)?ReadControl(c,input,depth):Read(name,alpha);
        var expected=c.GetProperty("unfiltered");var result=OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions());
        if(name=="lossless-copy") {
            Assert.True(expected.GetProperty("skipLeaves").GetInt32()>0);
            Assert.True(expected.GetProperty("copyLeaves").GetInt32()>0);
            Assert.True(expected.GetProperty("fractionalCopyLeaves").GetInt32()>0);
        }
        Assert.Equal(depth,result.BitDepth);
        int sb=sequence.Use128Superblock?128:64;
        Assert.Equal((long)((frame.Width+sb-1)/sb*sb)*((frame.Height+sb-1)/sb*sb)*(sequence.Monochrome?2:3),result.StorageBytes);
        Assert.Equal(frame.Width,result.Width);Assert.Equal(frame.Height,result.Height);Assert.Equal(expected.GetProperty("planes").GetArrayLength(),result.PlaneCount);
        for(int p=0;p<result.PlaneCount;p++) {
            byte[] pixels=Convert.FromBase64String(expected.GetProperty("planes")[p].GetString()!);
            int sub=p==0?0:1,w=frame.MiCols*4>>sub,h=frame.MiRows*4>>sub,sampleBytes=depth==10?2:1;
            Assert.Equal(w*h*sampleBytes,pixels.Length);
            for(int y=0;y<h;y++) for(int x=0;x<w;x++) {
                int offset=(y*w+x)*sampleBytes,native=pixels[offset]+(sampleBytes==2?pixels[offset+1]*256:0);
                Assert.True(native==result.Value(p,x,y),$"{name}/{alpha}/{depth}, plane {p} ({x},{y}): native {native}, managed {result.Value(p,x,y)}");
            }
        }
    }
    [Theory]
    [InlineData("avif-opaque",false)]
    [InlineData("avif-main10-420-full",false)]
    [InlineData("avif-opaque",true)]
    [InlineData("avif-main10-420-full",true)]
    public void FrameFailureCannotPublishPartialPlanesAndLimitsRejectBeforeAllocation(string name,bool deblocked) {
        var (bytes,sequence,frame)=Read(name,false);
        var stage=deblocked?OfficeAv1ReconstructionStage.Deblocked:OfficeAv1ReconstructionStage.Unfiltered;
        Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions {RetainedManagedBytes=long.MaxValue},stage));
        Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions {RetainedManagedBytes=OfficeRasterGuards.MaximumDecodedBytes-500000},stage));
        Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions {MaximumDecodedPixels=frame.Width*frame.Height-1},stage));
        Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions {MaximumEncodedBytes=bytes.Length-1},stage));
        Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions {MaximumInspectionWorkPixels=1},stage));
        using var cancel=new CancellationTokenSource();cancel.Cancel();
        Assert.Throws<OperationCanceledException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions {CancellationToken=cancel.Token},stage));
        var tile=frame.Tiles[0];frame.Tiles=new[] {new OfficeAv1Tile(tile.Offset,tile.Length-1,tile.MiRowStart,tile.MiRowEnd,tile.MiColStart,tile.MiColEnd)};
        Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions(),stage));
        frame.Tiles=Array.Empty<OfficeAv1Tile>();Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions(),stage));
    }
    [Fact]
    public void TenBitPlanesCannotEnterUnqualifiedFiltersOrEightBitComposition() {
        var (bytes,sequence,frame)=Read("avif-main10-420-full",false);var options=new OfficeRasterDecodeOptions();
        var result=OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,options);
        Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,options,OfficeAv1ReconstructionStage.Restored));
        Assert.Throws<FormatException>(()=>new OfficeAv1Restorer(frame,sequence,64,options));
        Assert.Throws<FormatException>(()=>OfficeAvifColorConverter.Compose(result,sequence.Color,null,options));
        frame.BitDepth=8;
        Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,options));
    }
    private static JsonDocument Open(int depth=8) {
        using var input=File.OpenRead(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif",depth==10?"reconstruction-main10-reference.json.gz":"reconstruction-reference.json.gz"));
        using var gzip=new GZipStream(input,CompressionMode.Decompress);return JsonDocument.Parse(gzip);
    }
    private static (byte[],OfficeAv1StillSequence,OfficeAv1StillFrame) ReadControl(JsonElement c,JsonElement input,int depth) {
        byte[] bytes=Convert.FromBase64String(input.GetString()!);var native=c.GetProperty("unfiltered");
        var item=new OfficeAvifImageItem(1,native.GetProperty("width").GetInt32(),native.GetProperty("height").GetInt32(),0,bytes.Length,
            new byte[] {0x81,(byte)native.GetProperty("level").GetInt32(),(byte)(depth==10?76:12),0},false,null);
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
