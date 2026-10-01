using System.IO.Compression;
using System.Text.Json;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Actual independent frame pixels protect the integration of entropy order, edges and inverse residuals.</summary>
public sealed class DrawingAv1ReconstructionTests {
    [Theory]
    [InlineData("avif-opaque",false)]
    [InlineData("avif-alpha",false)]
    [InlineData("avif-alpha",true)]
    [InlineData("multitile",false)]
    [InlineData("lossless-copy",false)]
    public void FrozenFramesMatchEveryNativePreFilterSample(string name,bool alpha) {
        using var fixture=Open();
        var c=fixture.RootElement.GetProperty("cases").EnumerateArray().Single(v=>v.GetProperty("name").GetString()==name && v.GetProperty("alpha").GetBoolean()==alpha);
        var (bytes,sequence,frame)=c.TryGetProperty("inputBase64",out var input)?ReadControl(c,input):Read(name,alpha);
        var expected=c.GetProperty("unfiltered");var result=OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions());
        if(name=="lossless-copy") {
            Assert.True(expected.GetProperty("skipLeaves").GetInt32()>0);
            Assert.True(expected.GetProperty("copyLeaves").GetInt32()>0);
            Assert.True(expected.GetProperty("fractionalCopyLeaves").GetInt32()>0);
        }
        Assert.Equal(frame.Width,result.Width);Assert.Equal(frame.Height,result.Height);Assert.Equal(expected.GetProperty("planes").GetArrayLength(),result.PlaneCount);
        for(int p=0;p<result.PlaneCount;p++) {
            byte[] pixels=Convert.FromBase64String(expected.GetProperty("planes")[p].GetString()!);
            int sub=p==0?0:1,w=frame.MiCols*4>>sub,h=frame.MiRows*4>>sub;Assert.Equal(w*h,pixels.Length);
            for(int y=0;y<h;y++) for(int x=0;x<w;x++)
                Assert.True(pixels[y*w+x]==result.Value(p,x,y),$"{name}/{alpha}, plane {p} ({x},{y}): native {pixels[y*w+x]}, managed {result.Value(p,x,y)}");
        }
    }
    [Fact]
    public void FrameFailureCannotPublishPartialPlanesAndLimitsRejectBeforeAllocation() {
        var (bytes,sequence,frame)=Read("avif-opaque",false);
        Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions {RetainedManagedBytes=long.MaxValue}));
        Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions {RetainedManagedBytes=OfficeRasterGuards.MaximumDecodedBytes-500000}));
        Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions {MaximumDecodedPixels=frame.Width*frame.Height-1}));
        Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions {MaximumEncodedBytes=bytes.Length-1}));
        using var cancel=new CancellationTokenSource();cancel.Cancel();
        Assert.Throws<OperationCanceledException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions {CancellationToken=cancel.Token}));
        var tile=frame.Tiles[0];frame.Tiles=new[] {new OfficeAv1Tile(tile.Offset,tile.Length-1,tile.MiRowStart,tile.MiRowEnd,tile.MiColStart,tile.MiColEnd)};
        Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions()));
        frame.Tiles=Array.Empty<OfficeAv1Tile>();Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions()));
    }
    private static JsonDocument Open() {
        using var input=File.OpenRead(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif","reconstruction-reference.json.gz"));
        using var gzip=new GZipStream(input,CompressionMode.Decompress);return JsonDocument.Parse(gzip);
    }
    private static (byte[],OfficeAv1StillSequence,OfficeAv1StillFrame) ReadControl(JsonElement c,JsonElement input) {
        byte[] bytes=Convert.FromBase64String(input.GetString()!);var native=c.GetProperty("unfiltered");
        var item=new OfficeAvifImageItem(1,native.GetProperty("width").GetInt32(),native.GetProperty("height").GetInt32(),0,bytes.Length,
            new byte[] {0x81,(byte)native.GetProperty("level").GetInt32(),12,0},false,null);
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
