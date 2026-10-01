using System.IO.Compression;
using System.Text.Json;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Actual libavif RGBA output protects matrix/range conversion, odd chroma edges and straight alpha.</summary>
public sealed class DrawingAvifColorTests {
    [Fact]
    public void NativeBilinearColorRowsMatchForEveryMatrixRangeAndAlphaControl() {
        using var file=File.OpenRead(Asset("color-reference.json.gz"));using var gzip=new GZipStream(file,CompressionMode.Decompress);
        using var fixture=JsonDocument.Parse(gzip);
        foreach(var c in fixture.RootElement.GetProperty("cases").EnumerateArray()) {
            int w=c.GetProperty("width").GetInt32(),h=c.GetProperty("height").GetInt32();
            var planes=c.GetProperty("planes").EnumerateArray().Select(p=>Hex(p.GetString()!)).ToArray();
            var frame=Padded(w,h,planes);var a=c.GetProperty("alpha");
            var alpha=a.ValueKind==JsonValueKind.Null?null:Padded(w,h,new[] {Hex(a.GetString()!)});
            byte[] expected=Hex(c.GetProperty("rgba").GetString()!);
            var metadata=new OfficeAvifColorDescription(2,2,c.GetProperty("matrix").GetInt32(),c.GetProperty("fullRange").GetBoolean());
            var image=OfficeAvifColorConverter.Compose(frame,metadata,alpha,new OfficeRasterDecodeOptions());
            Compare(expected,image.PixelBuffer,1,$"matrix {metadata.Matrix}, full {metadata.FullRange}, {w}x{h}");
        }
    }

    [Theory]
    [InlineData("avif-opaque")]
    [InlineData("avif-alpha")]
    public void FrozenFullAvifMatchesIndependentPixelsIncludingExactAlpha(string name) {
        byte[] bytes=File.ReadAllBytes(Asset(name+".avif"));var options=new OfficeRasterDecodeOptions();
        Assert.True(OfficeAvifContainerReader.TryRead(bytes,options,out var container));
        var (color,description)=Decode(bytes,container!.Color,options);
        OfficeAv1ReconstructedFrame? alpha=container.Alpha==null?null:Decode(bytes,container.Alpha,options).Item1;
        var image=OfficeAvifColorConverter.Compose(color,container.Color.ColorDescription??description,alpha,options);
        Compare(File.ReadAllBytes(Asset(name+".rgba")),image.PixelBuffer,3,name);
        if(alpha!=null)Assert.Contains(image.PixelBuffer.Where((v,i)=>i%4==3),v=>v>0 && v<255);
    }

    [Fact]
    public void CompositionRejectsUnsafeDimensionsUnsupportedMatricesAndCanceledWork() {
        var color=Padded(3,5,new[] {new byte[15],new byte[6],new byte[6]});
        var metadata=new OfficeAvifColorDescription(2,2,6,true);
        Assert.Throws<FormatException>(()=>OfficeAvifColorConverter.Compose(color,metadata,null,new OfficeRasterDecodeOptions {MaximumDecodedPixels=14}));
        Assert.Throws<FormatException>(()=>OfficeAvifColorConverter.Compose(color,metadata,null,new OfficeRasterDecodeOptions {MaximumInspectionWorkPixels=14}));
        Assert.Throws<FormatException>(()=>OfficeAvifColorConverter.Compose(color,metadata,null,new OfficeRasterDecodeOptions {
            RetainedManagedBytes=OfficeRasterGuards.MaximumDecodedBytes-color.StorageBytes-60+1}));
        Assert.Throws<NotSupportedException>(()=>OfficeAvifColorConverter.Compose(color,new OfficeAvifColorDescription(2,2,10,true),null,new OfficeRasterDecodeOptions()));
        Assert.Throws<FormatException>(()=>OfficeAvifColorConverter.Compose(color,metadata,Padded(1,1,new[] {new byte[1]}),new OfficeRasterDecodeOptions()));
        using var cancellation=new CancellationTokenSource();cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(()=>OfficeAvifColorConverter.Compose(color,metadata,null,new OfficeRasterDecodeOptions {CancellationToken=cancellation.Token}));
        var mono=Padded(1,1,new[] {new byte[] {128}});
        Assert.Equal(new byte[] {128,128,128,255},OfficeAvifColorConverter.Compose(mono,metadata,null,new OfficeRasterDecodeOptions()).PixelBuffer);
    }

    private static (OfficeAv1ReconstructedFrame,OfficeAvifColorDescription) Decode(byte[] bytes,OfficeAvifImageItem item,OfficeRasterDecodeOptions options) {
        Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes,item,options,out var sequence));
        Assert.True(OfficeAv1StillFrameReader.TryRead(bytes,item,sequence!,options,out var frame));
        return (OfficeAv1FrameReconstructor.Decode(bytes,sequence!,frame!,options,OfficeAv1ReconstructionStage.Restored),sequence!.Color);
    }
    private static OfficeAv1ReconstructedFrame Padded(int width,int height,byte[][] source) {
        int stride=(width+63)/64*64,rows=(height+63)/64*64;var planes=new ushort[source.Length][];
        for(int p=0;p<source.Length;p++) {
            int sub=p==0?0:1,w=(width+sub)>>sub,h=(height+sub)>>sub,pitch=stride>>sub;
            planes[p]=new ushort[pitch*(rows>>sub)];for(int y=0;y<h;y++)for(int x=0;x<w;x++)planes[p][y*pitch+x]=source[p][y*w+x];
        }
        return new OfficeAv1ReconstructedFrame(width,height,stride,rows,8,planes);
    }
    private static void Compare(byte[] expected,byte[] actual,int tolerance,string name) {
        Assert.Equal(expected.Length,actual.Length);
        for(int i=0;i<expected.Length;i++)Assert.True(Math.Abs(expected[i]-actual[i])<=(i%4==3?0:tolerance),
            $"{name}, pixel {i/4}, channel {i%4}: native {expected[i]}, managed {actual[i]}");
    }
    private static string Asset(string name)=>Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif",name);
    private static byte[] Hex(string value)=>Enumerable.Range(0,value.Length/2).Select(i=>Convert.ToByte(value.Substring(i*2,2),16)).ToArray();
}
