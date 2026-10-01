using System.IO.Compression;
using System.Text.Json;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Independent native row output protects phases, MI-edge clamping and odd chroma dimensions.</summary>
public sealed class DrawingAv1SuperresolutionTests {
    [Fact]
    public void NativeRowsMatchForEveryScaleAcrossOddWidthsChromaAndTileBoundaries() {
        using var file=File.OpenRead(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif","restored-reference.json.gz"));
        using var gzip=new GZipStream(file,CompressionMode.Decompress);using var fixture=JsonDocument.Parse(gzip);
        foreach(var c in fixture.RootElement.GetProperty("upscaleKernels").EnumerateArray()) {
            int coded=c.GetProperty("codedWidth").GetInt32(),width=c.GetProperty("width").GetInt32(),height=c.GetProperty("height").GetInt32();
            int plane=c.GetProperty("plane").GetInt32(),sub=plane==0?0:1,sourcePitch=c.GetProperty("inputStride").GetInt32();
            var frame=new OfficeAv1StillFrame {Width=coded,UpscaledWidth=width,Height=height,
                MiCols=2*((coded+7)/8),MiRows=2*((height+7)/8),SuperResolutionDenominator=c.GetProperty("denominator").GetInt32()};
            var sequence=new OfficeAv1StillSequence {SuperResolution=true,Monochrome=plane==0};
            int stride=(coded+63)/64*64;var input=new byte[sequence.Monochrome?1:3][];
            for(int p=0;p<input.Length;p++)input[p]=new byte[(stride>>(p==0?0:1))*(64>>(p==0?0:1))];
            byte[] samples=Hex(c.GetProperty("input").GetString()!),expected=Hex(c.GetProperty("output").GetString()!);
            int rows=(height+sub)>>sub,columns=frame.MiCols*4>>sub,pitch=stride>>sub,outWidth=(width+sub)>>sub;
            for(int y=0;y<rows;y++)Buffer.BlockCopy(samples,y*sourcePitch,input[plane],y*pitch,columns);
            byte[] saved=(byte[])input[plane].Clone();var owner=new OfficeAv1Upscaler(frame,sequence,64,new OfficeRasterDecodeOptions());
            byte[][] result=owner.Apply(input,stride);Assert.Equal(saved,input[plane]);
            for(int y=0;y<rows;y++)for(int x=0;x<outWidth;x++)
                Assert.True(expected[y*outWidth+x]==result[plane][y*(owner.Stride>>sub)+x],
                    $"Scale {frame.SuperResolutionDenominator}, width {width}, plane {plane}, tiles {c.GetProperty("tiles")}, ({x},{y})");
        }
    }
    [Fact]
    public void OutputDimensionsAndRepeatedUpscalingSharePixelMemoryWorkAndCancellationBounds() {
        var frame=new OfficeAv1StillFrame {Width=17,UpscaledWidth=33,Height=3,MiCols=6,MiRows=2,SuperResolutionDenominator=16};
        var sequence=new OfficeAv1StillSequence {SuperResolution=true,Monochrome=true};
        Assert.Throws<FormatException>(()=>new OfficeAv1Upscaler(frame,sequence,64,new OfficeRasterDecodeOptions {MaximumDecodedPixels=98}));
        Assert.Throws<FormatException>(()=>new OfficeAv1Upscaler(frame,sequence,64,new OfficeRasterDecodeOptions {
            RetainedManagedBytes=OfficeRasterGuards.MaximumDecodedBytes-OfficeAv1Upscaler.ContextBytes(frame,true,64)+1}));
        var source=new[] {new byte[64*64]};byte[] saved=(byte[])source[0].Clone();
        var limited=new OfficeAv1Upscaler(frame,sequence,64,new OfficeRasterDecodeOptions {MaximumInspectionWorkPixels=98});
        Assert.Throws<FormatException>(()=>limited.Apply(source,64));
        var cumulative=new OfficeAv1Upscaler(frame,sequence,64,new OfficeRasterDecodeOptions {MaximumInspectionWorkPixels=99});
        Assert.NotNull(cumulative.Apply(source,64));Assert.Throws<FormatException>(()=>cumulative.Apply(source,64));
        using var cancellation=new CancellationTokenSource();var owner=new OfficeAv1Upscaler(frame,sequence,64,new OfficeRasterDecodeOptions {CancellationToken=cancellation.Token});
        cancellation.Cancel();Assert.Throws<OperationCanceledException>(()=>owner.Apply(source,64));Assert.Equal(saved,source[0]);
    }
    private static byte[] Hex(string value)=>Enumerable.Range(0,value.Length/2).Select(i=>Convert.ToByte(value.Substring(i*2,2),16)).ToArray();
}
