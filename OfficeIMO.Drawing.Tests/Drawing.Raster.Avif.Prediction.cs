using System.Text.Json;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Actual native prediction pixels protect integer filtering, edge extension and CfL behavior.</summary>
public sealed class DrawingAv1PredictionTests {
    [Theory]
    [InlineData(8)]
    [InlineData(10)]
    public void EveryNativePredictionPixelMatchesAcrossModesEdgesAnglesFiltersAndPlanes(int bitDepth) {
        using var fixture=Open(bitDepth);var owner=new OfficeAv1IntraPredictor(bitDepth,new OfficeRasterDecodeOptions());
        var sizes=new HashSet<int>();var modes=new HashSet<int>();var filters=new HashSet<int>();var alphas=new HashSet<int>();
        int scenario=0;
        foreach(var item in fixture.RootElement.GetProperty("cases").EnumerateArray()) {
            int[] p=Ints(item.GetProperty("parameters"));int kind=item.GetProperty("kind").GetInt32();OfficeAv1Prediction actual;
            sizes.Add(p[0]);
            if(kind==0) {
                modes.Add(p[1]);if(p[5]>=0) filters.Add(p[5]);
                var above=Samples(item.GetProperty("above"));var left=Samples(item.GetProperty("left"));var originalTop=(ushort[])above.Clone();var originalLeft=(ushort[])left.Clone();
                actual=owner.Predict(p[0],(OfficeAv1IntraMode)p[1],p[2],new OfficeAv1PredictionEdges(above,left,(ushort)p[10],p[6],p[7],p[8],p[9]),p[3]!=0,p[4]!=0,p[5]);
                Assert.Equal(originalTop,above);Assert.Equal(originalLeft,left);
            } else if(kind==1) {
                alphas.Add(p[1]);actual=owner.PredictChromaFromLuma(p[0],p[1],(ushort)p[2],Samples(item.GetProperty("luma")),p[7],p[5],p[6],p[3],p[4]);
            } else {
                var colors=Samples(item.GetProperty("colors"));var map=Bytes(item.GetProperty("map"));
                ushort[] empty=Array.Empty<ushort>(),zeros=new ushort[colors.Length];byte[] emptyMap=Array.Empty<byte>();bool y=p[1]==0;
                var palette=new OfficeAv1Palette(y?colors:empty,y?empty:p[1]==1?colors:zeros,y?empty:p[1]==2?colors:zeros,y?map:emptyMap,y?emptyMap:map,y?p[2]:2*p[2],y?p[3]:2*p[3],-1);
                actual=owner.PredictPalette(p[0],p[1],palette,p[4],p[5]);
            }
            var expected=Samples(item.GetProperty("pixels"));
            Assert.True(expected.SequenceEqual(Enumerable.Range(0,actual.Count).Select(actual.Value)),"Native prediction mismatch in scenario "+scenario+" kind "+kind+" size "+p[0]);
            Assert.Equal(OfficeAv1TransformSize.Width(p[0]),actual.Width);Assert.Equal(OfficeAv1TransformSize.Height(p[0]),actual.Height);scenario++;
        }
        Assert.Equal(Enumerable.Range(0,19),sizes.OrderBy(x=>x));Assert.Equal(Enumerable.Range(0,13),modes.OrderBy(x=>x));
        Assert.Equal(Enumerable.Range(0,5),filters.OrderBy(x=>x));Assert.Equal(Enumerable.Range(-16,33),alphas.OrderBy(x=>x));
    }
    [Theory]
    [InlineData(8)]
    [InlineData(10)]
    public void ReturnedPredictionsSurviveLaterCallsAndScratchFilteringNeverChangesInputEdges(int bitDepth) {
        var owner=new OfficeAv1IntraPredictor(bitDepth,new OfficeRasterDecodeOptions());ushort[] a={1,20,90,200,255,33,64,128},b={255,0,128,1,200,50,80,100};
        if(bitDepth==10) {for(int i=0;i<a.Length;i++) {a[i]=(ushort)(4*a[i]);b[i]=(ushort)(4*b[i]);}}
        var first=owner.Predict(0,OfficeAv1IntraMode.Diagonal67,3,new OfficeAv1PredictionEdges(a,b,129,4,4,4,4),true,false);
        ushort[] saved=Enumerable.Range(0,first.Count).Select(first.Value).ToArray();a[0]=100;
        owner.Predict(0,OfficeAv1IntraMode.Smooth,0,new OfficeAv1PredictionEdges(a,b,0,4,4),false,true);
        Assert.Equal(saved,Enumerable.Range(0,first.Count).Select(first.Value));
    }
    [Theory]
    [InlineData(8)]
    [InlineData(10)]
    public void InvalidPredictionInputsMemoryPixelsAndCancellationFailBeforePublishingOutput(int bitDepth) {
        var edges=new OfficeAv1PredictionEdges(new ushort[4],new ushort[4],0,4,4);
        Assert.Throws<FormatException>(()=>new OfficeAv1IntraPredictor(bitDepth,new OfficeRasterDecodeOptions {RetainedManagedBytes=long.MaxValue}));
        Assert.Throws<FormatException>(()=>new OfficeAv1IntraPredictor(12,new OfficeRasterDecodeOptions()));
        var pixelBound=new OfficeAv1IntraPredictor(bitDepth,new OfficeRasterDecodeOptions {MaximumDecodedPixels=15});
        Assert.Throws<FormatException>(()=>pixelBound.Predict(0,OfficeAv1IntraMode.Dc,0,edges,false,false));
        var near=new OfficeAv1IntraPredictor(bitDepth,new OfficeRasterDecodeOptions {RetainedManagedBytes=OfficeRasterGuards.MaximumDecodedBytes-12384});
        Assert.Equal(16,near.Predict(0,OfficeAv1IntraMode.Dc,0,edges,false,false).Count);
        Assert.Throws<FormatException>(()=>near.Predict(4,OfficeAv1IntraMode.Dc,0,new OfficeAv1PredictionEdges(new ushort[64],new ushort[64],0,64,64),false,false));
        var owner=new OfficeAv1IntraPredictor(bitDepth,new OfficeRasterDecodeOptions());
        Assert.Throws<FormatException>(()=>owner.Predict(0,OfficeAv1IntraMode.Dc,1,edges,false,false));
        Assert.Throws<FormatException>(()=>owner.Predict(0,OfficeAv1IntraMode.Vertical,0,edges,false,false,0));
        Assert.Throws<FormatException>(()=>owner.Predict(0,OfficeAv1IntraMode.Dc,0,new OfficeAv1PredictionEdges(new ushort[129],new ushort[4],0,4,4),false,false));
        Assert.Throws<FormatException>(()=>owner.Predict(0,OfficeAv1IntraMode.Dc,0,new OfficeAv1PredictionEdges(new ushort[4],new ushort[4],0,2,4,1),false,false));
        Assert.Throws<FormatException>(()=>owner.PredictChromaFromLuma(0,17,0,new ushort[64],8,8,8,1,1));
        Assert.Throws<FormatException>(()=>owner.PredictChromaFromLuma(0,1,0,new ushort[64],8,7,8,1,1));
        var palette=new OfficeAv1Palette(new ushort[2],Array.Empty<ushort>(),Array.Empty<ushort>(),Enumerable.Repeat((byte)2,16).ToArray(),Array.Empty<byte>(),4,4,-1);
        Assert.Throws<FormatException>(()=>owner.PredictPalette(0,0,palette,0,0));
        ushort excessive=(ushort)(1<<bitDepth);
        Assert.Throws<FormatException>(()=>owner.Predict(0,OfficeAv1IntraMode.Dc,0,new OfficeAv1PredictionEdges(new ushort[] {excessive,0,0,0},new ushort[4],0,4,4),true,false));
        Assert.Throws<FormatException>(()=>owner.Predict(0,OfficeAv1IntraMode.Paeth,0,new OfficeAv1PredictionEdges(new ushort[4],new ushort[4],excessive,4,4),false,false));
        var luma=new ushort[64];luma[0]=excessive;
        Assert.Throws<FormatException>(()=>owner.PredictChromaFromLuma(0,1,0,luma,8,8,8,1,1));
        Assert.Throws<FormatException>(()=>owner.PredictChromaFromLuma(0,1,excessive,new ushort[64],8,8,8,1,1));
        var widePalette=new OfficeAv1Palette(new ushort[] {0,excessive},Array.Empty<ushort>(),Array.Empty<ushort>(),new byte[16],Array.Empty<byte>(),4,4,-1);
        Assert.Throws<FormatException>(()=>owner.PredictPalette(0,0,widePalette,0,0));
        using var canceled=new CancellationTokenSource();var cancellable=new OfficeAv1IntraPredictor(bitDepth,new OfficeRasterDecodeOptions {CancellationToken=canceled.Token});canceled.Cancel();
        Assert.Throws<OperationCanceledException>(()=>cancellable.Predict(0,OfficeAv1IntraMode.Dc,0,edges,false,false));
        Assert.Throws<OperationCanceledException>(()=>cancellable.PredictChromaFromLuma(0,1,128,new ushort[64],8,8,8,1,1));
        Assert.Throws<OperationCanceledException>(()=>cancellable.PredictPalette(0,0,palette,0,0));
    }
    private static int[] Ints(JsonElement element)=>element.EnumerateArray().Select(x=>x.GetInt32()).ToArray();
    private static byte[] Bytes(JsonElement element)=>element.EnumerateArray().Select(x=>x.GetByte()).ToArray();
    private static ushort[] Samples(JsonElement element)=>element.EnumerateArray().Select(x=>x.GetUInt16()).ToArray();
    private static JsonDocument Open(int bitDepth) {
        using var stream=File.OpenRead(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif",bitDepth==8?"prediction-reference.json.gz":"prediction-main10-reference.json.gz"));
        using var gzip=new System.IO.Compression.GZipStream(stream,System.IO.Compression.CompressionMode.Decompress);return JsonDocument.Parse(gzip);
    }
}
