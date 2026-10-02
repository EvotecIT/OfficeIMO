using System.Security.Cryptography;
using System.Text.Json;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Actual native inverse samples protect fixed-point reconstruction, not a self-round-trip.</summary>
public sealed class DrawingAv1ResidualTests {
    [Theory]
    [InlineData(8)]
    [InlineData(10)]
    public void EveryNativeResidualAndClippedSampleMatchesAcrossSizesTypesAndQuantization(int bitDepth) {
        using var fixture=Open(bitDepth);int cases=0;var sizes=new HashSet<int>();var types=new HashSet<int>();
        foreach(var item in fixture.RootElement.GetProperty("cases").EnumerateArray()) {
            int[] p=Ints(item.GetProperty("parameters"));var (frame,prelude,coefficients)=Input(item,bitDepth);
            sizes.Add(p[0]);types.Add(p[1]);
            var reader=new OfficeAv1ResidualTransform(frame,new OfficeRasterDecodeOptions());
            if(item.TryGetProperty("conforming",out var conforming) && !conforming.GetBoolean()) {
                Assert.Throws<FormatException>(()=>reader.Reconstruct(coefficients,prelude));cases++;continue;
            }
            var actual=reader.Reconstruct(coefficients,prelude);
            int[] expected=Ints(item.GetProperty("residual"));
            Assert.True(expected.SequenceEqual(Enumerable.Range(0,actual.Count).Select(actual.Value)),"Native residual mismatch in case "+cases+" size/type "+p[0]+"/"+p[1]);
            Assert.Equal(Ints(item.GetProperty("pixels")),Enumerable.Range(0,actual.Count).Select(i=>Math.Max(0,Math.Min((1<<bitDepth)-1,(1<<(bitDepth-1))+actual.Value(i)))));
            Assert.Equal(OfficeAv1TransformSize.Width(p[0]),actual.Width);Assert.Equal(OfficeAv1TransformSize.Height(p[0]),actual.Height);
            cases++;
        }
        Assert.Equal(Enumerable.Range(0,19),sizes.OrderBy(x=>x));Assert.Equal(Enumerable.Range(0,16),types.OrderBy(x=>x));
    }
    [Theory]
    [InlineData(8)]
    [InlineData(10)]
    public void QuantizationFactsMatchEveryNativeLookupAndMatrixElement(int bitDepth) {
        using var fixture=Open(bitDepth);var facts=fixture.RootElement.GetProperty("numericFacts");
        Assert.Equal(Ints(facts.GetProperty("dc")),(bitDepth==8?OfficeAv1QuantizationTables.Dc:OfficeAv1QuantizationTables.Dc10).Select(x=>(int)x));
        Assert.Equal(Ints(facts.GetProperty("ac")),(bitDepth==8?OfficeAv1QuantizationTables.Ac:OfficeAv1QuantizationTables.Ac10).Select(x=>(int)x));
        int[] offsets={0,16,80,336,336,1360,1392,1424,1552,1680,2192,336,336,2704,2768,2832,3088,1680,2192};
        var bytes=new byte[OfficeAv1QuantizationTables.MatrixBytes];
        for(int level=0;level<15;level++) for(int plane=0;plane<2;plane++) for(int size=0;size<19;size++) {
            int count=Math.Min(32,OfficeAv1TransformSize.Width(size))*Math.Min(32,OfficeAv1TransformSize.Height(size));
            for(int i=0;i<count;i++) bytes[(level*2+plane)*3344+offsets[size]+i]=(byte)OfficeAv1QuantizationTables.Weight(level,plane!=0,size,i);
        }
        using var hash=SHA256.Create();Assert.Equal(fixture.RootElement.GetProperty("matrixSha256").GetString(),BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-","").ToLowerInvariant());
    }
    [Theory]
    [InlineData(8)]
    [InlineData(10)]
    public void QuantizationSnapshotAndReturnedResidualsSurviveLaterCallsAndFrameChanges(int bitDepth) {
        using var fixture=Open(bitDepth);var item=fixture.RootElement.GetProperty("cases")[3];var (frame,prelude,coefficients)=Input(item,bitDepth);
        var reader=new OfficeAv1ResidualTransform(frame,new OfficeRasterDecodeOptions());
        frame.BitDepth=bitDepth==8?10:8;frame.BaseQIndex=255;frame.DeltaQYDc=63;frame.QMatrixLevels[0]=15;frame.SegmentFeatures[prelude.SegmentId,0]=true;frame.SegmentData[prelude.SegmentId,0]=255;
        var first=reader.Reconstruct(coefficients,prelude);var second=reader.Reconstruct(coefficients,prelude);
        Assert.Equal(Ints(item.GetProperty("residual")),Enumerable.Range(0,first.Count).Select(first.Value));
        Assert.Equal(Ints(item.GetProperty("residual")),Enumerable.Range(0,second.Count).Select(second.Value));
    }
    [Theory]
    [InlineData(8)]
    [InlineData(10)]
    public void ResidualMemoryPixelCoefficientAndCancellationBoundsRemainObservable(int bitDepth) {
        using var fixture=Open(bitDepth);var small=fixture.RootElement.GetProperty("cases")[0];var (frame,prelude,coefficients)=Input(small,bitDepth);
        frame.BitDepth=12;Assert.Throws<FormatException>(()=>new OfficeAv1ResidualTransform(frame,new OfficeRasterDecodeOptions()));frame.BitDepth=bitDepth;
        Assert.Throws<FormatException>(()=>new OfficeAv1ResidualTransform(frame,new OfficeRasterDecodeOptions {RetainedManagedBytes=long.MaxValue}));
        Assert.Throws<FormatException>(()=>new OfficeAv1ResidualTransform(frame,new OfficeRasterDecodeOptions {MaximumDecodedPixels=100}));
        var options=new OfficeRasterDecodeOptions {RetainedManagedBytes=OfficeRasterGuards.MaximumDecodedBytes-110000};
        Assert.Equal(Ints(small.GetProperty("residual")),Enumerable.Range(0,16).Select(new OfficeAv1ResidualTransform(frame,options).Reconstruct(coefficients,prelude).Value));
        var large=fixture.RootElement.GetProperty("cases").EnumerateArray().First(c=>c.GetProperty("parameters")[0].GetInt32()==4);
        var input=Input(large,bitDepth);Assert.Throws<FormatException>(()=>new OfficeAv1ResidualTransform(input.Item1,options).Reconstruct(input.Item3,input.Item2));
        var malformed=new OfficeAv1Coefficients(coefficients.Block,coefficients.Type,1,new[] {0x100000,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0});
        Assert.Throws<FormatException>(()=>new OfficeAv1ResidualTransform(frame,new OfficeRasterDecodeOptions()).Reconstruct(malformed,prelude));
        using var cancellation=new CancellationTokenSource();var reader=new OfficeAv1ResidualTransform(frame,new OfficeRasterDecodeOptions {CancellationToken=cancellation.Token});cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(()=>reader.Reconstruct(coefficients,prelude));
    }
    private static (OfficeAv1StillFrame,OfficeAv1BlockPrelude,OfficeAv1Coefficients) Input(JsonElement item,int bitDepth) {
        int[] p=Ints(item.GetProperty("parameters"));
        var frame=new OfficeAv1StillFrame {Width=64,Height=64,BitDepth=bitDepth,BaseQIndex=p[3],DeltaQPresent=p[5]!=0,SegmentationEnabled=p[7]!=0,UsingQMatrix=p[10]!=0,
            DeltaQYDc=p[14],DeltaQUDc=p[15],DeltaQUAc=p[16],DeltaQVDc=p[17],DeltaQVAc=p[18]};
        for(int i=0;i<3;i++) frame.QMatrixLevels[i]=p[11+i];frame.SegmentFeatures[p[6],0]=p[8]!=0;frame.SegmentData[p[6],0]=p[9];
        var prelude=new OfficeAv1BlockPrelude(p[19]!=0,p[6],p[20]!=0,0,p[4],new int[4]);
        var coefficients=new OfficeAv1Coefficients(new OfficeAv1TransformBlock(p[2],0,0,p[0]),(OfficeAv1TransformType)p[1],p[21],Ints(item.GetProperty("coefficients")));
        return (frame,prelude,coefficients);
    }
    private static int[] Ints(JsonElement value)=>value.EnumerateArray().Select(x=>x.GetInt32()).ToArray();
    private static JsonDocument Open(int bitDepth) {
        using var stream=File.OpenRead(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif",bitDepth==8?"residual-reference.json.gz":"residual-main10-reference.json.gz"));
        using var gzip=new System.IO.Compression.GZipStream(stream,System.IO.Compression.CompressionMode.Decompress);return JsonDocument.Parse(gzip);
    }
}
