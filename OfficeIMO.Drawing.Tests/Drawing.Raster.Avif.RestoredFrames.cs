using System.IO.Compression;
using System.Text.Json;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Independent full-frame output protects restoration units, stripe borders, rounding and cropped edges.</summary>
public sealed class DrawingAv1RestoredFrameTests {
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
    [InlineData("restore-control-0",false)]
    [InlineData("restore-control-1",false)]
    [InlineData("restore-control-2",false)]
    [InlineData("restore-control-3",false)]
    [InlineData("restore-control-4",false)]
    [InlineData("restore-control-5",false)]
[InlineData("restore-control-6",false)]
[InlineData("restore-control-7",false)]
[InlineData("restore-control-8",false)]
[InlineData("restore-control-9",false)]
[InlineData("restore-control-10",false)]
[InlineData("restore-control-11",false)]
[InlineData("restore-control-12",false)]
[InlineData("restore-control-13",false)]
[InlineData("restore-control-14",false)]
[InlineData("restore-control-15",false)]
    public void CroppedFrameMatchesNativeBeforeAndAfterRestoration(string name,bool alpha) {
        using var fixture=Open();var c=fixture.RootElement.GetProperty("cases").EnumerateArray().Single(v=>
            v.GetProperty("name").GetString()==name && v.GetProperty("alpha").GetBoolean()==alpha);
        var (bytes,sequence,frame)=Read(c,name,alpha);
        foreach(var stage in new[] {OfficeAv1ReconstructionStage.Cdef,OfficeAv1ReconstructionStage.Upscaled,OfficeAv1ReconstructionStage.Restored}) {
            var result=OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions(),stage);
            var native=c.GetProperty(stage==OfficeAv1ReconstructionStage.Cdef?"cdef":stage==OfficeAv1ReconstructionStage.Upscaled?"upscaled":"restored");
            var planes=native.GetProperty("planes");Assert.Equal(result.PlaneCount,planes.GetArrayLength());
            for(int p=0;p<result.PlaneCount;p++) {
                int sub=p==0?0:1,w=(result.Width+(1<<sub)-1)>>sub,h=(frame.Height+(1<<sub)-1)>>sub;
                int pitch=stage==OfficeAv1ReconstructionStage.Cdef?frame.MiCols*4>>sub:w;
                byte[] samples=Convert.FromBase64String(planes[p].GetString()!);
                Assert.Equal(stage==OfficeAv1ReconstructionStage.Cdef?pitch*(frame.MiRows*4>>sub):w*h,samples.Length);
                for(int y=0;y<h;y++)for(int x=0;x<w;x++)
                    Assert.True(samples[y*pitch+x]==result.Value(p,x,y),$"{name}/{alpha} {stage}, plane {p} ({x},{y}): native {samples[y*pitch+x]}, managed {result.Value(p,x,y)}");
            }
        }
    }
    [Fact]
    public void KernelsMatchNativeForAllSelfGuidedSetsAndWienerSignedExtremes() {
        using var fixture=Open();var filter=new OfficeAv1RestorationFilter(default);
        foreach(var c in fixture.RootElement.GetProperty("kernels").EnumerateArray()) {
            var taps=c.GetProperty("taps");int type=c.GetProperty("type").GetInt32();
            var unit=new OfficeAv1RestorationUnit(0,0,0,type,c.GetProperty("set").GetInt32(),
                taps[0][0].GetInt32(),taps[0][1].GetInt32(),taps[0][2].GetInt32(),taps[1][0].GetInt32(),taps[1][1].GetInt32(),taps[1][2].GetInt32(),
                c.GetProperty("x0").GetInt32(),c.GetProperty("x1").GetInt32());
            byte[] input=Hex(c.GetProperty("input").GetString()!),output=(byte[])input.Clone();
            filter.Apply(input,input,output,76,76,76,0,75,6,6,c.GetProperty("width").GetInt32(),c.GetProperty("height").GetInt32(),unit);
            Assert.True(output.SequenceEqual(Hex(c.GetProperty("output").GetString()!)),
                $"Restoration kernel type {type}, set {unit.SgrSet}, weights {unit.X0}/{unit.X1}, taps {taps}");
        }
    }
    [Fact]
    public void RestorationSnapshotsShareTheAggregateRetainedBudget() {
        using var fixture=Open();var c=fixture.RootElement.GetProperty("cases").EnumerateArray().First(v=>
            v.GetProperty("restored").GetProperty("restoration").EnumerateArray().Any(u=>u.GetProperty("type").GetInt32()!=0));
        string name=c.GetProperty("name").GetString()!;var (bytes,sequence,frame)=Read(c,name,c.GetProperty("alpha").GetBoolean());
        int sb=sequence.Use128Superblock?128:64;long storage=(long)((frame.Width+sb-1)/sb*sb)*((frame.Height+sb-1)/sb*sb)*(sequence.Monochrome?2:3)/2+
            (long)frame.MiRows*frame.MiCols*2;
        long cdef=sequence.Cdef && !frame.CodedLossless && !frame.AllowIntraBlockCopy?OfficeAv1Cdef.ContextBytes(frame,sequence.Monochrome,sb):0;
        long reserve=storage+131072+OfficeAv1IntraPredictor.ContextBytes+OfficeAv1ResidualTransform.ContextBytes+
            OfficeAv1Deblocker.ContextBytes(frame,sequence.Monochrome)+cdef+frame.Tiles.Max(t=>OfficeAv1TileReader.ContextBytes(frame,t))+bytes.LongLength;
        var options=new OfficeRasterDecodeOptions {RetainedManagedBytes=OfficeRasterGuards.MaximumDecodedBytes-reserve-8192};
        Assert.NotNull(OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,options,OfficeAv1ReconstructionStage.Cdef));
        Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,options,OfficeAv1ReconstructionStage.Restored));
    }
    [Fact]
    public void FilterRejectsOutputFeedbackAndObservesCancellationBeforeWriting() {
        using var fixture=Open();var c=fixture.RootElement.GetProperty("kernels").EnumerateArray().First(v=>v.GetProperty("type").GetInt32()==3);
        byte[] source=Hex(c.GetProperty("input").GetString()!),original=(byte[])source.Clone();
        var unit=new OfficeAv1RestorationUnit(0,0,0,3,c.GetProperty("set").GetInt32(),0,0,0,0,0,0,
            c.GetProperty("x0").GetInt32(),c.GetProperty("x1").GetInt32());
        var filter=new OfficeAv1RestorationFilter(default);
        Assert.Throws<FormatException>(()=>filter.Apply(source,source,source,76,76,76,0,75,6,6,17,9,unit));
        Assert.Equal(original,source);
        using var cancellation=new CancellationTokenSource();filter=new OfficeAv1RestorationFilter(cancellation.Token);
        byte[] output=(byte[])source.Clone();cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(()=>filter.Apply(source,source,output,76,76,76,0,75,6,6,17,9,unit));
        Assert.Equal(original,output);
    }
    private static JsonDocument Open() {
        using var input=File.OpenRead(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif","restored-reference.json.gz"));
        using var gzip=new GZipStream(input,CompressionMode.Decompress);return JsonDocument.Parse(gzip);
    }
    private static byte[] Hex(string value)=>Enumerable.Range(0,value.Length/2).Select(i=>Convert.ToByte(value.Substring(i*2,2),16)).ToArray();
    private static (byte[],OfficeAv1StillSequence,OfficeAv1StillFrame) Read(JsonElement c,string name,bool alpha) {
        byte[] bytes;OfficeAvifImageItem item;var options=new OfficeRasterDecodeOptions();
        if(c.TryGetProperty("inputBase64",out var input)) {
            bytes=Convert.FromBase64String(input.GetString()!);var native=c.GetProperty("cdef");bool mono=c.GetProperty("monochrome").GetBoolean();
            item=new OfficeAvifImageItem(1,c.GetProperty("restored").GetProperty("width").GetInt32(),native.GetProperty("height").GetInt32(),0,bytes.Length,
                new byte[] {0x81,(byte)native.GetProperty("level").GetInt32(),(byte)(mono?28:12),0},mono,null);
        } else {
            bytes=File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif",name+".avif"));
            Assert.True(OfficeAvifContainerReader.TryRead(bytes,options,out var container));item=alpha?container!.Alpha!:container!.Color;
        }
        Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes,item,options,out var sequence));
        Assert.True(OfficeAv1StillFrameReader.TryRead(bytes,item,sequence!,options,out var frame));return(bytes,sequence!,frame!);
    }
}
