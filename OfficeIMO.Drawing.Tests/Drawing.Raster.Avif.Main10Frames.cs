using System.IO.Compression;
using System.Text.Json;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Independent ten-bit frame samples protect frame-filter ordering, geometry and padded source edges.</summary>
public sealed class DrawingAv1Main10FrameTests {
    [Theory]
    [InlineData("avif-main10-420-full",false)]
    [InlineData("avif-main10-420-limited",false)]
    [InlineData("avif-main10-420-full-alpha",false)]
    [InlineData("avif-main10-420-full-alpha",true)]
    [InlineData("avif-main10-420-limited-alpha",false)]
    [InlineData("avif-main10-420-limited-alpha",true)]
    [InlineData("avif-main10-mono-full",false)]
    [InlineData("avif-main10-mono-limited",false)]
    [InlineData("avif-main10-mono-full-alpha",false)]
    [InlineData("avif-main10-mono-full-alpha",true)]
    [InlineData("avif-main10-mono-limited-alpha",false)]
    [InlineData("avif-main10-mono-limited-alpha",true)]
    [InlineData("lossless-copy",false)]
    [InlineData("odd-color",false)]
    [InlineData("odd-mono",false)]
    [InlineData("odd-tiled-color",false)]
    [InlineData("upscale-control-0",false)]
    [InlineData("upscale-control-1",false)]
    [InlineData("upscale-control-2",false)]
    [InlineData("upscale-control-3",false)]
    [InlineData("upscale-control-4",false)]
    [InlineData("upscale-control-5",false)]
    [InlineData("upscale-control-6",false)]
    [InlineData("upscale-control-7",false)]
    [InlineData("upscale-control-8",false)]
    [InlineData("upscale-control-9",false)]
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
    [InlineData("restore-control-16",false)]
    [InlineData("restore-control-17",false)]
    [InlineData("restore-control-18",false)]
    [InlineData("restore-control-19",false)]
    public void CompleteFramesMatchNativeCdefUpscaledAndRestoredSamples(string name,bool alpha) {
        using var fixture=Open();var c=Find(fixture,name,alpha);var (bytes,sequence,frame)=Read(c);
        if(name.StartsWith("upscale-control-",StringComparison.Ordinal)) {
            Assert.True(frame.Width<frame.UpscaledWidth);
            Assert.Equal(4,frame.Tiles.Length);
            Assert.Equal(name.EndsWith("-8") || name.EndsWith("-9"),sequence.Use128Superblock);
        }
        foreach(var stage in new[] {OfficeAv1ReconstructionStage.Cdef,OfficeAv1ReconstructionStage.Upscaled,OfficeAv1ReconstructionStage.Restored}) {
            var result=OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,new OfficeRasterDecodeOptions(),stage);
            var native=c.GetProperty(stage==OfficeAv1ReconstructionStage.Cdef?"cdef":stage==OfficeAv1ReconstructionStage.Upscaled?"upscaled":"restored");
            Assert.Equal(10,result.BitDepth);Assert.Equal(native.GetProperty("width").GetInt32(),result.Width);
            Assert.Equal(native.GetProperty("height").GetInt32(),result.Height);
            var planes=native.GetProperty("planes");Assert.Equal(planes.GetArrayLength(),result.PlaneCount);
            for(int p=0;p<result.PlaneCount;p++) {
                int sub=p==0?0:1,w=(result.Width+sub)>>sub,h=(result.Height+sub)>>sub;
                int pitch=stage==OfficeAv1ReconstructionStage.Cdef?frame.MiCols*4>>sub:w;
                byte[] expected=Convert.FromBase64String(planes[p].GetString()!);
                Assert.Equal(2*pitch*(stage==OfficeAv1ReconstructionStage.Cdef?frame.MiRows*4>>sub:h),expected.Length);
                for(int y=0;y<h;y++)for(int x=0;x<w;x++) {
                    int offset=2*(y*pitch+x),sample=expected[offset]+expected[offset+1]*256;
                    Assert.True(sample==result.Value(p,x,y),$"{name}/{alpha} {stage}, plane {p} ({x},{y}): native {sample}, managed {result.Value(p,x,y)}");
                }
            }
        }
    }
    [Fact]
    public void UpscaledOutputSharesAggregateMemoryWithLiveCodedPlanesAndFilters() {
        using var fixture=Open();var (bytes,sequence,frame)=Read(Find(fixture,"upscale-control-0",false));int sb=sequence.Use128Superblock?128:64;
        long storage=(long)((frame.Width+sb-1)/sb*sb)*((frame.Height+sb-1)/sb*sb)*(sequence.Monochrome?2:3)+(long)frame.MiRows*frame.MiCols*2;
        long cdef=sequence.Cdef && !frame.CodedLossless && !frame.AllowIntraBlockCopy?OfficeAv1Cdef.ContextBytes(frame,sequence.Monochrome,sb):0;
        long reserve=storage+131072+OfficeAv1IntraPredictor.ContextBytes+OfficeAv1ResidualTransform.ContextBytes+
            OfficeAv1Deblocker.ContextBytes(frame,sequence.Monochrome)+cdef+frame.Tiles.Max(t=>OfficeAv1TileReader.ContextBytes(frame,t))+bytes.LongLength;
        var options=new OfficeRasterDecodeOptions {RetainedManagedBytes=OfficeRasterGuards.MaximumDecodedBytes-reserve-8192};
        Assert.NotNull(OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,options,OfficeAv1ReconstructionStage.Cdef));
        Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,options,OfficeAv1ReconstructionStage.Upscaled));
    }
    [Fact]
    public void RestoredOutputReservesSnapshotsAlongsideLiveUpscaledPlanes() {
        using var fixture=Open();var (bytes,sequence,frame)=Read(Find(fixture,"restore-control-6",false));
        Assert.True(frame.Width<frame.UpscaledWidth);
        Assert.Contains(frame.RestorationTypes,type=>type!=0);
        int sb=sequence.Use128Superblock?128:64;
        long storage=(long)((frame.Width+sb-1)/sb*sb)*((frame.Height+sb-1)/sb*sb)*(sequence.Monochrome?2:3)+(long)frame.MiRows*frame.MiCols*2;
        long cdef=sequence.Cdef && !frame.CodedLossless && !frame.AllowIntraBlockCopy?OfficeAv1Cdef.ContextBytes(frame,sequence.Monochrome,sb):0;
        long reserve=storage+131072+OfficeAv1IntraPredictor.ContextBytes+OfficeAv1ResidualTransform.ContextBytes+
            OfficeAv1Deblocker.ContextBytes(frame,sequence.Monochrome)+cdef+OfficeAv1Upscaler.ContextBytes(frame,sequence.Monochrome,sb)+
            frame.Tiles.Max(t=>OfficeAv1TileReader.ContextBytes(frame,t))+bytes.LongLength;
        var options=new OfficeRasterDecodeOptions {RetainedManagedBytes=OfficeRasterGuards.MaximumDecodedBytes-reserve-8192};
        Assert.NotNull(OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,options,OfficeAv1ReconstructionStage.Upscaled));
        Assert.Throws<FormatException>(()=>OfficeAv1FrameReconstructor.Decode(bytes,sequence,frame,options,OfficeAv1ReconstructionStage.Restored));
    }
    private static JsonDocument Open() {
        using var file=File.OpenRead(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif","restored-main10-reference.json.gz"));
        using var gzip=new GZipStream(file,CompressionMode.Decompress);return JsonDocument.Parse(gzip);
    }
    private static JsonElement Find(JsonDocument fixture,string name,bool alpha)=>fixture.RootElement.GetProperty("cases").EnumerateArray()
        .Single(c=>c.GetProperty("name").GetString()==name && c.GetProperty("alpha").GetBoolean()==alpha);
    private static (byte[],OfficeAv1StillSequence,OfficeAv1StillFrame) Read(JsonElement c) {
        byte[] bytes;OfficeAvifImageItem item;var options=new OfficeRasterDecodeOptions();
        if(c.TryGetProperty("inputBase64",out var raw)) {
            bytes=Convert.FromBase64String(raw.GetString()!);var native=c.GetProperty("cdef");
            item=new OfficeAvifImageItem(1,c.GetProperty("upscaled").GetProperty("width").GetInt32(),native.GetProperty("height").GetInt32(),0,bytes.Length,
                new byte[] {0x81,(byte)native.GetProperty("level").GetInt32(),(byte)(c.GetProperty("monochrome").GetBoolean()?92:76),0},false,null);
        } else {
            bytes=File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif",c.GetProperty("name").GetString()+".avif"));
            Assert.True(OfficeAvifContainerReader.TryRead(bytes,options,out var container));item=c.GetProperty("alpha").GetBoolean()?container!.Alpha!:container!.Color;
        }
        Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes,item,options,out var sequence));
        Assert.True(OfficeAv1StillFrameReader.TryRead(bytes,item,sequence!,options,out var frame));return(bytes,sequence!,frame!);
    }
}
