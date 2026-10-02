using System.Text.Json;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Native restoration readers and geometry protect evolving references and superblock unit ownership.</summary>
public sealed class DrawingAv1RestorationTests {
    [Fact]
    public void NativeRestorationMatrixMatchesEveryUnitAndSignedFilterCoefficient() {
        using var stream=File.OpenRead(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif","restoration-reference.json.gz"));
        using var gzip=new System.IO.Compression.GZipStream(stream,System.IO.Compression.CompressionMode.Decompress);
        using var fixture=JsonDocument.Parse(gzip);
        Assert.True(fixture.RootElement.GetProperty("nativeSelfCheck").GetBoolean());
        foreach(var c in fixture.RootElement.GetProperty("cases").EnumerateArray()) {
            var sequence=new OfficeAv1StillSequence {Restoration=true,Monochrome=c.GetProperty("mono").GetInt32()!=0,Use128Superblock=c.GetProperty("large").GetInt32()!=0};
            int[] extent=Ints(c.GetProperty("extent"));
            var frame=new OfficeAv1StillFrame {Width=c.GetProperty("width").GetInt32(),UpscaledWidth=c.GetProperty("upwidth").GetInt32(),Height=c.GetProperty("height").GetInt32(),SuperResolutionDenominator=c.GetProperty("denom").GetInt32(),MiRows=extent[1],MiCols=extent[3]};
            for(int p=0;p<3;p++) {frame.RestorationTypes[p]=c.GetProperty("types")[p].GetInt32();frame.RestorationUnitSizes[p]=c.GetProperty("sizes")[p].GetInt32();}
            byte[] bytes=Hex(c.GetProperty("hex").GetString()!);
            var tile=new OfficeAv1Tile(0,bytes.Length,extent[0],extent[1],extent[2],extent[3]);
            var readers=Enumerable.Range(0,2).Select(_=>new OfficeAv1RestorationReader(frame,sequence,tile,new OfficeRasterDecodeOptions())).ToArray();
            var symbols=Enumerable.Range(0,2).Select(_=>new OfficeAv1SymbolReader(bytes,0,bytes.Length,100000,c.GetProperty("updates").GetInt32()==0)).ToArray();
            var sinks=Enumerable.Range(0,2).Select(_=>new Units(c.GetProperty("events"))).ToArray();
            // Each reader snapshots frame parameters; interleaving two readers proves references/CDFs are tile-owned.
            frame.RestorationTypes[0]=0;frame.RestorationUnitSizes[0]=32;frame.Height=1;frame.SuperResolutionDenominator=16;
            int units=sequence.Use128Superblock?32:16;
            for(int row=extent[0];row<extent[1];row+=units) for(int col=extent[2];col<extent[3];col+=units)
                for(int t=0;t<2;t++) readers[t].ReadSuperblock(symbols[t],row,col,sinks[t]);
            for(int t=0;t<2;t++) {symbols[t].Finish();Assert.Equal(c.GetProperty("events").GetArrayLength(),sinks[t].Count);}
        }
    }

    [Fact]
    public void InvalidGeometryOrderFailureCancellationAndMemoryBoundsCannotPublishMoreUnits() {
        var sequence=new OfficeAv1StillSequence {Restoration=true,Monochrome=true};
        var frame=new OfficeAv1StillFrame {Width=64,UpscaledWidth=64,Height=64,MiRows=16,MiCols=16};
        frame.RestorationTypes[0]=2;frame.RestorationUnitSizes[0]=64;
        var tile=new OfficeAv1Tile(0,1,0,16,0,16);var options=new OfficeRasterDecodeOptions();
        var reader=new OfficeAv1RestorationReader(frame,sequence,tile,options);
        var symbols=new OfficeAv1SymbolReader(new byte[] {255},0,1,1,true);
        var sink=new Units(default);
        Assert.Throws<FormatException>(()=>reader.ReadSuperblock(symbols,0,16,sink));
        Assert.Throws<FormatException>(()=>reader.ReadSuperblock(symbols,0,0,sink)); // Requires more than the one-symbol budget.
        Assert.Throws<FormatException>(()=>reader.ReadSuperblock(new OfficeAv1SymbolReader(new byte[] {255},0,1,100,true),0,0,sink));
        Assert.Equal(0,sink.Count);
        Assert.Throws<FormatException>(()=>new OfficeAv1RestorationReader(frame,sequence,tile,new OfficeRasterDecodeOptions {RetainedManagedBytes=long.MaxValue}));
        frame.RestorationUnitSizes[0]=63;Assert.Throws<FormatException>(()=>new OfficeAv1RestorationReader(frame,sequence,tile,options));frame.RestorationUnitSizes[0]=64;
        frame.AllowIntraBlockCopy=true;Assert.Throws<FormatException>(()=>new OfficeAv1RestorationReader(frame,sequence,tile,options));frame.AllowIntraBlockCopy=false;
        using var cancellation=new CancellationTokenSource();
        reader=new OfficeAv1RestorationReader(frame,sequence,tile,new OfficeRasterDecodeOptions {CancellationToken=cancellation.Token});cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(()=>reader.ReadSuperblock(symbols,0,0,sink));Assert.Equal(0,sink.Count);
    }
    private sealed class Units:IOfficeAv1TileConsumer {
        private readonly JsonElement _expected;internal int Count;
        internal Units(JsonElement expected) {_expected=expected;}
        public void Restoration(OfficeAv1RestorationUnit unit) {
            Assert.NotEqual(JsonValueKind.Undefined,_expected.ValueKind);
            var e=_expected[Count++];
            Assert.Equal(e.GetProperty("plane").GetInt32(),unit.Plane);Assert.Equal(e.GetProperty("row").GetInt32(),unit.Row);Assert.Equal(e.GetProperty("col").GetInt32(),unit.Col);
            Assert.Equal(e.GetProperty("type").GetInt32(),unit.Type);Assert.Equal(e.GetProperty("set").GetInt32(),unit.SgrSet);
            Assert.Equal(Ints(e.GetProperty("wiener")),Enumerable.Range(0,6).Select(i=>unit.WienerTap(i/3,i%3)));
            Assert.Equal(Ints(e.GetProperty("xqd")),new[] {unit.X0,unit.X1});
        }
        public void BeginBlock(OfficeAv1TileBlock block)=>throw new InvalidOperationException();
        public void Residual(OfficeAv1Coefficients coefficients)=>throw new InvalidOperationException();
        public void EndBlock()=>throw new InvalidOperationException();
        public void CompleteTile()=>throw new InvalidOperationException();
    }
    private static byte[] Hex(string value)=>Enumerable.Range(0,value.Length/2).Select(i=>Convert.ToByte(value.Substring(i*2,2),16)).ToArray();
    private static int[] Ints(JsonElement value)=>value.EnumerateArray().Select(v=>v.GetInt32()).ToArray();
}
