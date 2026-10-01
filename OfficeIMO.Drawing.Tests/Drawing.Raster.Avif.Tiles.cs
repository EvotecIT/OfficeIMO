using System.Text.Json;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Complete independently encoded tiles protect coordinated grammar and terminal publication.</summary>
public sealed class DrawingAv1TileTests {
    [Theory]
    [InlineData("avif-opaque",false)]
    [InlineData("avif-alpha",false)]
    [InlineData("avif-alpha",true)]
    [InlineData("multitile",false)]
    public void FrozenTilesConsumeTheirEntireEntropyPayload(string name,bool alpha) {
        var (bytes,sequence,frame)=Fixture(name,alpha);
        using var fixture=Open();
        var expected=fixture.RootElement.GetProperty("cases").EnumerateArray().Single(c=>c.GetProperty("name").GetString()==name && c.GetProperty("alpha").GetBoolean()==alpha);
        for(int t=0;t<frame.Tiles.Length;t++) {
            var tile=frame.Tiles[t];
            var reader=new OfficeAv1TileReader(frame,sequence,tile,new OfficeRasterDecodeOptions());
            var sink=new Trace(expected.GetProperty("tiles")[t],tile);reader.Read(bytes,sink);
            Assert.True(sink.Complete);Assert.True(sink.Leaves>0);Assert.True(sink.Residuals>0);
            Assert.Throws<InvalidOperationException>(()=>reader.Read(bytes,sink));
        }
    }
    [Fact]
    public void InvalidTerminationCannotPublishCompletionOrResumeTheTile() {
        var (original,sequence,frame)=Fixture("avif-opaque",false);var tile=frame.Tiles[0];
        byte[] bytes=new byte[tile.Offset+tile.Length+1];Array.Copy(original,bytes,original.Length);bytes[bytes.Length-1]=1;
        var extended=new OfficeAv1Tile(tile.Offset,tile.Length+1,tile.MiRowStart,tile.MiRowEnd,tile.MiColStart,tile.MiColEnd);
        var reader=new OfficeAv1TileReader(frame,sequence,extended,new OfficeRasterDecodeOptions());var sink=new FailureSink();
        Assert.Throws<FormatException>(()=>reader.Read(bytes,sink));Assert.False(sink.Complete);Assert.True(sink.Residuals>0);
        Assert.Throws<InvalidOperationException>(()=>reader.Read(original,sink));
        var truncated=new OfficeAv1Tile(tile.Offset,tile.Length-1,tile.MiRowStart,tile.MiRowEnd,tile.MiColStart,tile.MiColEnd);
        reader=new OfficeAv1TileReader(frame,sequence,truncated,new OfficeRasterDecodeOptions());sink=new FailureSink();
        Assert.Throws<FormatException>(()=>reader.Read(original,sink));Assert.False(sink.Complete);
    }

    [Theory]
    [InlineData("begin")]
    [InlineData("residual")]
    [InlineData("end")]
    [InlineData("complete")]
    public void ConsumerFailurePermanentlyRetiresTheCoordinator(string stage) {
        var (bytes,sequence,frame)=Fixture("avif-opaque",false);
        var reader=new OfficeAv1TileReader(frame,sequence,frame.Tiles[0],new OfficeRasterDecodeOptions());
        var sink=new FailureSink {ThrowStage=stage};
        Assert.Throws<ApplicationException>(()=>reader.Read(bytes,sink));Assert.False(sink.Complete);
        Assert.Throws<InvalidOperationException>(()=>reader.Read(bytes,new FailureSink()));
    }

    [Fact]
    public void CancellationAggregateMemoryAndEncodedLimitsAreEnforcedAcrossAllOwners() {
        var (bytes,sequence,frame)=Fixture("avif-opaque",false);var tile=frame.Tiles[0];
        Assert.Throws<FormatException>(()=>new OfficeAv1TileReader(frame,sequence,tile,new OfficeRasterDecodeOptions {RetainedManagedBytes=OfficeRasterGuards.MaximumDecodedBytes-250000}));
        Assert.Throws<FormatException>(()=>new OfficeAv1TileReader(frame,sequence,tile,new OfficeRasterDecodeOptions {RetainedManagedBytes=long.MaxValue}));
        Assert.Throws<FormatException>(()=>new OfficeAv1TileReader(frame,sequence,tile,new OfficeRasterDecodeOptions {MaximumDecodedPixels=100}));
        var reader=new OfficeAv1TileReader(frame,sequence,tile,new OfficeRasterDecodeOptions {MaximumEncodedBytes=bytes.Length-1});
        var sink=new FailureSink();Assert.Throws<FormatException>(()=>reader.Read(bytes,sink));Assert.Equal(0,sink.Residuals);
        using var cancellation=new CancellationTokenSource();
        reader=new OfficeAv1TileReader(frame,sequence,tile,new OfficeRasterDecodeOptions {CancellationToken=cancellation.Token});
        sink=new FailureSink {Cancel=cancellation};
        Assert.Throws<OperationCanceledException>(()=>reader.Read(bytes,sink));Assert.Equal(1,sink.Residuals);Assert.False(sink.Complete);
        Assert.Throws<InvalidOperationException>(()=>reader.Read(bytes,new FailureSink()));
    }
    [Fact]
    public void EncodedBackingArraySharesTheRetainedMemoryBudgetEvenOutsideTileBounds() {
        var (bytes,sequence,frame)=Fixture("avif-opaque",false);
        var options=new OfficeRasterDecodeOptions {RetainedManagedBytes=OfficeRasterGuards.MaximumDecodedBytes-1000000};
        var control=new FailureSink();
        new OfficeAv1TileReader(frame,sequence,frame.Tiles[0],options).Read(bytes,control);
        Assert.True(control.Complete);
        byte[] padded=new byte[1024*1024];Array.Copy(bytes,padded,bytes.Length);
        var reader=new OfficeAv1TileReader(frame,sequence,frame.Tiles[0],options);var sink=new FailureSink();
        Assert.Throws<FormatException>(()=>reader.Read(padded,sink));
        Assert.Equal(0,sink.Residuals);Assert.False(sink.Complete);
        Assert.Throws<InvalidOperationException>(()=>reader.Read(bytes,new FailureSink()));
    }
    private sealed class FailureSink:IOfficeAv1TileConsumer {
        internal string? ThrowStage;internal CancellationTokenSource? Cancel;internal int Residuals;internal bool Complete;
        private void Check(string stage) {if(ThrowStage==stage) throw new ApplicationException("Failed reconstruction sink.");}
        public void Restoration(OfficeAv1RestorationUnit unit)=>Check("restoration");
        public void BeginBlock(OfficeAv1TileBlock block)=>Check("begin");
        public void Residual(OfficeAv1Coefficients coefficients) {Residuals++;Check("residual");Cancel?.Cancel();}
        public void EndBlock()=>Check("end");
        public void CompleteTile() {Check("complete");Complete=true;}
    }
    private sealed class Trace:IOfficeAv1TileConsumer {
        private readonly JsonElement[] _events;
        private readonly bool[] _coverage;
        private readonly OfficeAv1Tile _tile;
        private int _index,_pendingResiduals;
        private bool _skip,_pending;
        internal int Leaves,Residuals;internal bool Complete;
        internal Trace(JsonElement expected,OfficeAv1Tile tile) {
            _events=expected.GetProperty("events").EnumerateArray().ToArray();_tile=tile;
            Assert.Equal(new[] {tile.MiRowStart,tile.MiRowEnd,tile.MiColStart,tile.MiColEnd},Ints(expected.GetProperty("extent")));
            _coverage=new bool[(tile.MiRowEnd-tile.MiRowStart)*(tile.MiColEnd-tile.MiColStart)];
        }
        private JsonElement Next(string kind) { Assert.True(_index<_events.Length);var e=_events[_index++];Assert.Equal(kind,e.GetProperty("event").GetString());return e; }
        public void Restoration(OfficeAv1RestorationUnit unit) {
            var e=Next("restoration");Assert.Equal(e.GetProperty("plane").GetInt32(),unit.Plane);
            Assert.Equal(e.GetProperty("row").GetInt32(),unit.Row);Assert.Equal(e.GetProperty("col").GetInt32(),unit.Col);
            Assert.Equal(e.GetProperty("type").GetInt32(),unit.Type);Assert.Equal(e.GetProperty("set").GetInt32(),unit.SgrSet);
            Assert.Equal(Ints(e.GetProperty("wiener")),Enumerable.Range(0,6).Select(i=>unit.WienerTap(i/3,i%3)));
            Assert.Equal(Ints(e.GetProperty("xqd")),new[] {unit.X0,unit.X1});
        }
        public void BeginBlock(OfficeAv1TileBlock block) {
            Assert.False(_pending);_pending=true;Leaves++;
            var e=Next("leaf");var b=block.Region;
            Assert.Equal(Ints(e.GetProperty("block")),new[] {b.MiRow,b.MiCol,b.Width,b.Height});
            _skip=block.Prelude.Skip;Assert.Equal(e.GetProperty("skip").GetInt32()!=0,_skip);
            Assert.Equal(e.GetProperty("segment").GetInt32(),block.Prelude.SegmentId);
            Assert.Equal(e.GetProperty("q").GetInt32(),block.Prelude.CurrentQIndex);
            Assert.Equal(e.GetProperty("y").GetInt32(),(int)block.Modes.YMode);
            if(block.Modes.HasChroma) Assert.Equal(e.GetProperty("uv").GetInt32(),(int)block.Modes.UvMode);
            Assert.Equal(e.GetProperty("copy").GetInt32()!=0,block.Modes.UseIntraBlockCopy);
            Assert.Equal(Ints(e.GetProperty("motion")),new[] {block.Motion.Row,block.Motion.Col});
            Assert.Equal(e.GetProperty("palette")[0].GetInt32(),block.Palette.SizeY);
            if(block.Modes.HasChroma) Assert.Equal(e.GetProperty("palette")[1].GetInt32(),block.Palette.SizeUv);
            Assert.Equal(e.GetProperty("filter").GetInt32(),block.Palette.FilterMode);
            Assert.Equal(e.GetProperty("angles")[0].GetInt32(),block.Modes.AngleDeltaY);
            if(block.Modes.HasChroma) Assert.Equal(e.GetProperty("angles")[1].GetInt32(),block.Modes.AngleDeltaUv);
            _pendingResiduals=block.Transforms.Count;
            int columns=_tile.MiColEnd-_tile.MiColStart;
            for(int r=b.MiRow;r<Math.Min(_tile.MiRowEnd,b.MiRow+b.Height/4);r++)
                for(int c=b.MiCol;c<Math.Min(_tile.MiColEnd,b.MiCol+b.Width/4);c++) {
                    int i=(r-_tile.MiRowStart)*columns+c-_tile.MiColStart;Assert.False(_coverage[i]);_coverage[i]=true;
                }
        }
        public void Residual(OfficeAv1Coefficients coefficients) {
            Assert.True(_pending);Assert.True(_pendingResiduals-->0);Residuals++;
            if(_skip) { Assert.Equal(0,coefficients.EndOfBlock);Assert.All(Enumerable.Range(0,coefficients.Count),i=>Assert.Equal(0,coefficients.Value(i)));return; }
            var e=Next("residual");var b=coefficients.Block;
            Assert.Equal(Ints(e.GetProperty("block")),new[] {b.Plane,b.X,b.Y,b.Size});
            Assert.Equal(e.GetProperty("type").GetInt32(),(int)coefficients.Type);
            Assert.Equal(e.GetProperty("eob").GetInt32(),coefficients.EndOfBlock);
            Assert.Equal(Ints(e.GetProperty("values")),Enumerable.Range(0,coefficients.Count).Select(coefficients.Value));
        }
        public void EndBlock() { Assert.True(_pending);Assert.Equal(0,_pendingResiduals);_pending=false; }
        public void CompleteTile() { Assert.False(_pending);Assert.Equal(_events.Length,_index);Assert.All(_coverage,Assert.True);Complete=true; }
    }
    private static int[] Ints(JsonElement value)=>value.EnumerateArray().Select(v=>v.GetInt32()).ToArray();
    private static JsonDocument Open() {
        using var stream=File.OpenRead(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif","tile-reference.json.gz"));
        using var gzip=new System.IO.Compression.GZipStream(stream,System.IO.Compression.CompressionMode.Decompress);
        return JsonDocument.Parse(gzip);
    }
    private static (byte[],OfficeAv1StillSequence,OfficeAv1StillFrame) Fixture(string name,bool alpha) {
        byte[] bytes=File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory,"TestAssets","Avif",name+".avif"));
        var options=new OfficeRasterDecodeOptions();Assert.True(OfficeAvifContainerReader.TryRead(bytes,options,out var container));
        var item=alpha?container!.Alpha!:container!.Color;
        Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes,item,options,out var sequence));
        Assert.True(OfficeAv1StillFrameReader.TryRead(bytes,item,sequence!,options,out var frame));
        return (bytes,sequence!,frame!);
    }
}
