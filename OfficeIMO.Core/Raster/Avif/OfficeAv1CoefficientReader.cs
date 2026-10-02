using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Tile-owned Main-8 coefficient grammar, transform types and evolving border contexts.</summary>
/// <remarks>BeginBlock receives already parsed geometry and modes; intra-copy motion precedes it.
/// CompleteBlock publishes one successful leaf. Pixel reconstruction and tile traversal remain separate.</remarks>
internal sealed partial class OfficeAv1CoefficientReader {
    private readonly OfficeAv1TileGeometry _geometry;
    private readonly CancellationToken _cancellation;
    private readonly int _baseQ;
    private readonly bool _reduced;
    private readonly int[] _segmentQ=new int[8];
    private readonly int[][] _skipCdf, _dc, _extra, _base, _br, _last, _intra, _inter;
    private readonly int[][][] _eob;
    // Level and DC category, plus a per-leaf written flag. No full-tile clone per leaf.
    private readonly byte[][] _above=new byte[3][], _left=new byte[3][];
    private readonly byte[][] _localAbove=new byte[3][], _localLeft=new byte[3][];
    private readonly byte[] _types=new byte[1024];
    private OfficeAv1BlockRegion _block;
    private OfficeAv1BlockPrelude _prelude;
    private OfficeAv1IntraModes _modes;
    private OfficeAv1TransformLayout? _layout;
    private int _filter, _index;
    private bool _pending, _failed;

    internal OfficeAv1CoefficientReader(OfficeAv1StillFrame frame, OfficeAv1StillSequence sequence,
        OfficeAv1Tile tile, OfficeRasterDecodeOptions options) {
        if (sequence==null) throw new ArgumentNullException(nameof(sequence));
        if (options==null) throw new ArgumentNullException(nameof(options));
        options.Validate(); options.CancellationToken.ThrowIfCancellationRequested();
        _geometry=new OfficeAv1TileGeometry(frame,tile,sequence.Use128Superblock?128:64);
        _geometry.EnsurePixelBudget(options.MaximumDecodedPixels);
        if ((uint)frame.BaseQIndex>255) throw new FormatException("Invalid AV1 coefficient quantizer.");
        _baseQ=frame.BaseQIndex; _reduced=frame.ReducedTransformSet; _cancellation=options.CancellationToken;
        for (int i=0;i<8;i++) {
            bool active=frame.SegmentationEnabled && frame.SegmentFeatures[i,0];
            if(active && (frame.SegmentData[i,0]<-255 || frame.SegmentData[i,0]>255))
                throw new FormatException("Invalid AV1 segment quantizer adjustment.");
            _segmentQ[i]=active?Math.Max(0,Math.Min(255,_baseQ+frame.SegmentData[i,0])):_baseQ;
        }
        // CDF payload/array headers, leaf overlays/type grid, scan and one at-most-1024 coefficient result.
        const int fixedBytes=196608;
        const string limit="AV1 coefficient context exceeds the retained-memory limit.";
        if (options.RetainedManagedBytes>OfficeRasterGuards.MaximumDecodedBytes-fixedBytes) throw new FormatException(limit);
        long retained=options.RetainedManagedBytes+fixedBytes;
        for (int p=0;p<3;p++) {
            int sub=p==0?0:1;
            _above[p]=new byte[OfficeRasterGuards.EnsureByteArrayLength(((tile.MiColEnd-tile.MiColStart)>>sub)*2,ref retained,limit)];
            _left[p]=new byte[OfficeRasterGuards.EnsureByteArrayLength(((tile.MiRowEnd-tile.MiRowStart)>>sub)*2,ref retained,limit)];
            _localAbove[p]=new byte[32*3]; _localLeft[p]=new byte[32*3];
        }
        int q=_baseQ<=20?0:_baseQ<=60?1:_baseQ<=120?2:3;
        _skipCdf=Slice(CreateSkip(),q,65); _dc=Slice(CreateDc(),q,6); _extra=Slice(CreateExtra(),q,90);
        _last=Slice(CreateLast(),q,40);
        _base=q==0?CreateBase0():q==1?CreateBase1():q==2?CreateBase2():CreateBase3();
        _br=q==0?CreateBr0():q==1?CreateBr1():q==2?CreateBr2():CreateBr3();
        _eob=new[] {Slice(CreateEob0(),q,4),Slice(CreateEob1(),q,4),Slice(CreateEob2(),q,4),
            Slice(CreateEob3(),q,4),Slice(CreateEob4(),q,4),Slice(CreateEob5(),q,4),Slice(CreateEob6(),q,4)};
        _intra=CreateIntra(); _inter=CreateInter();
    }

    internal void BeginBlock(OfficeAv1BlockRegion block, OfficeAv1BlockPrelude prelude, OfficeAv1IntraModes modes,
        OfficeAv1Palette palette, OfficeAv1TransformLayout layout) {
        Check(); _geometry.Validate(block);
        if (palette==null) throw new ArgumentNullException(nameof(palette));
        if (layout==null) throw new ArgumentNullException(nameof(layout));
        if (_pending) throw new InvalidOperationException("Complete the current AV1 coefficient leaf first.");
        if ((uint)prelude.SegmentId>=8 || (uint)modes.YMode>12 || (uint)modes.UvMode>13 ||
            palette.FilterMode < -1 || palette.FilterMode>4 || layout.Count<1 || layout.Count>1536)
            throw new FormatException("Invalid AV1 coefficient leaf inputs.");
        for(int i=0;i<layout.Count;i++) {
            var b=layout.Block(i); int sub=b.Plane==0?0:1;
            if ((uint)b.Plane>2 || (b.Plane>0 && !modes.HasChroma) ||
                b.X<((block.MiCol>>sub)*4) || b.Y<((block.MiRow>>sub)*4) ||
                b.X+b.Width>((block.MiCol>>sub)*4+Math.Max(4,block.Width>>sub)) ||
                b.Y+b.Height>((block.MiRow>>sub)*4+Math.Max(4,block.Height>>sub)) || (b.X&3)!=0 || (b.Y&3)!=0)
                throw new FormatException("Invalid AV1 residual coefficient geometry.");
        }
        _block=block; _prelude=prelude; _modes=modes; _filter=palette.FilterMode; _layout=layout; _index=0;
        Array.Clear(_types,0,_types.Length);
        for(int p=0;p<3;p++) { Array.Clear(_localAbove[p],0,96); Array.Clear(_localLeft[p],0,96); }
        _pending=true;
        if (prelude.Skip) ResetSkipped();
    }

    /// <summary>Reads the next residual in the supplied traversal order, without exposing mutable context or arrays.</summary>
    internal OfficeAv1Coefficients Read(OfficeAv1SymbolReader symbols) {
        Check();
        if (symbols==null) throw new ArgumentNullException(nameof(symbols));
        if (!_pending || _index>=_layout!.Count) throw new InvalidOperationException("No AV1 residual is pending.");
        try {
            var block=_layout.Block(_index);
            var result=ReadCoefficients(symbols,block);
            _index++; return result;
        } catch { _failed=true; throw; }
    }

    internal void CompleteBlock() {
        Check();
        if (!_pending || _index!=_layout!.Count) throw new InvalidOperationException("AV1 residual leaf is incomplete.");
        for(int p=0;p<(_modes.HasChroma?3:1);p++) {
            Publish(_above[p],_localAbove[p],(_block.MiCol- _geometry.Tile.MiColStart)>>(p==0?0:1));
            Publish(_left[p],_localLeft[p],(_block.MiRow- _geometry.Tile.MiRowStart)>>(p==0?0:1));
        }
        _pending=false; _layout=null;
    }
    private static void Publish(byte[] target,byte[] source,int origin) {
        for(int i=0;i<32 && origin+i<target.Length/2;i++) if(source[i*3+2]!=0) {
            target[(origin+i)*2]=source[i*3]; target[(origin+i)*2+1]=source[i*3+1];
        }
    }
    private static int[][] Slice(int[][] source,int q,int count) { var result=new int[count][]; Array.Copy(source,q*count,result,0,count); return result; }
    private void Check() { _cancellation.ThrowIfCancellationRequested(); if (_failed) throw new FormatException("The AV1 coefficient context has failed."); }
}
