using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Tile-owned AV1 restoration unit syntax and reference taps, sections 5.11.57–58.</summary>
/// <remarks>All planes share the type CDFs; each plane has separate reference coefficients.
/// Units are streamed before partition syntax, without retaining a frame-sized filter grid.</remarks>
internal sealed class OfficeAv1RestorationReader {
    private readonly OfficeAv1TileGeometry _geometry;
    private readonly int _height, _width, _denominator, _planes;
    private readonly int[] _types=new int[3], _sizes=new int[3];
    private readonly int[] _wiener={11570,32768,0}, _sgr={16855,32768,0}, _switch={9413,22581,32768,0};
    private readonly int[,,] _referenceWiener=new int[3,2,3];
    private readonly int[,] _referenceSgr=new int[3,2];
    private readonly CancellationToken _cancellation;
    private int _nextRow, _nextCol;
    private bool _failed;

    internal OfficeAv1RestorationReader(OfficeAv1StillFrame frame, OfficeAv1StillSequence sequence,
        OfficeAv1Tile tile, OfficeRasterDecodeOptions options) {
        if(sequence==null) throw new ArgumentNullException(nameof(sequence));
        if(options==null) throw new ArgumentNullException(nameof(options));
        options.Validate(); options.CancellationToken.ThrowIfCancellationRequested();
        _geometry=new OfficeAv1TileGeometry(frame,tile,sequence.Use128Superblock?128:64);
        _geometry.EnsurePixelBudget(options.MaximumDecodedPixels);
        if(frame.Width<1 || frame.Width>65536 || frame.Height<1 || frame.Height>65536 || frame.UpscaledWidth<frame.Width || frame.UpscaledWidth>65536 ||
            frame.MiCols!=2*((frame.Width+7)/8) || frame.MiRows!=2*((frame.Height+7)/8) ||
            frame.SuperResolutionDenominator<8 || frame.SuperResolutionDenominator>16 ||
            (long)frame.UpscaledWidth*frame.Height>options.MaximumDecodedPixels)
            throw new FormatException("Invalid AV1 restoration frame geometry.");
        const int fixedBytes=2048;
        if(options.RetainedManagedBytes>OfficeRasterGuards.MaximumDecodedBytes-fixedBytes)
            throw new FormatException("AV1 restoration context exceeds the retained-memory limit.");
        _height=frame.Height; _width=frame.UpscaledWidth; _denominator=frame.SuperResolutionDenominator;
        _planes=sequence.Monochrome?1:3; _cancellation=options.CancellationToken;
        for(int p=0;p<3;p++) {
            int type=frame.RestorationTypes[p], size=frame.RestorationUnitSizes[p];
            if((uint)type>3 || (p>=_planes && type!=0) ||
                (type!=0 && (!sequence.Restoration || frame.AllLossless || frame.AllowIntraBlockCopy ||
                    (size!=32 && size!=64 && size!=128 && size!=256) || (p==0 && size<64))))
                throw new FormatException("Invalid AV1 restoration parameters.");
            _types[p]=type; _sizes[p]=size;
            for(int pass=0;pass<2;pass++) {
                _referenceSgr[p,pass]=pass==0?-32:31;
                _referenceWiener[p,pass,0]=3; _referenceWiener[p,pass,1]=-7; _referenceWiener[p,pass,2]=15;
            }
        }
        _nextRow=tile.MiRowStart; _nextCol=tile.MiColStart;
    }

    /// <summary>Reads each unit intersecting the next superblock in tile raster order.</summary>
    internal void ReadSuperblock(OfficeAv1SymbolReader symbols, int row, int col, IOfficeAv1TileConsumer consumer) {
        _cancellation.ThrowIfCancellationRequested();
        if(_failed) throw new FormatException("The AV1 restoration context has failed.");
        if(symbols==null) throw new ArgumentNullException(nameof(symbols));
        if(consumer==null) throw new ArgumentNullException(nameof(consumer));
        if(row!=_nextRow || col!=_nextCol || row>=_geometry.Tile.MiRowEnd)
            throw new FormatException("Invalid AV1 restoration superblock order.");
        try {
            int units=_geometry.SuperblockPixels/4;
            for(int p=0;p<_planes;p++) if(_types[p]!=0) {
                int sub=p==0?0:1, size=_sizes[p];
                int rows=Math.Max((((_height+(1<<sub)-1)>>sub)+size/2)/size,1);
                int cols=Math.Max((((_width+(1<<sub)-1)>>sub)+size/2)/size,1);
                int firstRow=(row*(4>>sub)+size-1)/size;
                int endRow=Math.Min(rows,((row+units)*(4>>sub)+size-1)/size);
                int numerator=(4>>sub)*_denominator, denominator=size*8;
                int firstCol=(col*numerator+denominator-1)/denominator;
                int endCol=Math.Min(cols,((col+units)*numerator+denominator-1)/denominator);
                for(int y=firstRow;y<endRow;y++) for(int x=firstCol;x<endCol;x++) {
                    _cancellation.ThrowIfCancellationRequested();
                    consumer.Restoration(ReadUnit(symbols,p,y,x));
                    _cancellation.ThrowIfCancellationRequested();
                }
            }
            _nextCol+=units;
            if(_nextCol>=_geometry.Tile.MiColEnd) { _nextCol=_geometry.Tile.MiColStart; _nextRow+=units; }
        } catch { _failed=true; throw; }
    }

    private OfficeAv1RestorationUnit ReadUnit(OfficeAv1SymbolReader symbols,int p,int row,int col) {
        int type=_types[p]==2?(symbols.ReadSymbol(_wiener)!=0?2:0)
            :_types[p]==3?(symbols.ReadSymbol(_sgr)!=0?3:0):symbols.ReadSymbol(_switch);
        if(_types[p]==1 && type!=0) type++; // Switchable entropy order is none, Wiener, self-guided.
        int set=0, x0=0, x1=0;
        if(type==2) {
            for(int pass=0;pass<2;pass++) for(int j=p==0?0:1;j<3;j++) {
                int low=j==0?-5:j==1?-23:-17, high=j==0?11:j==1?9:47;
                _referenceWiener[p,pass,j]=ReadSigned(symbols,low,high,j+1,_referenceWiener[p,pass,j]);
            }
        } else if(type==3) {
            set=symbols.ReadLiteral(4);
            x0=set>=10 && set<=13?0:ReadSigned(symbols,-96,32,4,_referenceSgr[p,0]);
            x1=set>=14?Math.Max(-32,Math.Min(95,128-x0)):ReadSigned(symbols,-32,96,4,_referenceSgr[p,1]);
            _referenceSgr[p,0]=x0; _referenceSgr[p,1]=x1;
        }
        return new OfficeAv1RestorationUnit(p,row,col,type,set,
            type==2 && p==0?_referenceWiener[p,0,0]:0,type==2?_referenceWiener[p,0,1]:0,type==2?_referenceWiener[p,0,2]:0,
            type==2 && p==0?_referenceWiener[p,1,0]:0,type==2?_referenceWiener[p,1,1]:0,type==2?_referenceWiener[p,1,2]:0,x0,x1);
    }

    private static int ReadSigned(OfficeAv1SymbolReader symbols,int low,int high,int k,int reference) {
        int count=high-low, r=reference-low, i=0, baseValue=0, value;
        while(true) {
            int bits=i==0?k:k+i-1, range=1<<bits;
            if(count<=baseValue+3*range) { value=baseValue+symbols.ReadNonSymmetric(count-baseValue); break; }
            if(!symbols.ReadBool()) { value=baseValue+symbols.ReadLiteral(bits); break; }
            i++; baseValue+=range;
        }
        return low+((r<<1)<=count?Recenter(r,value):count-1-Recenter(count-1-r,value));
    }
    private static int Recenter(int reference,int value) => value>2*reference?value
        :(value&1)!=0?reference-((value+1)>>1):reference+(value>>1);
}
