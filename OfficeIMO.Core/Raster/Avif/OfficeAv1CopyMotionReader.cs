using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Tile-owned reduced-still intra-block-copy prediction, displacement syntax and source validation.</summary>
/// <remarks>AV1 5.11.26/31/32, 6.10.25 and the intra-only specialization of 7.10.2.
/// The caller visits leaves in partition decode order and completes only after residuals/reconstruction succeed.
/// Completed ordinary intra leaves also publish their size for spatial candidate stepping.</remarks>
internal sealed partial class OfficeAv1CopyMotionReader {
    private readonly OfficeAv1TileGeometry _geometry;
    private readonly bool _allowed, _monochrome;
    private readonly CancellationToken _cancellation;
    // Per-Mi cell: signed row/col (16 bits each), width/height in Mi units (6 bits each), decoded/copy flags.
    private readonly ulong[] _cells;
    private readonly OfficeAv1CopyMotion[] _stack=new OfficeAv1CopyMotion[8];
    private readonly int[] _weights=new int[8];
    private readonly int[] _joint={4096,11264,19328,32768,0};
    private readonly int[][] _sign={new[] {16384,32768,0},new[] {16384,32768,0}};
    private readonly int[][] _classes=CreateClasses(), _class0=CreateClass0(), _bits=CreateBits();
    private OfficeAv1BlockRegion _block;
    private OfficeAv1CopyMotion _motion;
    private bool _pending, _copy, _failed;
    private int _count;

    internal OfficeAv1CopyMotionReader(OfficeAv1StillFrame frame, OfficeAv1StillSequence sequence, OfficeAv1Tile tile,
        OfficeRasterDecodeOptions options) {
        if (sequence==null) throw new ArgumentNullException(nameof(sequence));
        if (options==null) throw new ArgumentNullException(nameof(options));
        options.Validate(); options.CancellationToken.ThrowIfCancellationRequested();
        _geometry=new OfficeAv1TileGeometry(frame,tile,sequence.Use128Superblock?128:64);
        _geometry.EnsurePixelBudget(options.MaximumDecodedPixels);
        _allowed=frame.AllowIntraBlockCopy; _monochrome=sequence.Monochrome; _cancellation=options.CancellationToken;
        const string limit="AV1 copy context exceeds the retained-memory limit.";
        if (options.RetainedManagedBytes>OfficeRasterGuards.MaximumDecodedBytes-4096) throw new FormatException(limit);
        long retained=options.RetainedManagedBytes+4096;
        long count=(long)(tile.MiRowEnd-tile.MiRowStart)*(tile.MiColEnd-tile.MiColStart);
        OfficeRasterGuards.EnsureByteArrayLength(count*8,ref retained,limit);
        _cells=new ulong[(int)count];
    }

    /// <summary>Reads motion only for copy leaves; an ordinary intra leaf consumes no symbols.</summary>
    internal OfficeAv1CopyMotion Read(OfficeAv1SymbolReader symbols, OfficeAv1BlockRegion block, OfficeAv1IntraModes modes) {
        Check();
        if (symbols==null) throw new ArgumentNullException(nameof(symbols));
        _geometry.Validate(block);
        if (_pending) throw new InvalidOperationException("Complete the current AV1 copy leaf before reading another.");
        bool chroma=!_monochrome && !(block.Height==4 && (block.MiRow&1)==0) && !(block.Width==4 && (block.MiCol&1)==0);
        if (modes.HasChroma!=chroma || (modes.UseIntraBlockCopy && !_allowed) || Cell(block.MiRow,block.MiCol)!=0)
            throw new FormatException("Invalid AV1 copy leaf contract.");
        try {
            _block=block; _copy=modes.UseIntraBlockCopy; _motion=default;
            if (_copy) {
                var prediction=Prediction();
                int joint=symbols.ReadSymbol(_joint);
                int row=prediction.Row+(joint==2||joint==3?Component(symbols,0):0);
                int col=prediction.Col+(joint==1||joint==3?Component(symbols,1):0);
                _motion=new OfficeAv1CopyMotion(row,col);
                if (!ValidSource(chroma)) throw new FormatException("Invalid AV1 intra-block-copy source displacement.");
            }
            _pending=true; return _motion;
        } catch { _failed=true; throw; }
    }

    /// <summary>Publishes one successful leaf; no partial or failed leaf becomes a predictor.</summary>
    internal void CompleteBlock() {
        Check();
        if (!_pending) throw new InvalidOperationException("No AV1 copy leaf is pending completion.");
        var t=_geometry.Tile; var b=_block;
        ulong cell=(ulong)(ushort)_motion.Row|((ulong)(ushort)_motion.Col<<16)|((ulong)(b.Width/4)<<32)|
            ((ulong)(b.Height/4)<<38)|(1UL<<44)|(_copy?1UL<<45:0);
        int bottom=Math.Min(t.MiRowEnd,b.MiRow+b.Height/4), right=Math.Min(t.MiColEnd,b.MiCol+b.Width/4);
        for (int row=b.MiRow;row<bottom;row++)
            for (int col=b.MiCol;col<right;col++) _cells[Index(row,col)]=cell;
        _pending=false;
    }

    private int Component(OfficeAv1SymbolReader symbols,int component) {
        int sign=symbols.ReadSymbol(_sign[component]), kind=symbols.ReadSymbol(_classes[component]);
        int integer=0, magnitude;
        if (kind==0) magnitude=(symbols.ReadSymbol(_class0[component])+1)*8;
        else {
            for (int bit=0;bit<kind;bit++) integer|=symbols.ReadSymbol(_bits[component*10+bit])<<bit;
            magnitude=(2<<(kind+2))+(integer+1)*8;
        }
        // Reduced still frames force integer motion and disable high precision: fractional fields are implied 3/1.
        return sign==0?magnitude:-magnitude;
    }
    private static int[][] CreateClasses() => new[] {
        new[] {28672,30976,31858,32320,32551,32656,32740,32757,32762,32767,32768,0},
        new[] {28672,30976,31858,32320,32551,32656,32740,32757,32762,32767,32768,0}
    };
    private static int[][] CreateClass0() => new[] {new[] {27648,32768,0},new[] {27648,32768,0}};
    private static int[][] CreateBits() {
        int[] probabilities={17408,17920,18944,20480,22528,24576,28672,29952,29952,30720};
        var result=new int[20][];
        for (int i=0;i<result.Length;i++) result[i]=new[] {probabilities[i%10],32768,0};
        return result;
    }
    private int Index(int row,int col) => (row-_geometry.Tile.MiRowStart)*(_geometry.Tile.MiColEnd-_geometry.Tile.MiColStart)+col-_geometry.Tile.MiColStart;
    private ulong Cell(int row,int col) => _cells[Index(row,col)];
    private void Check() {
        _cancellation.ThrowIfCancellationRequested();
        if (_failed) throw new FormatException("The AV1 copy motion context has failed.");
    }
}
