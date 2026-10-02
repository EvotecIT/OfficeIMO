using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Tile-local Main-8 reduced-still luma/chroma modes, directional angles and chroma-from-luma alphas.</summary>
/// <remarks>Reads AV1 5.11.7 through the palette boundary. CompleteBlock publishes luma neighbors only
/// after palette/filter-intra, transforms and residual syntax succeeds. The caller follows partition decode order.
/// One tile owns its CDFs; luma/chroma share angle CDFs and U/V share the CfL alpha CDFs.</remarks>
internal sealed partial class OfficeAv1IntraModeReader {
    private readonly OfficeAv1TileGeometry _geometry;
    private readonly bool _monochrome, _allowCopy;
    private readonly byte[] _aboveModes, _leftModes;
    private readonly CancellationToken _cancellation;
    private readonly int[][] _lumaCdfs=CreateLumaCdfs(), _uvWithoutCfl=CreateUvWithoutCflCdfs(), _uvWithCfl=CreateUvWithCflCdfs();
    private readonly int[][] _angleCdfs=CreateAngleCdfs(), _alphaCdfs=CreateAlphaCdfs();
    private readonly int[] _signCdf={1418,2123,13340,18405,26972,28343,32294,32768,0}, _copyCdf={30531,32768,0};
    private bool _pending, _failed;
    private OfficeAv1BlockRegion _pendingBlock;
    private OfficeAv1IntraModes _pendingModes;

    internal OfficeAv1IntraModeReader(OfficeAv1StillFrame frame, OfficeAv1StillSequence sequence, OfficeAv1Tile tile,
        OfficeRasterDecodeOptions options) {
        if (sequence==null) throw new ArgumentNullException(nameof(sequence));
        if (options==null) throw new ArgumentNullException(nameof(options));
        options.Validate(); options.CancellationToken.ThrowIfCancellationRequested();
        _geometry=new OfficeAv1TileGeometry(frame,tile,sequence.Use128Superblock?128:64);
        _geometry.EnsurePixelBudget(options.MaximumDecodedPixels);
        _monochrome=sequence.Monochrome; _allowCopy=frame.AllowIntraBlockCopy; _cancellation=options.CancellationToken;
        const string limit="AV1 intra context exceeds the retained-memory limit.";
        if (options.RetainedManagedBytes>OfficeRasterGuards.MaximumDecodedBytes-8192) throw new FormatException(limit);
        long retained=options.RetainedManagedBytes+8192; // fixed CDF/state payloads
        int cols=OfficeRasterGuards.EnsureByteArrayLength(tile.MiColEnd-tile.MiColStart,ref retained,limit);
        int rows=OfficeRasterGuards.EnsureByteArrayLength(tile.MiRowEnd-tile.MiRowStart,ref retained,limit);
        _aboveModes=new byte[cols]; _leftModes=new byte[rows];
    }

    /// <summary>Consumes prediction selection, not palette/filter-intra or intra-block-copy motion syntax.</summary>
    internal OfficeAv1IntraModes Read(OfficeAv1SymbolReader symbols, OfficeAv1BlockRegion block, OfficeAv1BlockPrelude prelude) {
        Check();
        if (symbols==null) throw new ArgumentNullException(nameof(symbols));
        _geometry.Validate(block);
        if (_pending) throw new InvalidOperationException("Complete the current AV1 leaf before reading another mode.");
        try {
            bool chroma=!_monochrome && !(block.Height==4 && (block.MiRow&1)==0) && !(block.Width==4 && (block.MiCol&1)==0);
            bool copy=_allowCopy && symbols.ReadSymbol(_copyCdf)!=0;
            var y=OfficeAv1IntraMode.Dc; var uv=OfficeAv1IntraMode.Dc;
            int angleY=0, angleUv=0, alphaU=0, alphaV=0;
            bool allowed=false;
            if (!copy) {
                var tile=_geometry.Tile;
                int above=block.MiRow>tile.MiRowStart ? _aboveModes[block.MiCol-tile.MiColStart] : 0;
                int left=block.MiCol>tile.MiColStart ? _leftModes[block.MiRow-tile.MiRowStart] : 0;
                y=(OfficeAv1IntraMode)symbols.ReadSymbol(_lumaCdfs[ModeContext(above)*5+ModeContext(left)]);
                bool angles=block.Width*block.Height>=64; // All sizes except 4x4, 4x8 and 8x4 (including 4x16/16x4).
                angleY=ReadAngle(symbols,y,angles);
                if (chroma) {
                    // Main-8 is 4:2:0: chroma residuals round each dimension up to at least four pixels.
                    allowed=prelude.Lossless ? block.Width<=8 && block.Height<=8 : block.Width<=32 && block.Height<=32;
                    uv=(OfficeAv1IntraMode)symbols.ReadSymbol((allowed?_uvWithCfl:_uvWithoutCfl)[(int)y]);
                    if (uv==OfficeAv1IntraMode.ChromaFromLuma) {
                        int signs=symbols.ReadSymbol(_signCdf), signU=(signs+1)/3, signV=(signs+1)%3;
                        if (signU!=0) alphaU=ReadAlpha(symbols,(signU-1)*3+signV,signU);
                        if (signV!=0) alphaV=ReadAlpha(symbols,(signV-1)*3+signU,signV);
                    }
                    angleUv=ReadAngle(symbols,uv,angles);
                }
            }
            _pendingBlock=block;
            _pendingModes=new OfficeAv1IntraModes(copy,chroma,allowed,y,uv,angleY,angleUv,alphaU,alphaV);
            _pending=true;
            return _pendingModes;
        } catch { _failed=true; throw; }
    }

    /// <summary>Retains the completed leaf's luma mode across its visible Mi cells; coded edge padding is not stored.</summary>
    internal void CompleteBlock() {
        Check();
        if (!_pending) throw new InvalidOperationException("No AV1 intra mode is pending completion.");
        var tile=_geometry.Tile; var b=_pendingBlock;
        byte mode=(byte)_pendingModes.YMode;
        int cols=Math.Min(tile.MiColEnd,b.MiCol+b.Width/4), rows=Math.Min(tile.MiRowEnd,b.MiRow+b.Height/4);
        for (int c=b.MiCol;c<cols;c++) _aboveModes[c-tile.MiColStart]=mode;
        for (int r=b.MiRow;r<rows;r++) _leftModes[r-tile.MiRowStart]=mode;
        _pending=false;
    }

    private int ReadAngle(OfficeAv1SymbolReader symbols, OfficeAv1IntraMode mode, bool enabled) =>
        enabled && (int)mode>=1 && (int)mode<=8
            ? symbols.ReadSymbol(_angleCdfs[(int)mode-1])-3 : 0;
    private int ReadAlpha(OfficeAv1SymbolReader symbols,int context,int sign) {
        int alpha=1+symbols.ReadSymbol(_alphaCdfs[context]); return sign==1?-alpha:alpha;
    }
    private static int ModeContext(int mode) {
        switch (mode) {
            case 1: case 10:return 1;
            case 2: case 11:return 2;
            case 3: case 8:return 3;
            case 4: case 5: case 6: case 7:return 4;
            default:return 0;
        }
    }
    private void Check() {
        _cancellation.ThrowIfCancellationRequested();
        if (_failed) throw new FormatException("The AV1 intra mode context has failed.");
    }
}
