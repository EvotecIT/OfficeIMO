using System;
using System.Collections.Generic;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Tile-owned transform-size syntax and Main-8 residual traversal geometry.</summary>
/// <remarks>For intra-block copy the caller first consumes motion syntax. Coefficient/type syntax is a later owner.
/// Neighbor state is published by CompleteBlock only after the entire residual leaf succeeds.</remarks>
internal sealed partial class OfficeAv1TransformReader {
    private readonly OfficeAv1TileGeometry _geometry;
    private readonly int _mode;
    private readonly CancellationToken _cancellation;
    // Each border cell retains transform size, preceding block dimension and inter/skip flags.
    private readonly byte[] _above, _left;
    private readonly int[][] _depth = CreateDepth(), _split = CreateSplit();
    private bool _pending, _failed, _inter, _skip;
    private OfficeAv1BlockRegion _block;
    private OfficeAv1TransformLayout? _layout;

    internal OfficeAv1TransformReader(OfficeAv1StillFrame frame, OfficeAv1StillSequence sequence, OfficeAv1Tile tile,
        OfficeRasterDecodeOptions options) {
        if (sequence==null) throw new ArgumentNullException(nameof(sequence));
        if (options==null) throw new ArgumentNullException(nameof(options));
        options.Validate(); options.CancellationToken.ThrowIfCancellationRequested();
        _geometry=new OfficeAv1TileGeometry(frame,tile,sequence.Use128Superblock?128:64);
        _geometry.EnsurePixelBudget(options.MaximumDecodedPixels);
        if ((uint)frame.TransformMode>2) throw new FormatException("Invalid AV1 transform mode.");
        _mode=frame.TransformMode; _cancellation=options.CancellationToken;
        // Includes CDFs, a 32x32 leaf grid and both temporary/final arrays for at most 1536 residual blocks.
        const int fixedBytes=65536;
        const string limit="AV1 transform context exceeds the retained-memory limit.";
        if (options.RetainedManagedBytes>OfficeRasterGuards.MaximumDecodedBytes-fixedBytes) throw new FormatException(limit);
        long retained=options.RetainedManagedBytes+fixedBytes;
        _above=new byte[OfficeRasterGuards.EnsureByteArrayLength((tile.MiColEnd-tile.MiColStart)*3,ref retained,limit)];
        _left=new byte[OfficeRasterGuards.EnsureByteArrayLength((tile.MiRowEnd-tile.MiRowStart)*3,ref retained,limit)];
    }

    internal OfficeAv1TransformLayout Read(OfficeAv1SymbolReader symbols, OfficeAv1BlockRegion block,
        OfficeAv1BlockPrelude prelude, OfficeAv1IntraModes modes) {
        Check();
        if (symbols==null) throw new ArgumentNullException(nameof(symbols));
        _geometry.Validate(block);
        if (_pending) throw new InvalidOperationException("Complete the current AV1 transform leaf first.");
        try {
            int width=block.Width/4, height=block.Height/4, maximum=OfficeAv1TransformSize.Maximum(block.Width,block.Height);
            byte[] grid=new byte[width*height];
            int last=maximum;
            if (_mode==2 && (width>1 || height>1) && modes.UseIntraBlockCopy && !prelude.Skip && !prelude.Lossless) {
                int stepW=OfficeAv1TransformSize.Width(maximum)/4, stepH=OfficeAv1TransformSize.Height(maximum)/4;
                for (int row=0; row<height; row+=stepH) for (int col=0; col<width; col+=stepW)
                    ReadTree(symbols,block,grid,row,col,maximum,0,ref last);
            } else {
                if (prelude.Lossless) last=0;
                else if (maximum!=0 && _mode==2 && (!prelude.Skip || !modes.UseIntraBlockCopy)) {
                    int category=OfficeAv1TransformSize.Depth(maximum)-1;
                    int context=(DepthNeighbor(block,true)>=OfficeAv1TransformSize.Width(maximum)?1:0)+
                        (DepthNeighbor(block,false)>=OfficeAv1TransformSize.Height(maximum)?1:0);
                    int depth=symbols.ReadSymbol(_depth[category*3+context]);
                    while(depth-->0) last=OfficeAv1TransformSize.Split(last);
                }
                for (int i=0; i<grid.Length; i++) grid[i]=(byte)last;
            }
            var blocks=BuildResiduals(block,grid,last,prelude.Lossless,modes);
            _layout=new OfficeAv1TransformLayout(grid,width,last,blocks);
            _block=block; _inter=modes.UseIntraBlockCopy; _skip=prelude.Skip; _pending=true;
            return _layout;
        } catch { _failed=true; throw; }
    }

    private void ReadTree(OfficeAv1SymbolReader symbols, OfficeAv1BlockRegion b, byte[] grid,
        int row, int col, int size, int depth, ref int last) {
        Check();
        if (b.MiRow+row>=_geometry.MiRows || b.MiCol+col>=_geometry.MiCols) return;
        int w=OfficeAv1TransformSize.Width(size)/4, h=OfficeAv1TransformSize.Height(size)/4;
        bool split=false;
        if (size!=0 && depth<2) {
            int maxSquare=OfficeAv1TransformSize.SquareUp(OfficeAv1TransformSize.Maximum(b.Width,b.Height));
            int context=(OfficeAv1TransformSize.SquareUp(size)!=maxSquare?3:0)+(4-maxSquare)*6+
                (TreeNeighbor(b,grid,row,col,true)<w*4?1:0)+(TreeNeighbor(b,grid,row,col,false)<h*4?1:0);
            split=symbols.ReadSymbol(_split[context])!=0;
        }
        if (split) {
            int child=OfficeAv1TransformSize.Split(size), sw=OfficeAv1TransformSize.Width(child)/4, sh=OfficeAv1TransformSize.Height(child)/4;
            for (int y=0; y<h; y+=sh) for (int x=0; x<w; x+=sw) ReadTree(symbols,b,grid,row+y,col+x,child,depth+1,ref last);
        } else {
            for (int y=row; y<row+h; y++) for (int x=col; x<col+w; x++) grid[y*(b.Width/4)+x]=(byte)size;
            last=size;
        }
    }

    private int DepthNeighbor(OfficeAv1BlockRegion b, bool above) {
        var tile=_geometry.Tile;
        if (above?b.MiRow==tile.MiRowStart:b.MiCol==tile.MiColStart) return 0;
        byte[] axis=above?_above:_left;
        int offset=(above?b.MiCol-tile.MiColStart:b.MiRow-tile.MiRowStart)*3;
        return (axis[offset+2]&1)!=0?axis[offset+1]:above?OfficeAv1TransformSize.Width(axis[offset]):OfficeAv1TransformSize.Height(axis[offset]);
    }
    private int TreeNeighbor(OfficeAv1BlockRegion b, byte[] grid, int row, int col, bool above) {
        if (above?row!=0:col!=0) {
            int size=grid[(above?row-1:row)*(b.Width/4)+(above?col:col-1)];
            return above?OfficeAv1TransformSize.Width(size):OfficeAv1TransformSize.Height(size);
        }
        var tile=_geometry.Tile;
        if (above?b.MiRow==tile.MiRowStart:b.MiCol==tile.MiColStart) return 64;
        byte[] axis=above?_above:_left;
        int offset=(above?b.MiCol+col-tile.MiColStart:b.MiRow+row-tile.MiRowStart)*3;
        return axis[offset+2]==3?axis[offset+1]:above?OfficeAv1TransformSize.Width(axis[offset]):OfficeAv1TransformSize.Height(axis[offset]);
    }

    internal void CompleteBlock() {
        Check();
        if (!_pending) throw new InvalidOperationException("No AV1 transform leaf is pending completion.");
        var b=_block; var tile=_geometry.Tile; var layout=_layout!;
        int rows=Math.Min(b.Height/4,tile.MiRowEnd-b.MiRow), cols=Math.Min(b.Width/4,tile.MiColEnd-b.MiCol);
        byte flags=(byte)((_inter?1:0)|(_skip?2:0));
        for (int col=0; col<cols; col++) Publish(_above,(b.MiCol+col-tile.MiColStart)*3,layout.SizeAt(rows-1,col),b.Width,flags);
        for (int row=0; row<rows; row++) Publish(_left,(b.MiRow+row-tile.MiRowStart)*3,layout.SizeAt(row,cols-1),b.Height,flags);
        _pending=false; _layout=null;
    }
    private static void Publish(byte[] axis,int offset,int size,int dimension,byte flags) {
        axis[offset]=(byte)size; axis[offset+1]=(byte)dimension; axis[offset+2]=flags;
    }
    private void Check() {
        _cancellation.ThrowIfCancellationRequested();
        if (_failed) throw new FormatException("The AV1 transform context has failed.");
    }
}
