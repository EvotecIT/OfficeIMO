using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeAv1CoefficientReader {
    private static readonly int[][] IntraInverse={Array.Empty<int>(),new[] {9,0,10,11,3,1,2},new[] {9,0,3,1,2}};
    private static readonly int[][] InterInverse={Array.Empty<int>(),new[] {9,10,11,12,13,14,15,0,1,2,4,5,3,6,7,8},
        new[] {9,10,11,0,1,2,4,5,3,6,7,8},new[] {9,0}};
    private static readonly int[] ModeTypes={0,1,2,0,3,1,2,2,1,3,1,2,3,0},FilterDirections={0,1,2,6,0};
    private int Set(int size) {
        int square=Square(size),up=OfficeAv1TransformSize.SquareUp(size);
        if(up>3) return 0;
        if(_modes.UseIntraBlockCopy) return _reduced || up==3?3:square==2?2:1;
        return up==3?0:_reduced || square==2?2:1;
    }
    private int ReadType(OfficeAv1SymbolReader symbols,int size) {
        int set=Set(size),square=Square(size);
        if(set==0 || _segmentQ[_prelude.SegmentId]==0) return 0;
        if(_modes.UseIntraBlockCopy) return InterInverse[set][symbols.ReadSymbol(_inter[set*4+square])];
        int direction=_filter>=0?FilterDirections[_filter]:(int)_modes.YMode;
        return IntraInverse[set][symbols.ReadSymbol(_intra[(set*4+square)*13+direction])];
    }
    private int PlaneType(OfficeAv1TransformBlock b) {
        if(_prelude.Lossless || OfficeAv1TransformSize.SquareUp(b.Size)>3) return 0;
        int type;
        if(b.Plane==0) return _types[(b.Y/4-_block.MiRow)*(_block.Width/4)+b.X/4-_block.MiCol];
        if(_modes.UseIntraBlockCopy) {
            int x=Math.Max(_block.MiCol,b.X/2),y=Math.Max(_block.MiRow,b.Y/2);
            type=_types[(y-_block.MiRow)*(_block.Width/4)+x-_block.MiCol];
        } else type=ModeTypes[(int)_modes.UvMode];
        return InSet(Set(b.Size),type)?type:0;
    }
    private bool InSet(int set,int type) {
        if(set==0) return type==0;
        if(_modes.UseIntraBlockCopy) return set==1 || (set==2?type<12:type==0 || type==9);
        return type<4 || type==9 || (set==1 && (type==10 || type==11));
    }
    private void StoreType(OfficeAv1TransformBlock b,int type) {
        for(int y=b.Y/4-_block.MiRow;y<(b.Y+b.Height)/4-_block.MiRow;y++)
            for(int x=b.X/4-_block.MiCol;x<(b.X+b.Width)/4-_block.MiCol;x++) _types[y*(_block.Width/4)+x]=(byte)type;
    }
    internal static int TransformClass(int type) => type>=10?(type%2==0?2:1):0;
    /// <summary>Generates normative row-major scans; native scan fixtures independently protect all shapes/classes.</summary>
    internal static int[] CreateScan(int width,int height,int txClass) {
        if(width<4 || height<4 || width>32 || height>32 || (width&(width-1))!=0 || (height&(height-1))!=0 || (uint)txClass>2)
            throw new ArgumentOutOfRangeException();
        var scan=new int[width*height];int i=0;
        if(txClass==2) { for(int j=0;j<scan.Length;j++) scan[j]=j; }
        else if(txClass==1) { for(int x=0;x<width;x++) for(int y=0;y<height;y++) scan[i++]=y*width+x; }
        else for(int d=0;d<width+height-1;d++) {
            int first=Math.Max(0,d-width+1),last=Math.Min(height-1,d);
            if(width<height || (width==height && (d&1)!=0)) for(int y=first;y<=last;y++) scan[i++]=y*width+d-y;
            else for(int y=last;y>=first;y--) scan[i++]=y*width+d-y;
        }
        return scan;
    }
}
