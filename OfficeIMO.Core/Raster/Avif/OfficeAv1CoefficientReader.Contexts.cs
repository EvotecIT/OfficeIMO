using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeAv1CoefficientReader {
    private static readonly byte[,,] BaseOffsets={{{0,1},{1,0},{1,1},{0,2},{2,0}},{{0,1},{1,0},{0,2},{0,3},{0,4}},{{0,1},{1,0},{2,0},{3,0},{4,0}}};
    private static readonly byte[,,] BrOffsets={{{0,1},{1,0},{1,1}},{{0,1},{1,0},{0,2}},{{0,1},{1,0},{2,0}}};
    private static int BaseContext(int size,int[] values,int w,int h,int pos,int txClass) {
        int row=pos/w,col=pos%w,mag=0;
        for(int i=0;i<5;i++) {
            int y=row+BaseOffsets[txClass,i,0],x=col+BaseOffsets[txClass,i,1];
            if(y<h && x<w) mag+=Math.Min(Math.Abs(values[y*w+x]),3);
        }
        int context=Math.Min((mag+1)>>1,4);
        if(txClass!=0) return context+26+Math.Min(txClass==2?row:col,2)*5;
        if(pos==0) return 0;
        int ow=OfficeAv1TransformSize.Width(size),oh=OfficeAv1TransformSize.Height(size);
        int offset=ow<oh && row<2?11:ow>oh && col<2?16:row+col<2?1:row+col<4?6:21;
        return context+offset;
    }
    private static int BrContext(int[] values,int w,int h,int pos,int txClass) {
        int row=pos/w,col=pos%w,mag=0;
        for(int i=0;i<3;i++) {
            int y=row+BrOffsets[txClass,i,0],x=col+BrOffsets[txClass,i,1];
            if(y<h && x<w) mag+=Math.Min(values[y*w+x],15);
        }
        mag=Math.Min((mag+1)>>1,6);
        return mag+(pos==0?0:(txClass==0?row<2 && col<2:txClass==1?col==0:row==0)?7:14);
    }
    private int Border(int plane,bool above,int absolute,int component) {
        int sub=plane==0?0:1;
        int max=(above?_geometry.MiCols:_geometry.MiRows)>>sub;
        if(absolute>=max) return 0;
        int origin=(above?_block.MiCol:_block.MiRow)>>sub,local=absolute-origin;
        byte[] pending=above?_localAbove[plane]:_localLeft[plane];
        if(local>=0 && local<32 && pending[local*3+2]!=0) return pending[local*3+component];
        if(above?_block.MiRow==_geometry.Tile.MiRowStart:_block.MiCol==_geometry.Tile.MiColStart) return 0;
        int tileStart=(above?_geometry.Tile.MiColStart:_geometry.Tile.MiRowStart)>>sub;
        byte[] completed=above?_above[plane]:_left[plane];
        return completed[(absolute-tileStart)*2+component];
    }
    private int SkipContext(OfficeAv1TransformBlock b) {
        int above=0,left=0;
        for(int i=0;i<b.Width/4;i++) {
            int level=Border(b.Plane,true,b.X/4+i,0);
            above=b.Plane==0?Math.Max(above,level):above|level|Border(b.Plane,true,b.X/4+i,1);
        }
        for(int i=0;i<b.Height/4;i++) {
            int level=Border(b.Plane,false,b.Y/4+i,0);
            left=b.Plane==0?Math.Max(left,level):left|level|Border(b.Plane,false,b.Y/4+i,1);
        }
        if(b.Plane==0) return _block.Width==b.Width && _block.Height==b.Height?0:above==0 && left==0?1:
            above==0 || left==0?2+(Math.Max(above,left)>3?1:0):Math.Max(above,left)<=3?4:Math.Min(above,left)<=3?5:6;
        return 7+(above!=0?1:0)+(left!=0?1:0)+(Math.Max(4,_block.Width/2)*Math.Max(4,_block.Height/2)>b.Width*b.Height?3:0);
    }
    private int DcContext(OfficeAv1TransformBlock b) {
        int sign=0;
        for(int i=0;i<b.Width/4;i++) sign+=Sign(Border(b.Plane,true,b.X/4+i,1));
        for(int i=0;i<b.Height/4;i++) sign+=Sign(Border(b.Plane,false,b.Y/4+i,1));
        return sign<0?1:sign>0?2:0;
    }
    private static int Sign(int category) => category==1?-1:category==2?1:0;
    private void WriteBorders(OfficeAv1TransformBlock b,int level,int dc) {
        int sub=b.Plane==0?0:1,x=b.X/4-(_block.MiCol>>sub),y=b.Y/4-(_block.MiRow>>sub);
        for(int i=x;i<x+b.Width/4;i++) Write(_localAbove[b.Plane],i,level,dc);
        for(int i=y;i<y+b.Height/4;i++) Write(_localLeft[b.Plane],i,level,dc);
    }
    private void ResetSkipped() {
        for(int p=0;p<(_modes.HasChroma?3:1);p++) {
            int sub=p==0?0:1;
            int cols=((_block.MiCol+_block.Width/4)>>sub)-(_block.MiCol>>sub);
            int rows=((_block.MiRow+_block.Height/4)>>sub)-(_block.MiRow>>sub);
            for(int i=0;i<cols;i++) Write(_localAbove[p],i,0,0);
            for(int i=0;i<rows;i++) Write(_localLeft[p],i,0,0);
        }
    }
    private static void Write(byte[] target,int index,int level,int dc) { target[index*3]=(byte)level; target[index*3+1]=(byte)dc; target[index*3+2]=1; }
}
