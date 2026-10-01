using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeAv1FrameReconstructor {
    // AV1 5.11.3/5.11.35: decoded 4x4 cells plus a one-cell superblock border.
    private void ResetDecoded() {
        for(int p=0;p<_pixels.Length;p++) {
            int sub=p==0?0:1,n=_sb/4>>sub,w=(_tile.MiColEnd-_sbCol)>>sub,h=(_tile.MiRowEnd-_sbRow)>>sub;
            Array.Clear(_decoded[p],0,_decoded[p].Length);
            for(int y=-1;y<=n;y++) for(int x=-1;x<=n;x++)
                _decoded[p][(y+1)*34+x+1]=(y<0 && x<w) || (x<0 && y<h);
            _decoded[p][(n+1)*34]=false;
        }
    }
    private void MarkDecoded(OfficeAv1TransformBlock b) {
        int sub=b.Plane==0?0:1,x=(b.X-(_sbCol*4>>sub))/4,y=(b.Y-(_sbRow*4>>sub))/4;
        for(int r=0;r<b.Height/4;r++) for(int c=0;c<b.Width/4;c++) _decoded[b.Plane][(y+r+1)*34+x+c+1]=true;
    }

    private OfficeAv1PredictionEdges Edges(OfficeAv1TransformBlock b) {
        int p=b.Plane,sub=p==0?0:1,stride=_stride>>sub;
        int originX=(_block.Region.MiCol>>sub)*4,originY=(_block.Region.MiRow>>sub)*4;
        bool top=b.Y>originY || (_block.Region.MiRow>>sub)>(_tile.MiRowStart>>sub);
        bool left=b.X>originX || (_block.Region.MiCol>>sub)>(_tile.MiColStart>>sub);
        int nTop=top?Math.Min(b.Width,(_frame.MiCols*4>>sub)-b.X):0;
        int nLeft=left?Math.Min(b.Height,(_frame.MiRows*4>>sub)-b.Y):0;
        int x=(b.X-(_sbCol*4>>sub))/4,y=(b.Y-(_sbRow*4>>sub))/4;
        bool right=b.X+b.Width<(_tile.MiColEnd*4>>sub) && _decoded[p][y*34+x+b.Width/4+1];
        bool below=b.Y+b.Height<(_tile.MiRowEnd*4>>sub) && _decoded[p][(y+b.Height/4+1)*34+x];
        int tr=top && nTop==b.Width && right?Math.Min(Math.Min(b.Width,b.Height),(_frame.MiCols*4>>sub)-b.X-b.Width):0;
        int bl=left && nLeft==b.Height && below?Math.Min(Math.Min(b.Height,b.Width),(_frame.MiRows*4>>sub)-b.Y-b.Height):0;
        for(int i=0;i<nTop+tr;i++) _above[i]=_pixels[p][(b.Y-1)*stride+b.X+i];
        for(int i=0;i<nLeft+bl;i++) _left[i]=_pixels[p][(b.Y+i)*stride+b.X-1];
        byte corner=top && left?_pixels[p][(b.Y-1)*stride+b.X-1]:(byte)128;
        return new OfficeAv1PredictionEdges(_above,_left,corner,nTop,nLeft,tr,bl);
    }
    private bool SmoothNeighbors(int plane) {
        var b=_block.Region;int r=b.MiRow,c=b.MiCol;
        int ar=r-1,ac=c,lr=r,lc=c-1;
        if(plane>0) {if((c&1)==0)ac++;if((r&1)!=0)ar--;if((c&1)!=0)lc--;if((r&1)==0)lr++;}
        var modes=plane==0?_yModes:_uvModes;
        bool above=(r>>(plane>0?1:0))>(_tile.MiRowStart>>(plane>0?1:0));
        bool left=(c>>(plane>0?1:0))>(_tile.MiColStart>>(plane>0?1:0));
        return above && Smooth(modes[ar*_frame.MiCols+ac]) || left && Smooth(modes[lr*_frame.MiCols+lc]);
    }
    private static bool Smooth(byte mode)=>mode>=9 && mode<=11;
}
