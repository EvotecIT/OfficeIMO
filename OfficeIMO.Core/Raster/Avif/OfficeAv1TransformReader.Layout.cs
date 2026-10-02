using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeAv1TransformReader {
    private OfficeAv1TransformBlock[] BuildResiduals(OfficeAv1BlockRegion b, byte[] grid, int last, bool lossless, OfficeAv1IntraModes modes) {
        var output=new List<OfficeAv1TransformBlock>(16);
        int widthChunks=Math.Max(1,b.Width/64), heightChunks=Math.Max(1,b.Height/64);
        int chunkW=widthChunks>1 || heightChunks>1?64:b.Width, chunkH=widthChunks>1 || heightChunks>1?64:b.Height;
        for (int cy=0; cy<heightChunks; cy++) for (int cx=0; cx<widthChunks; cx++) {
            for (int plane=0; plane<(modes.HasChroma?3:1); plane++) {
                Check();
                int sub=plane==0?0:1;
                int x=((b.MiCol+cx*16)>>sub)*4, y=((b.MiRow+cy*16)>>sub)*4;
                int width=Math.Max(4,chunkW>>sub), height=Math.Max(4,chunkH>>sub);
                if (modes.UseIntraBlockCopy && !lossless && plane==0) ResidualTree(output,b,grid,x,y,width,height);
                else {
                    int size=lossless?0:plane==0?last:ChromaSize(b);
                    for (int yy=0; yy<height; yy+=OfficeAv1TransformSize.Height(size))
                        for (int xx=0; xx<width; xx+=OfficeAv1TransformSize.Width(size)) Add(output,plane,x+xx,y+yy,size);
                }
            }
        }
        return output.ToArray();
    }
    private static int ChromaSize(OfficeAv1BlockRegion b) {
        int size=OfficeAv1TransformSize.Maximum(Math.Max(4,b.Width/2),Math.Max(4,b.Height/2));
        int w=OfficeAv1TransformSize.Width(size), h=OfficeAv1TransformSize.Height(size);
        if (w==64 || h==64) return OfficeAv1TransformSize.Find(w==16?16:32,h==16?16:32);
        return size;
    }
    private void ResidualTree(List<OfficeAv1TransformBlock> output, OfficeAv1BlockRegion b, byte[] grid,int x,int y,int w,int h) {
        Check();
        if (x>=_geometry.MiCols*4 || y>=_geometry.MiRows*4) return;
        int size=grid[(y/4-b.MiRow)*(b.Width/4)+(x/4-b.MiCol)];
        if (w<=OfficeAv1TransformSize.Width(size) && h<=OfficeAv1TransformSize.Height(size))
            Add(output,0,x,y,OfficeAv1TransformSize.Find(w,h));
        else if (w>h) { ResidualTree(output,b,grid,x,y,w/2,h); ResidualTree(output,b,grid,x+w/2,y,w/2,h); }
        else if (w<h) { ResidualTree(output,b,grid,x,y,w,h/2); ResidualTree(output,b,grid,x,y+h/2,w,h/2); }
        else {
            ResidualTree(output,b,grid,x,y,w/2,h/2); ResidualTree(output,b,grid,x+w/2,y,w/2,h/2);
            ResidualTree(output,b,grid,x,y+h/2,w/2,h/2); ResidualTree(output,b,grid,x+w/2,y+h/2,w/2,h/2);
        }
    }
    private void Add(List<OfficeAv1TransformBlock> output,int plane,int x,int y,int size) {
        Check(); int sub=plane==0?0:1;
        if (x>=(_geometry.MiCols*4>>sub) || y>=(_geometry.MiRows*4>>sub)) return;
        if (output.Count>=1536) throw new FormatException("AV1 residual transform limit exceeded.");
        output.Add(new OfficeAv1TransformBlock(plane,x,y,size));
    }
}
