using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeAv1CopyMotionReader {
    private bool ValidSource(bool chroma) {
        var m=_motion; var b=_block; var t=_geometry.Tile;
        if (Math.Abs(m.Row)>=16384 || Math.Abs(m.Col)>=16384 || (m.Row&7)!=0 || (m.Col&7)!=0) return false;
        int top=b.MiRow*4+(m.Row>>3),left=b.MiCol*4+(m.Col>>3);
        int bottom=top+b.Height,right=left+b.Width;
        // Main-8 is 4:2:0. Tiny chroma-bearing leaves reference the preceding shared chroma footprint.
        if (chroma) {if (b.Width<8) left-=4;if (b.Height<8) top-=4;}
        if (top<t.MiRowStart*4 || left<t.MiColStart*4 || bottom>t.MiRowEnd*4 || right>t.MiColEnd*4) return false;
        int activeRow=b.MiRow*4/_geometry.SuperblockPixels, activeCol=b.MiCol>>4;
        int sourceRow=(bottom-1)/_geometry.SuperblockPixels,sourceCol=(right-1)>>6;
        int perRow=((t.MiColEnd-t.MiColStart-1)>>4)+1;
        if (sourceRow*perRow+sourceCol>=activeRow*perRow+activeCol-4) return false;
        int wavefront=(5+(_geometry.SuperblockPixels==128?1:0))*(activeRow-sourceRow);
        return sourceRow<=activeRow && sourceCol<activeCol-4+wavefront;
    }
}
