using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeAv1CopyMotionReader {
    private OfficeAv1CopyMotion Prediction() {
        _count=0;
        Scan(true,-1); Scan(false,-1);
        if (Math.Max(_block.Width,_block.Height)<=64) Point(-1,_block.Width/4);
        int nearest=_count;
        for (int i=0;i<nearest;i++) _weights[i]+=640;
        Point(-1,-1); Scan(true,-3); Scan(false,-3);
        if (_block.Height>4) Scan(true,-5);
        if (_block.Width>4) Scan(false,-5);
        Sort(0,nearest); Sort(nearest,_count);
        // In a still intra frame no external-reference candidate can survive the extra search.
        for (int i=_count;i<2;i++) _stack[i]=default;
        for (int i=0;i<_count;i++) {
            var m=_stack[i]; var b=_block;
            _stack[i]=new OfficeAv1CopyMotion(
                Clip(m.Row,-b.MiRow*32-128-b.Height*8,(_geometry.MiRows-b.Height/4-b.MiRow)*32+128+b.Height*8),
                Clip(m.Col,-b.MiCol*32-128-b.Width*8,(_geometry.MiCols-b.Width/4-b.MiCol)*32+128+b.Width*8));
        }
        var prediction=_stack[0];
        if (prediction.Row==0 && prediction.Col==0) prediction=_stack[1];
        if (prediction.Row==0 && prediction.Col==0) {
            int size=_geometry.SuperblockPixels;
            prediction=_block.MiRow-size/4<_geometry.Tile.MiRowStart
                ? new OfficeAv1CopyMotion(0,-(size+256)*8) : new OfficeAv1CopyMotion(-size*8,0);
        }
        return prediction;
    }
    private void Scan(bool horizontal,int delta) {
        var b=_block;
        int extent=(horizontal?b.Width:b.Height)/4;
        int end=Math.Min(Math.Min(extent,(horizontal?_geometry.MiCols-b.MiCol:_geometry.MiRows-b.MiRow)),16);
        int dr=horizontal?delta:0, dc=horizontal?0:delta;
        if (Math.Abs(delta)>1) {
            if (horizontal) {dr+=b.MiRow&1;dc=1-(b.MiCol&1);}
            else {dr=1-(b.MiRow&1);dc+=b.MiCol&1;}
        }
        for (int i=0;i<end;) {
            int row=b.MiRow+dr+(horizontal?0:i), col=b.MiCol+dc+(horizontal?i:0);
            if (!Inside(row,col)) break;
            ulong cell=Cell(row,col);
            int dimension=(int)((cell>>(horizontal?32:38))&63);
            int length=Math.Min(extent,Math.Max(1,dimension));
            if (Math.Abs(horizontal?dr:dc)>1) length=Math.Max(2,length);
            if (extent>=16) length=Math.Max(4,length);
            Add(cell,length*2); i+=length;
        }
    }
    private void Point(int dr,int dc) {
        int row=_block.MiRow+dr,col=_block.MiCol+dc;
        if (Inside(row,col)) Add(Cell(row,col),4);
    }
    private void Add(ulong cell,int weight) {
        if ((cell&(1UL<<45))==0) return;
        var m=new OfficeAv1CopyMotion((short)(ushort)cell,(short)(ushort)(cell>>16));
        int index=0;
        while (index<_count && (_stack[index].Row!=m.Row || _stack[index].Col!=m.Col)) index++;
        if (index<_count) _weights[index]+=weight;
        else if (_count<8) {_stack[_count]=m;_weights[_count++]=weight;}
    }
    private void Sort(int start,int end) {
        for (int i=start+1;i<end;i++) {
            var m=_stack[i];int weight=_weights[i],j=i;
            while (j>start && _weights[j-1]<weight) {_stack[j]=_stack[j-1];_weights[j]=_weights[j-1];j--;}
            _stack[j]=m;_weights[j]=weight;
        }
    }
    private bool Inside(int row,int col) {
        var t=_geometry.Tile;
        return row>=t.MiRowStart && row<t.MiRowEnd && col>=t.MiColStart && col<t.MiColEnd;
    }
    private static int Clip(int value,int min,int max) => Math.Max(min,Math.Min(max,value));
}
