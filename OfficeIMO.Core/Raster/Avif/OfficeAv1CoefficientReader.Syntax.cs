using System;

namespace OfficeIMO.Drawing;

internal sealed partial class OfficeAv1CoefficientReader {
    private OfficeAv1Coefficients ReadCoefficients(OfficeAv1SymbolReader symbols,OfficeAv1TransformBlock b) {
        int w=Math.Min(32,b.Width),h=Math.Min(32,b.Height),txContext=(Square(b.Size)+OfficeAv1TransformSize.SquareUp(b.Size)+1)>>1;
        int ptype=b.Plane==0?0:1;
        var values=new int[w*h];
        if (_prelude.Skip) return new OfficeAv1Coefficients(b,OfficeAv1TransformType.DctDct,0,values);
        bool zero=symbols.ReadSymbol(_skipCdf[txContext*13+SkipContext(b)])!=0;
        int type=0,eob=0,levelSum=0,dcCategory=0;
        if (!zero) {
            if (b.Plane==0) StoreType(b,ReadType(symbols,b.Size));
            type=PlaneType(b);
            int txClass=TransformClass(type),multi=Log2(w)+Log2(h)-4;
            int point=symbols.ReadSymbol(_eob[multi][ptype*2+(txClass==0?0:1)])+1;
            eob=point<2?point:(1<<(point-2))+1;
            if(point>=3) {
                if(symbols.ReadSymbol(_extra[(txContext*2+ptype)*9+point-3])!=0) eob+=1<<(point-3);
                for(int shift=point-4;shift>=0;shift--) if(symbols.ReadBool()) eob+=1<<shift;
            }
            if(eob>values.Length) throw new FormatException("Invalid AV1 coefficient end position.");
            int[] scan=CreateScan(w,h,txClass);
            for(int c=eob-1;c>=0;c--) {
                Check(); int pos=scan[c];
                int level=c==eob-1?symbols.ReadSymbol(_last[(txContext*2+ptype)*4+LastContext(c,values.Length)])+1:
                    symbols.ReadSymbol(_base[(txContext*2+ptype)*42+BaseContext(b.Size,values,w,h,pos,txClass)]);
                if(level>2) {
                    int context=BrContext(values,w,h,pos,txClass);
                    for(int i=0;i<4;i++) {
                        int increment=symbols.ReadSymbol(_br[(Math.Min(txContext,3)*2+ptype)*21+context]);
                        level+=increment; if(increment<3) break;
                    }
                }
                values[pos]=level;
            }
            for(int c=0;c<eob;c++) {
                Check(); int pos=scan[c],level=values[pos]; bool negative=false;
                if(level!=0) negative=c==0?symbols.ReadSymbol(_dc[ptype*3+DcContext(b)])!=0:symbols.ReadBool();
                if(level>14) level=ReadGolomb(symbols)+14;
                if(pos==0 && level>0) dcCategory=negative?1:2;
                level &= 0xfffff; levelSum=Math.Min(63,levelSum+level); values[pos]=negative?-level:level;
            }
        } else if(b.Plane==0) StoreType(b,0);
        WriteBorders(b,levelSum,dcCategory);
        return new OfficeAv1Coefficients(b,(OfficeAv1TransformType)type,eob,values);
    }
    // Section 6.11.39 requires a terminating length bit by length 20; no unbounded unary loop.
    private static int ReadGolomb(OfficeAv1SymbolReader symbols) {
        int length=1;
        while(!symbols.ReadBool()) { if(length==20) throw new FormatException("AV1 coefficient Golomb length exceeds 20."); length++; }
        return (1<<(length-1))|symbols.ReadLiteral(length-1);
    }
    private static int LastContext(int c,int area) => c==0?0:c<=area/8?1:c<=area/4?2:3;
    private static int Square(int size) => Log2(Math.Min(OfficeAv1TransformSize.Width(size),OfficeAv1TransformSize.Height(size))/4);
    private static int Log2(int value) { int n=0; while((value>>=1)!=0)n++; return n; }
}
