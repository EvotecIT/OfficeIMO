using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Owns Main-8 CDEF skip/index maps and an immutable snapshot of pre-CDEF samples.</summary>
internal sealed class OfficeAv1Cdef {
    private readonly int _rows,_cols,_regions,_bits,_damping;
    private readonly int[,] _strengths;
    private readonly byte[] _skip,_indices;
    private readonly int[] _partial=new int[120],_cost=new int[8];
    private readonly CancellationToken _cancellation;
    internal static long ContextBytes(OfficeAv1StillFrame frame,bool monochrome,int sb) {
        long pixels=(long)((frame.Width+sb-1)/sb*sb)*((frame.Height+sb-1)/sb*sb)*(monochrome?2:3)/2;
        return pixels*2+(long)frame.MiRows*frame.MiCols+((frame.MiRows+15)/16)*((frame.MiCols+15)/16)+2048;
    }
    internal OfficeAv1Cdef(OfficeAv1StillFrame frame,bool monochrome,int sb,OfficeRasterDecodeOptions options) {
        options.CancellationToken.ThrowIfCancellationRequested();
        if(frame.BitDepth!=8) throw new FormatException("AV1 high-bit-depth CDEF is not qualified.");
        if(options.RetainedManagedBytes>OfficeRasterGuards.MaximumDecodedBytes-ContextBytes(frame,monochrome,sb))
            throw new FormatException("AV1 CDEF contexts exceed retained memory.");
        if(frame.CdefBits<0 || frame.CdefBits>3 || frame.CdefDamping<3 || frame.CdefDamping>6)
            throw new FormatException("Invalid AV1 CDEF parameters.");
        _strengths=(int[,])frame.CdefStrengths.Clone();_bits=frame.CdefBits;_damping=frame.CdefDamping;
        for(int i=0;i<(1<<_bits);i++) for(int j=0;j<4;j++) {
            int s=_strengths[i,j];
            if((j&1)==0?(s<0 || s>15):(s!=0 && s!=1 && s!=2 && s!=4))
                throw new FormatException("Invalid AV1 CDEF strength.");
        }
        _rows=frame.MiRows;_cols=frame.MiCols;_regions=(_cols+15)/16;_cancellation=options.CancellationToken;
        _skip=new byte[checked(_rows*_cols)];_indices=new byte[((_rows+15)/16)*_regions];
        for(int i=0;i<_indices.Length;i++) _indices[i]=255;
    }
    internal void Leaf(OfficeAv1TileBlock block) {
        var b=block.Region;int index=block.Prelude.CdefIndex;
        if(index < -1 || index >= (1<<_bits)) throw new FormatException("Invalid AV1 CDEF index.");
        int endRow=Math.Min(_rows,b.MiRow+b.Height/4),endCol=Math.Min(_cols,b.MiCol+b.Width/4);
        for(int r=b.MiRow;r<endRow;r++) {
            _cancellation.ThrowIfCancellationRequested();
            for(int c=b.MiCol;c<endCol;c++) _skip[r*_cols+c]=block.Prelude.Skip?(byte)1:(byte)0;
        }
        if(index<0) return;
        for(int r=b.MiRow/16;r<=(endRow-1)/16;r++) for(int c=b.MiCol/16;c<=(endCol-1)/16;c++) {
            int slot=r*_regions+c;
            if(_indices[slot]!=255 && _indices[slot]!=index) throw new FormatException("Inconsistent AV1 CDEF region index.");
            _indices[slot]=(byte)index;
        }
    }
    /// <summary>All neighborhoods read a separately charged snapshot; modified pixels never feed later blocks.</summary>
    internal void Apply(ushort[][] planes,int stride) {
        var input=new ushort[planes.Length][];
        for(int p=0;p<planes.Length;p++) {
            _cancellation.ThrowIfCancellationRequested();int pitch=stride>>(p>0?1:0);
            input[p]=new ushort[planes[p].Length];
            for(int y=0;y<planes[p].Length/pitch;y++) {
                _cancellation.ThrowIfCancellationRequested();Array.Copy(planes[p],y*pitch,input[p],y*pitch,pitch);
            }
        }
        for(int r=0;r<_rows;r+=2) {
            _cancellation.ThrowIfCancellationRequested();
            for(int c=0;c<_cols;c+=2) {
                int index=_indices[r/16*_regions+c/16];
                if(index==255 || (_skip[r*_cols+c]!=0 && _skip[r*_cols+c+1]!=0 &&
                   _skip[(r+1)*_cols+c]!=0 && _skip[(r+1)*_cols+c+1]!=0)) continue;
                int direction=OfficeAv1CdefFilter.FindDirection(input[0],stride,r*4*stride+c*4,_partial,_cost,out int variance);
                for(int p=0;p<planes.Length;p++) {
                    int sub=p==0?0:1,primary=_strengths[index,p==0?0:2],secondary=_strengths[index,p==0?1:3];
                    int dir=primary==0?0:direction;
                    if(p==0) primary=OfficeAv1CdefFilter.AdjustStrength(primary,variance);
                    OfficeAv1CdefFilter.Apply(input[p],planes[p],stride>>sub,c*4>>sub,r*4>>sub,8>>sub,
                        _cols*4>>sub,_rows*4>>sub,primary,secondary,_damping-sub,dir);
                }
            }
        }
    }
}
