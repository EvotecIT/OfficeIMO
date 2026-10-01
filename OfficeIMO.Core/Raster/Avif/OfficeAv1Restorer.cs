using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Retains restoration units and immutable deblocked/CDEF sources through complete-frame filtering.</summary>
internal sealed class OfficeAv1Restorer {
    private readonly int _width,_height;
    private readonly int[] _sizes=new int[3],_rows=new int[3],_cols=new int[3],_types=new int[3];
    private readonly OfficeAv1RestorationUnit[][] _units;
    private readonly bool[][] _assigned;
    private readonly OfficeAv1RestorationFilter _filter;
    private readonly CancellationToken _cancellation;
    private readonly long _workLimit;
    private ushort[][]? _deblocked;
    internal static long ContextBytes(OfficeAv1StillFrame frame,bool monochrome,int sb) {
        long pixels=(long)((frame.UpscaledWidth+sb-1)/sb*sb)*((frame.Height+sb-1)/sb*sb)*(monochrome?2:3)/2,units=0;
        for(int p=0;p<(monochrome?1:3);p++) if(frame.RestorationTypes[p]!=0) {
            int sub=p==0?0:1,size=frame.RestorationUnitSizes[p];
            if(size!=32 && size!=64 && size!=128 && size!=256) throw new FormatException("Invalid AV1 restoration unit size.");
            units+=(long)Math.Max(((((frame.Height+(1<<sub)-1)>>sub)+size/2)/size),1)*
                Math.Max(((((frame.UpscaledWidth+(1<<sub)-1)>>sub)+size/2)/size),1);
        }
        return pixels*4+units*64+OfficeAv1RestorationFilter.ContextBytes;
    }
    internal OfficeAv1Restorer(OfficeAv1StillFrame frame,OfficeAv1StillSequence sequence,int sb,OfficeRasterDecodeOptions options) {
        options.CancellationToken.ThrowIfCancellationRequested();
        if((sequence.BitDepth!=8 && sequence.BitDepth!=10) || frame.BitDepth!=sequence.BitDepth)
            throw new FormatException("Invalid or inconsistent AV1 restoration bit depth.");
        if(options.RetainedManagedBytes>OfficeRasterGuards.MaximumDecodedBytes-ContextBytes(frame,sequence.Monochrome,sb))
            throw new FormatException("AV1 restoration exceeds retained memory.");
        _width=frame.UpscaledWidth;_height=frame.Height;_cancellation=options.CancellationToken;_workLimit=options.MaximumInspectionWorkPixels;
        _units=new OfficeAv1RestorationUnit[sequence.Monochrome?1:3][];_assigned=new bool[_units.Length][];
        for(int p=0;p<_units.Length;p++) {
            int sub=p==0?0:1,size=frame.RestorationUnitSizes[p],type=frame.RestorationTypes[p];_types[p]=type;_sizes[p]=size;
            if((uint)type>3 || (type!=0 && (!sequence.Restoration || frame.AllLossless || frame.AllowIntraBlockCopy || (p==0 && size<64))))
                throw new FormatException("Invalid AV1 restoration plane parameters.");
            if(type==0) {_units[p]=Array.Empty<OfficeAv1RestorationUnit>();_assigned[p]=Array.Empty<bool>();continue;}
            _rows[p]=Math.Max((((_height+(1<<sub)-1)>>sub)+size/2)/size,1);
            _cols[p]=Math.Max((((_width+(1<<sub)-1)>>sub)+size/2)/size,1);
            _units[p]=new OfficeAv1RestorationUnit[_rows[p]*_cols[p]];_assigned[p]=new bool[_units[p].Length];
        }
        _filter=new OfficeAv1RestorationFilter(sequence.BitDepth,_cancellation);
    }
    internal void Unit(OfficeAv1RestorationUnit unit) {
        _cancellation.ThrowIfCancellationRequested();int p=unit.Plane;
        if((uint)p>=(uint)_units.Length || (uint)unit.Row>=(uint)_rows[p] || (uint)unit.Col>=(uint)_cols[p] ||
           (unit.Type!=0 && unit.Type!=2 && unit.Type!=3) || (_types[p]!=1 && unit.Type!=0 && unit.Type!=_types[p]))
            throw new FormatException("Invalid AV1 restoration unit.");
        int index=unit.Row*_cols[p]+unit.Col;
        if(_assigned[p][index]) throw new FormatException("Duplicate AV1 restoration unit.");
        _units[p][index]=unit;_assigned[p][index]=true;
    }
    internal void CaptureDeblocked(ushort[][] pixels,int stride) { _deblocked=Copy(pixels,stride); }
    /// <summary>Applies the same normative upscaler to deblocked stripe samples before restoration.</summary>
    internal void UpscaleDeblocked(OfficeAv1Upscaler upscaler,int stride) {
        if(_deblocked==null) throw new InvalidOperationException("AV1 restoration requires deblocked stripe sources.");
        _deblocked=upscaler.Apply(_deblocked,stride);
    }
    internal void Apply(ushort[][] pixels,int stride) {
        if(_deblocked==null) throw new InvalidOperationException("AV1 restoration requires deblocked stripe sources.");
        for(int p=0;p<_assigned.Length;p++) for(int i=0;i<_assigned[p].Length;i++)
            if(!_assigned[p][i]) throw new FormatException("Incomplete AV1 restoration unit grid.");
        ushort[][] source=Copy(pixels,stride);long work=0;
        for(int p=0;p<pixels.Length;p++) if(_types[p]!=0) {
            int sub=p==0?0:1,width=(_width+(1<<sub)-1)>>sub,height=(_height+(1<<sub)-1)>>sub;
            for(int y=0;y<height;) {
                _cancellation.ThrowIfCancellationRequested();int stripe=(y+(8>>sub))/(64>>sub),start=(-8+stripe*64)>>sub,end=start+(64>>sub)-1;
                int h=Math.Min(height-y,end-y+1),row=Math.Min(_rows[p]-1,(y+(8>>sub))/_sizes[p]);
                for(int x=0;x<width;) {
                    int col=Math.Min(_cols[p]-1,x/_sizes[p]),stop=col==_cols[p]-1?width:Math.Min(width,(col+1)*_sizes[p]);
                    int w=Math.Min(64,stop-x);var unit=_units[p][row*_cols[p]+col];
                    if(unit.Type!=0) {
                        work+=(long)w*h;if(work>_workLimit) throw new FormatException("AV1 restoration exceeds filter work.");
                        _filter.Apply(_deblocked[p],source[p],pixels[p],stride>>sub,width,height,start,end,x,y,w,h,unit);
                    }
                    x+=w;
                }
                y+=h;
            }
        }
        _cancellation.ThrowIfCancellationRequested();
    }
    private ushort[][] Copy(ushort[][] pixels,int stride) {
        var copy=new ushort[pixels.Length][];
        for(int p=0;p<pixels.Length;p++) {
            _cancellation.ThrowIfCancellationRequested();int pitch=stride>>(p==0?0:1);copy[p]=new ushort[pixels[p].Length];
            for(int y=0;y<pixels[p].Length/pitch;y++) {_cancellation.ThrowIfCancellationRequested();Array.Copy(pixels[p],y*pitch,copy[p],y*pitch,pitch);}
        }
        return copy;
    }
}
