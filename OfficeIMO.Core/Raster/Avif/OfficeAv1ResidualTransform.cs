using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Tile-owned quantization snapshot and reusable Main-8 residual transform scratch.</summary>
/// <remarks>Sequential use only. Returned residuals are immutable; callers budget retained results separately.
/// Prediction and final sample clipping belong to the reconstruction consumer.</remarks>
internal sealed class OfficeAv1ResidualTransform {
    private readonly int _baseQ;
    private readonly bool _deltaQ;
    private readonly int[] _dc,_ac,_altQ,_matrix;
    private readonly bool[] _lossless;
    private readonly int[] _line,_copy;
    private readonly long _retained;
    private readonly CancellationToken _cancellation;

    internal OfficeAv1ResidualTransform(OfficeAv1StillFrame frame,OfficeRasterDecodeOptions options) {
        if(frame==null) throw new ArgumentNullException(nameof(frame));
        if(options==null) throw new ArgumentNullException(nameof(options));
        options.Validate();options.CancellationToken.ThrowIfCancellationRequested();
        const long context=4096+OfficeAv1QuantizationTables.MatrixBytes; // scratch, snapshots, fixed facts and array/object overhead
        if(options.RetainedManagedBytes>OfficeRasterGuards.MaximumDecodedBytes-context)
            throw new FormatException("AV1 residual contexts exceed the retained-memory limit.");
        if(frame.Width<1 || frame.Width>65536 || frame.Height<1 || frame.Height>65536 ||
           (long)frame.Width*frame.Height>options.MaximumDecodedPixels || (uint)frame.BaseQIndex>255)
            throw new FormatException("Invalid AV1 residual frame geometry or quantizer.");
        _retained=options.RetainedManagedBytes+context;_cancellation=options.CancellationToken;
        _dc=new int[3];_ac=new int[3];_altQ=new int[8];_matrix=new int[3];_lossless=new bool[8];
        _line=new int[64];_copy=new int[64];
        _baseQ=frame.BaseQIndex;_deltaQ=frame.DeltaQPresent;
        _dc[0]=frame.DeltaQYDc;_dc[1]=frame.DeltaQUDc;_dc[2]=frame.DeltaQVDc;
        _ac[1]=frame.DeltaQUAc;_ac[2]=frame.DeltaQVAc;
        bool zeroDeltas=true;
        for(int p=0;p<3;p++) {
            if(_dc[p]< -64 || _dc[p]>63 || _ac[p]< -64 || _ac[p]>63 || (uint)frame.QMatrixLevels[p]>15)
                throw new FormatException("Invalid AV1 residual quantization parameters.");
            zeroDeltas &= _dc[p]==0 && _ac[p]==0;
            _matrix[p]=frame.UsingQMatrix?frame.QMatrixLevels[p]:15;
        }
        for(int s=0;s<8;s++) {
            _altQ[s]=frame.SegmentationEnabled && frame.SegmentFeatures[s,0]?frame.SegmentData[s,0]:0;
            if(_altQ[s]< -255 || _altQ[s]>255) throw new FormatException("Invalid AV1 segment quantizer.");
            _lossless[s]=ClipQ(_baseQ+_altQ[s])==0 && zeroDeltas;
        }
    }

    /// <summary>Returns dequantized inverse-transform residuals, with FLIPADST paint orientation applied.</summary>
    internal OfficeAv1ResidualBlock Reconstruct(OfficeAv1Coefficients coefficients,OfficeAv1BlockPrelude prelude) {
        if(coefficients==null) throw new ArgumentNullException(nameof(coefficients));
        _cancellation.ThrowIfCancellationRequested();
        var block=coefficients.Block;int plane=block.Plane,type=(int)coefficients.Type;
        if((uint)plane>=3 || (uint)prelude.SegmentId>=8 || (uint)prelude.CurrentQIndex>255)
            throw new FormatException("Invalid AV1 residual plane or leaf quantizer.");
        bool lossless=_lossless[prelude.SegmentId];
        if(lossless!=prelude.Lossless) throw new FormatException("AV1 leaf lossless state disagrees with its segment.");
        OfficeAv1InverseTransform.Validate(block.Size,type,lossless);
        int w=block.Width,h=block.Height,tw=Math.Min(32,w),th=Math.Min(32,h);
        if(coefficients.Count!=tw*th || (uint)coefficients.EndOfBlock>(uint)coefficients.Count ||
           (prelude.Skip && coefficients.EndOfBlock!=0)) throw new FormatException("Invalid AV1 coefficient dimensions or end position.");
        long storage=(long)(coefficients.Count+w*h)*4+48;
        if(storage>OfficeRasterGuards.MaximumDecodedBytes-_retained)
            throw new FormatException("AV1 coefficient and residual storage exceed the retained-memory limit.");
        int qindex=ClipQ((_deltaQ?prelude.CurrentQIndex:_baseQ)+_altQ[prelude.SegmentId]);
        int dc=OfficeAv1QuantizationTables.Dc[ClipQ(qindex+_dc[plane])],ac=OfficeAv1QuantizationTables.Ac[ClipQ(qindex+_ac[plane])];
        int denom=w*h>1024?4:w*h>256?2:1;
        int level=lossless || type>=9?15:_matrix[plane];
        var values=new int[w*h];
        for(int row=0;row<th;row++) {
            _cancellation.ThrowIfCancellationRequested();
            for(int col=0;col<tw;col++) {
                int input=coefficients.Value(row*tw+col);
                if(input< -0xfffff || input>0xfffff) throw new FormatException("AV1 quantized coefficient exceeds its 20-bit bound.");
                int q=row==0 && col==0?dc:ac;
                if(level<15) q=(OfficeAv1QuantizationTables.Weight(level,plane>0,block.Size,row*tw+col)*q+16)>>5;
                long magnitude=(Math.Abs((long)input)*q & 0xffffff)/denom;
                long signed=input<0?-magnitude:magnitude;
                values[row*w+col]=(int)Math.Max(-32768,Math.Min(32767,signed));
            }
        }
        OfficeAv1InverseTransform.Apply(values,block.Size,type,lossless,_line,_copy,_cancellation);
        _cancellation.ThrowIfCancellationRequested();
        return new OfficeAv1ResidualBlock(block,values);
    }
    private static int ClipQ(int q)=>Math.Max(0,Math.Min(255,q));
}

/// <summary>Signed row-major residuals in final sample orientation, before adding prediction and clipping.</summary>
internal sealed class OfficeAv1ResidualBlock {
    private readonly int[] _values;
    internal OfficeAv1ResidualBlock(OfficeAv1TransformBlock block,int[] values) {Block=block;_values=values;}
    internal OfficeAv1TransformBlock Block {get;}
    internal int Width=>Block.Width;
    internal int Height=>Block.Height;
    internal int Count=>_values.Length;
    internal int Value(int index)=>_values[index];
}
