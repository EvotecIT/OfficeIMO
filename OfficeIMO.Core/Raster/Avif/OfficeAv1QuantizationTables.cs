using System;

namespace OfficeIMO.Drawing;

/// <summary>Normative AV1 8-bit and 10-bit quantization facts, independent of reference-code runtime packages.</summary>
internal static partial class OfficeAv1QuantizationTables {
    internal const int MatrixBytes=15*2*3344;
    private static readonly int[] Offsets={0,16,80,336,336,1360,1392,1424,1552,1680,2192,336,336,2704,2768,2832,3088,1680,2192};
    private static readonly byte[] Matrices=ReadMatrices();
    internal static int Weight(int level,bool chroma,int size,int position) => Matrices[(level*2+(chroma?1:0))*3344+Offsets[size]+position];
    private static byte[] ReadMatrices() {
        using var stream=typeof(OfficeAv1QuantizationTables).Assembly.GetManifestResourceStream("OfficeIMO.Core.Av1QuantizerMatrices");
        if(stream==null) throw new InvalidOperationException("The AV1 numeric matrix resource is missing.");
        var result=new byte[MatrixBytes];int offset=0;
        while(offset<result.Length) {int count=stream.Read(result,offset,result.Length-offset);if(count==0) throw new InvalidOperationException("The AV1 numeric matrix resource is truncated.");offset+=count;}
        if(stream.ReadByte()!=-1) throw new InvalidOperationException("The AV1 numeric matrix resource has an invalid length.");
        return result;
    }
}
