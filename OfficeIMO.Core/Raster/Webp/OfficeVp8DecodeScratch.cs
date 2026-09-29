namespace OfficeIMO.Drawing;

/// <summary>Operation-owned scratch reused for every VP8 macroblock on all target frameworks.</summary>
internal sealed class OfficeVp8DecodeScratch {
    internal readonly byte[] Predicted = new byte[256];
    internal readonly int[] BlockTop = new int[16];
    internal readonly int[] BlockLeft = new int[16];
    internal readonly int[] SubblockTop = new int[8];
    internal readonly int[] SubblockLeft = new int[8];
    internal readonly int[] Coefficients = new int[16];
    internal readonly int[] TransformTemp = new int[16];
    internal readonly int[] TransformOutput = new int[16];
    internal readonly int[] Y2Dc = new int[16];
}
