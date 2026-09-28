namespace OfficeIMO.Drawing;

/// <summary>Operation-owned scratch reused for every predicted block on all target frameworks.</summary>
internal sealed class OfficeVp8PredictionScratch {
    internal readonly byte[] Predicted = new byte[256];
    internal readonly int[] BlockTop = new int[16];
    internal readonly int[] BlockLeft = new int[16];
    internal readonly int[] SubblockTop = new int[8];
    internal readonly int[] SubblockLeft = new int[8];
}
