using System;

namespace OfficeIMO.Drawing;

internal static partial class OfficeJpegXrDecoder {
    internal sealed class BandData {
        internal int[] Primary = Array.Empty<int>(), Alpha = Array.Empty<int>();
        internal int[][][] Quantizers = Array.Empty<int[][]>(), AlphaQuantizers = Array.Empty<int[][]>();
        internal byte[] ModelBits = Array.Empty<byte>(), AlphaModelBits = Array.Empty<byte>();
        internal int[] Patterns = Array.Empty<int>(), AlphaPatterns = Array.Empty<int>();
        internal int[] QuantizerIndices = Array.Empty<int>(), AlphaQuantizerIndices = Array.Empty<int>();
    }

}
