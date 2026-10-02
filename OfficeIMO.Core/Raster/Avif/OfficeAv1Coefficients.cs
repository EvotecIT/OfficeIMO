namespace OfficeIMO.Drawing;

/// <summary>AV1 transform identifiers in normative TX_TYPE order.</summary>
internal enum OfficeAv1TransformType {
    DctDct, AdstDct, DctAdst, AdstAdst, FlipAdstDct, DctFlipAdst, FlipAdstFlipAdst,
    AdstFlipAdst, FlipAdstAdst, Identity, VerticalDct, HorizontalDct, VerticalAdst,
    HorizontalAdst, VerticalFlipAdst, HorizontalFlipAdst
}

/// <summary>Immutable signed quantized coefficients in specification row-major order, before dequantization.</summary>
internal sealed class OfficeAv1Coefficients {
    private readonly int[] _values;
    internal OfficeAv1Coefficients(OfficeAv1TransformBlock block, OfficeAv1TransformType type, int eob, int[] values) {
        Block=block; Type=type; EndOfBlock=eob; _values=values;
    }
    internal OfficeAv1TransformBlock Block { get; }
    internal OfficeAv1TransformType Type { get; }
    internal int EndOfBlock { get; }
    internal int Width => System.Math.Min(32,Block.Width);
    internal int Height => System.Math.Min(32,Block.Height);
    internal int Count => _values.Length;
    internal int Value(int index) => _values[index];
}
