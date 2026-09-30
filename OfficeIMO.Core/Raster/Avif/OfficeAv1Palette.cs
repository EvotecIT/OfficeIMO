namespace OfficeIMO.Drawing;

/// <summary>One leaf's 8-bit palette colors, padded index maps and optional filter-intra mode.</summary>
/// <remarks>Owned arrays are not exposed: publishing a cache cannot be altered by a caller's later writes.</remarks>
internal sealed class OfficeAv1Palette {
    private readonly byte[] _y, _u, _v, _mapY, _mapUv;
    internal OfficeAv1Palette(byte[] y, byte[] u, byte[] v, byte[] mapY, byte[] mapUv,
        int width, int height, int filterMode) {
        _y = y; _u = u; _v = v; _mapY = mapY; _mapUv = mapUv;
        Width = width; Height = height; FilterMode = filterMode;
    }
    internal int SizeY => _y.Length;
    internal int SizeUv => _u.Length;
    internal int Width { get; }
    internal int Height { get; }
    internal int ChromaWidth => System.Math.Max(4, Width / 2);
    internal int ChromaHeight => System.Math.Max(4, Height / 2);
    /// <summary>-1 means filter-intra is not selected; otherwise the AV1 mode is 0..4.</summary>
    internal int FilterMode { get; }
    internal byte Color(int plane, int index) => (plane == 0 ? _y : plane == 1 ? _u : _v)[index];
    internal byte Index(bool chroma, int row, int col) =>
        (chroma ? _mapUv : _mapY)[row * (chroma ? ChromaWidth : Width) + col];
}
