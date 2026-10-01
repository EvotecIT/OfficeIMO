using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Tile-owned Main-8 palette/filter-intra syntax and palette tokens through the transform-size boundary.</summary>
/// <remarks>The caller supplies the parsed modes and follows partition order. CompleteBlock publishes palettes
/// only after the remaining transform/residual syntax succeeds. Above color reuse resets on each 64-pixel row,
/// independently of the above palette-presence context. Intra-block-copy motion remains a separate owner.</remarks>
internal sealed partial class OfficeAv1PaletteReader {
    private readonly OfficeAv1TileGeometry _geometry;
    private readonly bool _screen, _filter;
    private readonly CancellationToken _cancellation;
    private readonly byte[] _aboveSizes, _leftSizes;
    private readonly ushort[] _aboveColors, _leftColors;
    private readonly int[][] _hasY = CreateHasY(), _hasUv = CreateHasUv(), _sizeY = CreateSizeY(), _sizeUv = CreateSizeUv();
    private readonly int[][] _indexY = CreateIndexY(), _indexUv = CreateIndexUv(), _useFilter = CreateUseFilter();
    private readonly int[] _filterModes = { 8949, 12776, 17211, 29558, 32768, 0 };
    private readonly ushort[] _cache = new ushort[16];
    private readonly int[] _scores = new int[8], _order = new int[8];
    private bool _pending, _failed;
    private OfficeAv1BlockRegion _block;
    private OfficeAv1Palette? _palette;

    internal OfficeAv1PaletteReader(OfficeAv1StillFrame frame, OfficeAv1StillSequence sequence, OfficeAv1Tile tile,
        OfficeRasterDecodeOptions options) {
        if (sequence == null) throw new ArgumentNullException(nameof(sequence));
        if (options == null) throw new ArgumentNullException(nameof(options));
        options.Validate(); options.CancellationToken.ThrowIfCancellationRequested();
        _geometry = new OfficeAv1TileGeometry(frame, tile, sequence.Use128Superblock ? 128 : 64);
        _geometry.EnsurePixelBudget(options.MaximumDecodedPixels);
        _screen = frame.AllowScreenContentTools; _filter = sequence.FilterIntra; _cancellation = options.CancellationToken;
        const string limit = "AV1 palette context exceeds the retained-memory limit.";
        // Includes fixed CDF/scratch storage and the largest one-leaf colors/maps (64x64 and 32x32).
        const int fixedBytes = 16384;
        if (options.RetainedManagedBytes > OfficeRasterGuards.MaximumDecodedBytes - fixedBytes) throw new FormatException(limit);
        long retained = options.RetainedManagedBytes + fixedBytes;
        int cols = tile.MiColEnd - tile.MiColStart, rows = tile.MiRowEnd - tile.MiRowStart;
        _aboveSizes = new byte[OfficeRasterGuards.EnsureByteArrayLength(cols * 2, ref retained, limit)];
        _leftSizes = new byte[OfficeRasterGuards.EnsureByteArrayLength(rows * 2, ref retained, limit)];
        _aboveColors = new ushort[OfficeRasterGuards.EnsureByteArrayLength(cols * 32, ref retained, limit) / 2];
        _leftColors = new ushort[OfficeRasterGuards.EnsureByteArrayLength(rows * 32, ref retained, limit) / 2];
    }

    internal OfficeAv1Palette Read(OfficeAv1SymbolReader symbols, OfficeAv1BlockRegion block, OfficeAv1IntraModes modes) {
        Check();
        if (symbols == null) throw new ArgumentNullException(nameof(symbols));
        _geometry.Validate(block);
        if (_pending) throw new InvalidOperationException("Complete the current AV1 palette leaf first.");
        try {
            ushort[] y = Array.Empty<ushort>(), u = Array.Empty<ushort>(), v = Array.Empty<ushort>();
            int filterMode = -1;
            if (!modes.UseIntraBlockCopy) {
                bool paletteAllowed = _screen && block.Width * block.Height >= 64 && block.Width <= 64 && block.Height <= 64;
                if (paletteAllowed) {
                    int sizeContext = Log2(block.Width / 4) + Log2(block.Height / 4) - 2;
                    if (modes.YMode == OfficeAv1IntraMode.Dc) {
                        int context = (AboveSize(0, block) > 0 ? 1 : 0) + (LeftSize(0, block) > 0 ? 1 : 0);
                        if (symbols.ReadSymbol(_hasY[sizeContext * 3 + context]) != 0)
                            y = ReadColors(symbols, block, 0, symbols.ReadSymbol(_sizeY[sizeContext]) + 2);
                    }
                    if (modes.HasChroma && modes.UvMode == OfficeAv1IntraMode.Dc && symbols.ReadSymbol(_hasUv[y.Length > 0 ? 1 : 0]) != 0) {
                        int count = symbols.ReadSymbol(_sizeUv[sizeContext]) + 2;
                        u = ReadColors(symbols, block, 1, count);
                        v = ReadVColors(symbols, count);
                    }
                }
                if (_filter && modes.YMode == OfficeAv1IntraMode.Dc && y.Length == 0 && Math.Max(block.Width, block.Height) <= 32 &&
                    symbols.ReadSymbol(_useFilter[BlockSizeIndex(block)]) != 0)
                    filterMode = symbols.ReadSymbol(_filterModes);
            }
            int onWidth = Math.Min(block.Width, (_geometry.MiCols - block.MiCol) * 4);
            int onHeight = Math.Min(block.Height, (_geometry.MiRows - block.MiRow) * 4);
            byte[] mapY = y.Length == 0 ? Array.Empty<byte>() : ReadMap(symbols, y.Length, block.Width, block.Height, onWidth, onHeight, _indexY);
            int uvWidth = Math.Max(4, block.Width / 2), uvHeight = Math.Max(4, block.Height / 2);
            int onUvWidth = onWidth / 2 + (block.Width < 8 ? 2 : 0), onUvHeight = onHeight / 2 + (block.Height < 8 ? 2 : 0);
            byte[] mapUv = u.Length == 0 ? Array.Empty<byte>() : ReadMap(symbols, u.Length, uvWidth, uvHeight, onUvWidth, onUvHeight, _indexUv);
            _block = block; _palette = new OfficeAv1Palette(y, u, v, mapY, mapUv, block.Width, block.Height, filterMode);
            _pending = true;
            return _palette;
        } catch { _failed = true; throw; }
    }

    internal void CompleteBlock() {
        Check();
        if (!_pending) throw new InvalidOperationException("No AV1 palette leaf is pending completion.");
        var tile = _geometry.Tile; var b = _block; var palette = _palette!;
        for (int plane = 0; plane < 2; plane++) {
            int count = plane == 0 ? palette.SizeY : palette.SizeUv;
            for (int col = b.MiCol; col < Math.Min(tile.MiColEnd, b.MiCol + b.Width / 4); col++)
                Publish(_aboveSizes, _aboveColors, col - tile.MiColStart, plane, count, palette);
            for (int row = b.MiRow; row < Math.Min(tile.MiRowEnd, b.MiRow + b.Height / 4); row++)
                Publish(_leftSizes, _leftColors, row - tile.MiRowStart, plane, count, palette);
        }
        _pending = false; _palette = null;
    }

    private static void Publish(byte[] sizes, ushort[] colors, int cell, int plane, int count, OfficeAv1Palette palette) {
        sizes[cell * 2 + plane] = (byte)count;
        for (int i = 0; i < count; i++) colors[cell * 16 + plane * 8 + i] = palette.Color(plane, i);
    }
    private int AboveSize(int plane, OfficeAv1BlockRegion b) => b.MiRow > _geometry.Tile.MiRowStart ? _aboveSizes[(b.MiCol - _geometry.Tile.MiColStart) * 2 + plane] : 0;
    private int LeftSize(int plane, OfficeAv1BlockRegion b) => b.MiCol > _geometry.Tile.MiColStart ? _leftSizes[(b.MiRow - _geometry.Tile.MiRowStart) * 2 + plane] : 0;
    private static int Log2(int value) { int n = 0; while ((value >>= 1) != 0) n++; return n; }
    private static int BlockSizeIndex(OfficeAv1BlockRegion b) {
        // AV1 BLOCK_SIZE numeric order; geometry validation has already checked the dimensions.
        switch (b.Width * 256 + b.Height) {
            case 4 * 256 + 4: return 0;
            case 4 * 256 + 8: return 1;
            case 8 * 256 + 4: return 2;
            case 8 * 256 + 8: return 3;
            case 8 * 256 + 16: return 4;
            case 16 * 256 + 8: return 5;
            case 16 * 256 + 16: return 6;
            case 16 * 256 + 32: return 7;
            case 32 * 256 + 16: return 8;
            case 32 * 256 + 32: return 9;
            case 32 * 256 + 64: return 10;
            case 64 * 256 + 32: return 11;
            case 64 * 256 + 64: return 12;
            case 64 * 256 + 128: return 13;
            case 128 * 256 + 64: return 14;
            case 128 * 256 + 128: return 15;
            case 4 * 256 + 16: return 16;
            case 16 * 256 + 4: return 17;
            case 8 * 256 + 32: return 18;
            case 32 * 256 + 8: return 19;
            case 16 * 256 + 64: return 20;
            case 64 * 256 + 16: return 21;
            default: throw new FormatException("Invalid AV1 block size.");
        }
    }
    private void Check() {
        _cancellation.ThrowIfCancellationRequested();
        if (_failed) throw new FormatException("The AV1 palette context has failed.");
    }
}
