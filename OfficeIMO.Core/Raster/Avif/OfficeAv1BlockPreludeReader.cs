using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>Tile-local reduced-still segment, skip, CDEF and delta state (AV1 5.11.7–13, 5.11.56).</summary>
/// <remarks>The decoder calls BeginSuperblock in tile raster order, Read in partition decode order, then
/// CompleteBlock after all remaining leaf syntax succeeds. No neighbor state is published by Read.</remarks>
internal sealed partial class OfficeAv1BlockPreludeReader {
    private readonly OfficeAv1Tile _tile;
    private readonly OfficeAv1TileGeometry _geometry;
    private readonly int _superblockUnits, _columns;
    private readonly bool _segmentation, _preSkip, _cdef, _deltaQ, _deltaLf, _multiLf;
    private readonly int _lastSegment, _cdefBits, _qResolution, _lfResolution, _filterCount;
    private readonly byte _skipSegments, _losslessSegments;
    private readonly byte[] _aboveSkip, _leftSkip;
    private readonly byte[]? _segments;
    private readonly int[] _cdefIndices = { -1, -1, -1, -1 }, _filters = new int[4];
    private readonly int[][] _skipCdfs = { new[] {31671,32768,0}, new[] {16515,32768,0}, new[] {4576,32768,0} };
    private readonly int[][] _segmentCdfs = {
        new[] {5622,7893,16093,18233,27809,28373,32533,32768,0},
        new[] {14274,18230,22557,24935,29980,30851,32344,32768,0},
        new[] {27527,28487,28723,28890,32397,32647,32679,32768,0}
    };
    private readonly int[] _qCdf = DeltaCdf(), _lfCdf = DeltaCdf();
    private readonly int[][] _multiLfCdfs = { DeltaCdf(), DeltaCdf(), DeltaCdf(), DeltaCdf() };
    private readonly CancellationToken _cancellation;
    private int _q, _superblockRow = -1, _superblockCol = -1;
    private bool _readDeltas, _pending, _failed;
    private OfficeAv1BlockRegion _pendingBlock;
    private OfficeAv1BlockPrelude _pendingState;

    internal OfficeAv1BlockPreludeReader(OfficeAv1StillFrame frame, OfficeAv1StillSequence sequence,
        OfficeAv1Tile tile, OfficeRasterDecodeOptions options) {
        if (frame == null) throw new ArgumentNullException(nameof(frame));
        if (sequence == null) throw new ArgumentNullException(nameof(sequence));
        if (options == null) throw new ArgumentNullException(nameof(options));
        options.Validate(); options.CancellationToken.ThrowIfCancellationRequested();
        _superblockUnits = sequence.Use128Superblock ? 32 : 16;
        _geometry = new OfficeAv1TileGeometry(frame,tile,_superblockUnits*4);
        _geometry.EnsurePixelBudget(options.MaximumDecodedPixels);
        if (frame.BaseQIndex < 0 || frame.BaseQIndex > 255 || frame.LastActiveSegmentId < 0 || frame.LastActiveSegmentId > 7 ||
            frame.CdefBits < 0 || frame.CdefBits > 3 || frame.DeltaQResolution < 0 || frame.DeltaQResolution > 3 ||
            frame.DeltaLoopFilterResolution < 0 || frame.DeltaLoopFilterResolution > 3)
            throw new FormatException("Invalid AV1 block prelude geometry or parameters.");
        _tile = tile;
        _columns = tile.MiColEnd - tile.MiColStart; _cancellation = options.CancellationToken;
        _segmentation = frame.SegmentationEnabled; _preSkip = frame.SegmentIdPreSkip;
        _lastSegment = frame.LastActiveSegmentId; _q = frame.BaseQIndex;
        _cdef = sequence.Cdef && !frame.CodedLossless && !frame.AllowIntraBlockCopy; _cdefBits = frame.CdefBits;
        _deltaQ = frame.DeltaQPresent; _qResolution = frame.DeltaQResolution;
        _deltaLf = frame.DeltaLoopFilterPresent; _lfResolution = frame.DeltaLoopFilterResolution;
        _multiLf = frame.DeltaLoopFilterMulti; _filterCount = _multiLf ? sequence.Monochrome ? 2 : 4 : 1;
        for (int i = 0; i < 8; i++) {
            if (_segmentation && frame.SegmentFeatures[i, 6]) _skipSegments |= (byte)(1 << i);
            if (frame.LosslessSegments[i]) _losslessSegments |= (byte)(1 << i);
        }
        const string limit = "AV1 block context exceeds the retained-memory limit.";
        if (options.RetainedManagedBytes > OfficeRasterGuards.MaximumDecodedBytes - 1024) throw new FormatException(limit);
        long retained = options.RetainedManagedBytes + 1024; // bounded CDF and fixed state payloads
        int rows = tile.MiRowEnd - tile.MiRowStart;
        int aboveLength = OfficeRasterGuards.EnsureByteArrayLength(_columns, ref retained, limit);
        int leftLength = OfficeRasterGuards.EnsureByteArrayLength(rows, ref retained, limit);
        int segmentLength = _segmentation ? OfficeRasterGuards.EnsureByteArrayLength((long)rows * _columns, ref retained, limit) : 0;
        _aboveSkip = new byte[aboveLength]; _leftSkip = new byte[leftLength];
        _segments = _segmentation ? new byte[segmentLength] : null;
    }

    /// <summary>Resets CDEF regions and the first-leaf delta flag; running quantizer/filter deltas persist across superblocks.</summary>
    internal void BeginSuperblock(int row, int col) {
        Check();
        if (_pending || row < _tile.MiRowStart || row >= _tile.MiRowEnd || col < _tile.MiColStart || col >= _tile.MiColEnd ||
            row % _superblockUnits != 0 || col % _superblockUnits != 0)
            throw new FormatException("Invalid AV1 superblock boundary.");
        _superblockRow = row; _superblockCol = col; _readDeltas = _deltaQ;
        for (int i = 0; i < 4; i++) _cdefIndices[i] = -1;
    }

    internal OfficeAv1BlockPrelude Read(OfficeAv1SymbolReader symbols, OfficeAv1BlockRegion block) {
        Check();
        if (symbols == null) throw new ArgumentNullException(nameof(symbols));
        ValidateBlock(block);
        if (_pending) throw new InvalidOperationException("Complete the current AV1 leaf before reading another.");
        try {
            bool skip = false;
            int segment = _preSkip ? ReadSegment(symbols, block, false) : 0;
            if (_preSkip && (_skipSegments & (1 << segment)) != 0) skip = true;
            else {
                int context = (block.MiRow > _tile.MiRowStart ? _aboveSkip[block.MiCol - _tile.MiColStart] : 0) +
                    (block.MiCol > _tile.MiColStart ? _leftSkip[block.MiRow - _tile.MiRowStart] : 0);
                skip = symbols.ReadSymbol(_skipCdfs[context]) != 0;
            }
            if (!_preSkip) segment = ReadSegment(symbols, block, skip);
            ReadCdef(symbols, block, skip);
            if (_readDeltas && !(skip && block.Width == _superblockUnits * 4 && block.Height == _superblockUnits * 4)) {
                _q = Clip(_q + (ReadDelta(symbols, _qCdf) << _qResolution), 1, 255);
                if (_deltaLf) for (int i = 0; i < _filterCount; i++)
                    _filters[i] = Clip(_filters[i] + (ReadDelta(symbols, _multiLf ? _multiLfCdfs[i] : _lfCdf) << _lfResolution), -63, 63);
            }
            _readDeltas = false;
            _pendingBlock = block;
            _pendingState = new OfficeAv1BlockPrelude(skip, segment, (_losslessSegments & (1 << segment)) != 0,
                _cdefIndices[CdefSlot(block.MiRow, block.MiCol)], _q, _filters);
            _pending = true;
            return _pendingState;
        } catch { _failed = true; throw; }
    }

    /// <summary>Publishes skip/segment neighbors only after prediction and residual syntax have also succeeded.</summary>
    internal void CompleteBlock() {
        Check();
        if (!_pending) throw new InvalidOperationException("No AV1 leaf is pending completion.");
        int row = _pendingBlock.MiRow, col = _pendingBlock.MiCol;
        int endRow = Math.Min(_tile.MiRowEnd, row + _pendingBlock.Height / 4);
        int endCol = Math.Min(_tile.MiColEnd, col + _pendingBlock.Width / 4);
        for (int c = col; c < endCol; c++) _aboveSkip[c - _tile.MiColStart] = _pendingState.Skip ? (byte)1 : (byte)0;
        for (int r = row; r < endRow; r++) {
            _cancellation.ThrowIfCancellationRequested();
            _leftSkip[r - _tile.MiRowStart] = _pendingState.Skip ? (byte)1 : (byte)0;
            if (_segments != null) for (int c = col; c < endCol; c++) _segments[SegmentIndex(r, c)] = (byte)_pendingState.SegmentId;
        }
        _pending = false;
    }

    private void Check() {
        _cancellation.ThrowIfCancellationRequested();
        if (_failed) throw new FormatException("The AV1 leaf context has failed.");
    }
    private void ValidateBlock(OfficeAv1BlockRegion b) {
        _geometry.Validate(b);
        if (_superblockRow < 0 || b.MiRow < _superblockRow || b.MiCol < _superblockCol ||
            b.MiRow+b.Height/4 > _superblockRow+_superblockUnits || b.MiCol+b.Width/4 > _superblockCol+_superblockUnits)
            throw new FormatException("Invalid AV1 leaf superblock geometry.");
    }
    private static int[] DeltaCdf() => new[] {28160,32120,32677,32768,0};
    private static int Clip(int value, int minimum, int maximum) => Math.Max(minimum, Math.Min(maximum, value));
}
