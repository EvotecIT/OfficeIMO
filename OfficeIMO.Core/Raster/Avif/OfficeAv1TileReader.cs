using System;
using System.Threading;

namespace OfficeIMO.Drawing;

/// <summary>One-shot Main-8 reduced-still tile traversal from restoration units through entropy termination.</summary>
/// <remarks>Each tile owns all mutable probability and neighbor contexts. A consumer receives one leaf's
/// immutable syntax and residuals at a time; pixel prediction, inverse transforms and filters are separate.
/// Failure, cancellation or a consumer exception permanently retires this reader and its partial output.</remarks>
internal sealed class OfficeAv1TileReader {
    private readonly OfficeAv1TileGeometry _geometry;
    private readonly OfficeAv1PartitionContext _partitions;
    private readonly OfficeAv1BlockPreludeReader _preludes;
    private readonly OfficeAv1IntraModeReader _modes;
    private readonly OfficeAv1CopyMotionReader? _motion;
    private readonly OfficeAv1PaletteReader _palettes;
    private readonly OfficeAv1TransformReader _transforms;
    private readonly OfficeAv1CoefficientReader _coefficients;
    private readonly OfficeAv1RestorationReader _restoration;
    private readonly CancellationToken _cancellation;
    private readonly bool _disableCdfUpdate;
    private readonly int _maximumEncodedBytes;
    private readonly long _retainedContextBytes;
    private bool _started;

    internal OfficeAv1TileReader(OfficeAv1StillFrame frame, OfficeAv1StillSequence sequence,
        OfficeAv1Tile tile, OfficeRasterDecodeOptions options) {
        if(sequence==null) throw new ArgumentNullException(nameof(sequence));
        if(options==null) throw new ArgumentNullException(nameof(options));
        options.Validate(); options.CancellationToken.ThrowIfCancellationRequested();
        _geometry=new OfficeAv1TileGeometry(frame,tile,sequence.Use128Superblock?128:64);
        _geometry.EnsurePixelBudget(options.MaximumDecodedPixels);
        _cancellation=options.CancellationToken; _disableCdfUpdate=frame.DisableCdfUpdate;
        _maximumEncodedBytes=options.MaximumEncodedBytes;
        long axis=(long)(tile.MiColEnd-tile.MiColStart)+(tile.MiRowEnd-tile.MiRowStart);
        long area=(long)(tile.MiColEnd-tile.MiColStart)*(tile.MiRowEnd-tile.MiRowStart);
        // Reserve all simultaneously live contexts, including one-leaf temporaries, before allocating any.
        // The partition owner has no options parameter; its CDFs and border storage are included here.
        long partitionBytes=4096+axis, preludeBytes=1024+axis+(frame.SegmentationEnabled?area:0);
        long modeBytes=8192+axis, paletteBytes=16384+18*axis, transformBytes=65536+3*axis;
        long coefficientBytes=196608+4*axis, motionBytes=frame.AllowIntraBlockCopy?4096+8*area:0;
        const long restorationBytes=2048;
        long all=partitionBytes+preludeBytes+modeBytes+paletteBytes+transformBytes+coefficientBytes+motionBytes+restorationBytes;
        if(options.RetainedManagedBytes>OfficeRasterGuards.MaximumDecodedBytes-all)
            throw new FormatException("AV1 tile contexts exceed the retained-memory limit.");
        _retainedContextBytes=options.RetainedManagedBytes+all;
        _partitions=new OfficeAv1PartitionContext(frame,tile,_geometry.SuperblockPixels,_cancellation);
        _preludes=new OfficeAv1BlockPreludeReader(frame,sequence,tile,options.WithAdditionalRetainedManagedBytes(all-preludeBytes));
        _modes=new OfficeAv1IntraModeReader(frame,sequence,tile,options.WithAdditionalRetainedManagedBytes(all-modeBytes));
        _palettes=new OfficeAv1PaletteReader(frame,sequence,tile,options.WithAdditionalRetainedManagedBytes(all-paletteBytes));
        _transforms=new OfficeAv1TransformReader(frame,sequence,tile,options.WithAdditionalRetainedManagedBytes(all-transformBytes));
        _coefficients=new OfficeAv1CoefficientReader(frame,sequence,tile,options.WithAdditionalRetainedManagedBytes(all-coefficientBytes));
        _restoration=new OfficeAv1RestorationReader(frame,sequence,tile,options.WithAdditionalRetainedManagedBytes(all-restorationBytes));
        if(frame.AllowIntraBlockCopy)
            _motion=new OfficeAv1CopyMotionReader(frame,sequence,tile,options.WithAdditionalRetainedManagedBytes(all-motionBytes));
    }

    /// <summary>Consumes the bounded tile once. CompleteTile is called only after the mandatory trailing bits pass.</summary>
    internal void Read(byte[] bytes, IOfficeAv1TileConsumer consumer) {
        if(bytes==null) throw new ArgumentNullException(nameof(bytes));
        if(consumer==null) throw new ArgumentNullException(nameof(consumer));
        if(_started) throw new InvalidOperationException("The AV1 tile reader cannot be resumed or reused.");
        _started=true; _cancellation.ThrowIfCancellationRequested();
        if(bytes.Length>_maximumEncodedBytes) throw new FormatException("AV1 tile input exceeds the encoded-byte limit.");
        // The entropy reader retains the backing array, including bytes outside this tile's slice.
        if(bytes.LongLength>OfficeRasterGuards.MaximumDecodedBytes-_retainedContextBytes)
            throw new FormatException("AV1 tile input and contexts exceed the retained-memory limit.");
        var tile=_geometry.Tile;
        // One coefficient can require up to 20 Golomb prefix bits plus 19 suffix bits and sign/base syntax.
        // This conservative per-padded-pixel bound also covers partition, mode, palette and restoration syntax.
        long maximumSymbols=(long)(tile.MiRowEnd-tile.MiRowStart)*(tile.MiColEnd-tile.MiColStart)*16*128+4096;
        var symbols=new OfficeAv1SymbolReader(bytes,tile.Offset,tile.Length,maximumSymbols,_disableCdfUpdate,_cancellation);
        int units=_geometry.SuperblockPixels/4;
        for(int row=tile.MiRowStart;row<tile.MiRowEnd;row+=units)
            for(int col=tile.MiColStart;col<tile.MiColEnd;col+=units) {
                _cancellation.ThrowIfCancellationRequested();
                _preludes.BeginSuperblock(row,col);
                _restoration.ReadSuperblock(symbols,row,col,consumer);
                ReadPartition(symbols,row,col,_geometry.SuperblockPixels,consumer);
            }
        symbols.Finish(); _cancellation.ThrowIfCancellationRequested(); consumer.CompleteTile();
    }

    private void ReadPartition(OfficeAv1SymbolReader symbols,int row,int col,int pixels,IOfficeAv1TileConsumer consumer) {
        _cancellation.ThrowIfCancellationRequested();
        var layout=_partitions.Read(symbols,row,col,pixels);
        for(int i=0;i<layout.Count;i++) {
            var block=layout.Child(i);
            if(block.MiRow>=_geometry.MiRows || block.MiCol>=_geometry.MiCols) continue;
            if(layout.Kind==OfficeAv1Partition.Split) ReadPartition(symbols,block.MiRow,block.MiCol,block.Width,consumer);
            else ReadBlock(symbols,block,consumer);
        }
    }

    private void ReadBlock(OfficeAv1SymbolReader symbols,OfficeAv1BlockRegion block,IOfficeAv1TileConsumer consumer) {
        var prelude=_preludes.Read(symbols,block);
        var modes=_modes.Read(symbols,block,prelude);
        var motion=_motion==null?default:_motion.Read(symbols,block,modes);
        var palette=_palettes.Read(symbols,block,modes);
        var transforms=_transforms.Read(symbols,block,prelude,modes);
        _coefficients.BeginBlock(block,prelude,modes,palette,transforms);
        consumer.BeginBlock(new OfficeAv1TileBlock(block,prelude,modes,motion,palette,transforms));
        for(int i=0;i<transforms.Count;i++) {
            _cancellation.ThrowIfCancellationRequested();
            consumer.Residual(_coefficients.Read(symbols));
        }
        _cancellation.ThrowIfCancellationRequested(); consumer.EndBlock(); _cancellation.ThrowIfCancellationRequested();
        _coefficients.CompleteBlock(); _transforms.CompleteBlock(); _palettes.CompleteBlock();
        _motion?.CompleteBlock(); _modes.CompleteBlock(); _preludes.CompleteBlock(); _partitions.RecordBlock(block);
    }
}
