using System;
using System.IO;
using System.IO.Compression;
using System.Threading;

namespace OfficeIMO.Drawing;

public static partial class OfficePngWriter {
    private const int MaterializedProbeMinimumRgbaBytes = 1024 * 1024;
    private const int MaterializedProbeMaximumPayloadBytes = 4 * 1024 * 1024;

    private static byte[] EncodeRgbaMaterialized(
        int width, int height, byte[] rgba, OfficePngCompression compression,
        double? dpiX, double? dpiY, CancellationToken cancellationToken) {
        using var output = new MemoryStream();
        EncodeRgbaStreaming(width, height, rgba, output, compression, dpiX, dpiY,
            cancellationToken, ownedOutput: output);
        return output.ToArray();
    }

    private static void WriteMaterializedOptimalZlib(
        MemoryStream output, PngIdatChunkStream destination, int width, int height,
        byte[] rgba, CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver) {
        long idatStart = output.Position;
        var workspace = new PngFilteringWorkspace(width, bilevel: false);
        using var adaptiveSize = new PngForwardProbeStream(destination, MaterializedProbeMaximumPayloadBytes);
        WriteRgbaZlib(adaptiveSize, height, rgba, workspace, adaptiveFiltering: true,
            cancellationToken, checkpointObserver);

        // A complete provisional candidate lives only in our own output. Once the
        // retention limit is exceeded, finish counting without retaining further bytes.
        if (!adaptiveSize.FullyRetained) RewindMaterializedOutput(output, idatStart);
        destination.Complete();

        using var unfilteredSize = new PngSizeProbeStream(destination);
        WriteBoundedUnfilteredProbe(unfilteredSize, height, rgba, workspace,
            adaptiveSize.Length, cancellationToken, checkpointObserver);
        bool adaptiveFiltering = adaptiveSize.Length <= unfilteredSize.Length;
        cancellationToken.ThrowIfCancellationRequested();
        if (adaptiveFiltering && adaptiveSize.FullyRetained) {
            destination.DiscardProbe();
            return;
        }

        // Replacing a provisional candidate must remove its complete chunk framing,
        // including any longer tail, without touching the already-written PNG metadata.
        RewindMaterializedOutput(output, idatStart);
        destination.ResumeWriting();
        if (!adaptiveFiltering && unfilteredSize.FullyCaptured) return;
        destination.DiscardProbe();
        WriteRgbaZlib(destination, height, rgba, workspace, adaptiveFiltering,
            cancellationToken, checkpointObserver);
    }

    // This destination only counts/captures a size probe; it never writes the final
    // PNG. Once emitted bytes exceed the complete adaptive candidate, remaining
    // rows cannot make unfiltered compression win. Finish deflate normally and
    // discard the losing partial probe through the existing selection logic.
    private static void WriteBoundedUnfilteredProbe(
        PngSizeProbeStream destination, int height, byte[] rgba,
        PngFilteringWorkspace workspace, long adaptiveLength,
        CancellationToken cancellationToken,
        Action<OfficeRasterEncodingCheckpoint>? checkpointObserver) {
        destination.WriteByte(0x78);
        destination.WriteByte(0x9C);
        byte[] row = workspace.Row;
        byte[] batch = workspace.Batch;
        int stride = workspace.Stride;
        int batchLength = 0;
        uint adlerA = 1;
        uint adlerB = 0;

        using (var deflate = new DeflateStream(destination, CompressionLevel.Optimal, leaveOpen: true)) {
            for (int y = 0; y < height; y++) {
                checkpointObserver?.Invoke(OfficeRasterEncodingCheckpoint.PngCompressionRow);
                cancellationToken.ThrowIfCancellationRequested();
                row[0] = 0;
                Buffer.BlockCopy(rgba, y * stride, row, 1, stride);
                if (row.Length > batch.Length - batchLength) {
                    deflate.Write(batch, 0, batchLength);
                    batchLength = 0;
                    if (destination.Length > adaptiveLength) break;
                }
                Buffer.BlockCopy(row, 0, batch, batchLength, row.Length);
                batchLength += row.Length;
                UpdateAdler32(row, 0, row.Length, ref adlerA, ref adlerB,
                    cancellationToken, checkpointObserver);
            }
            if (batchLength > 0) deflate.Write(batch, 0, batchLength);
        }
        WriteAdler32(destination, (adlerB << 16) | adlerA);
    }

    private static void RewindMaterializedOutput(MemoryStream output, long position) {
        output.SetLength(position);
        output.Position = position;
    }

    // Count exactly the same zlib stream as the ordinary probe while retaining a
    // bounded candidate through the existing IDAT writer. No second payload buffer
    // is allocated, and no caller-owned or byte-budgeted stream enters this path.
    private sealed class PngForwardProbeStream : Stream {
        private PngIdatChunkStream? _forward;
        private readonly int _maximumPayloadBytes;
        private long _length;

        internal PngForwardProbeStream(PngIdatChunkStream destination, int maximumPayloadBytes) {
            _forward = destination;
            _maximumPayloadBytes = maximumPayloadBytes;
        }

        internal bool FullyRetained => _forward != null;
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => _length;
        public override long Position { get => _length; set => throw new NotSupportedException(); }
        public override void Flush() { }

        private bool Retain(int count) {
            _length = checked(_length + count);
            if (_forward == null) return false;
            if (_length <= _maximumPayloadBytes) return true;
            _forward.DiscardProbe();
            _forward = null;
            return false;
        }

        public override void Write(byte[] buffer, int offset, int count) {
            if (Retain(count)) _forward!.Write(buffer, offset, count);
        }

        public override void WriteByte(byte value) {
            if (Retain(1)) _forward!.WriteByte(value);
        }

#if NET8_0_OR_GREATER
        public override void Write(ReadOnlySpan<byte> buffer) {
            if (Retain(buffer.Length)) _forward!.Write(buffer);
        }
#endif
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
    }
}
