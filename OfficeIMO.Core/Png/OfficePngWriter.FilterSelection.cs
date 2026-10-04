using System;
using System.IO;

namespace OfficeIMO.Drawing;

public static partial class OfficePngWriter {
    // Both size probes and the final stream share the same bounded scanline scratch.
    // A probe may retain a small compressed candidate in the existing IDAT chunk;
    // large candidates still keep only their size and use the final write pass.
    private sealed class PngFilteringWorkspace {
        internal PngFilteringWorkspace(int width, bool bilevel) {
            RgbaStride = checked(width * 4);
            Stride = bilevel ? checked((int)((width + 7L) / 8L)) : RgbaStride;
            Row = new byte[checked(Stride + 1)];
            Paeth = new byte[Stride];
            Batch = new byte[Math.Max(Row.Length, 64 * 1024)];
            BilevelRows = bilevel ? new byte[checked(Stride * 2)] : null;
        }

        internal int RgbaStride { get; }
        internal int Stride { get; }
        internal byte[] Row { get; }
        internal byte[] Paeth { get; }
        internal byte[] Batch { get; }
        internal byte[]? BilevelRows { get; }
    }

    private sealed class PngSizeProbeStream : Stream {
        private long _length;
        private PngIdatChunkStream? _capture;

        internal PngSizeProbeStream(PngIdatChunkStream? capture = null) => _capture = capture;
        internal bool FullyCaptured => _capture != null;
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => _length;
        public override long Position { get => _length; set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override void Write(byte[] buffer, int offset, int count) {
            _length = checked(_length + count);
            if (_capture != null && !_capture.TryCaptureProbe(buffer, offset, count)) DropCapture();
        }

        public override void WriteByte(byte value) {
            _length = checked(_length + 1L);
            if (_capture != null && !_capture.TryCaptureProbeByte(value)) DropCapture();
        }
#if NET8_0_OR_GREATER
        public override void Write(ReadOnlySpan<byte> buffer) {
            _length = checked(_length + buffer.Length);
            if (_capture != null && !_capture.TryCaptureProbe(buffer)) DropCapture();
        }
#endif
        private void DropCapture() {
            _capture!.DiscardProbe();
            _capture = null;
        }
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
    }
}
