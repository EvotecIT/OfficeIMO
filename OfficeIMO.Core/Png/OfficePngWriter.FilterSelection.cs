using System;
using System.IO;

namespace OfficeIMO.Drawing;

public static partial class OfficePngWriter {
    // Both size probes and the final stream share the same bounded scanline scratch.
    // No candidate image or compressed payload is retained by the probes.
    private sealed class PngFilteringWorkspace {
        internal PngFilteringWorkspace(int width) {
            Stride = checked(width * 4);
            Row = new byte[checked(Stride + 1)];
            Paeth = new byte[Stride];
            Batch = new byte[Math.Max(Row.Length, 64 * 1024)];
        }

        internal int Stride { get; }
        internal byte[] Row { get; }
        internal byte[] Paeth { get; }
        internal byte[] Batch { get; }
    }

    private sealed class PngSizeProbeStream : Stream {
        private long _length;
        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => true;
        public override long Length => _length;
        public override long Position { get => _length; set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override void Write(byte[] buffer, int offset, int count) => _length = checked(_length + count);
        public override void WriteByte(byte value) => _length = checked(_length + 1L);
#if NET8_0_OR_GREATER
        public override void Write(ReadOnlySpan<byte> buffer) => _length = checked(_length + buffer.Length);
#endif
        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
    }
}
