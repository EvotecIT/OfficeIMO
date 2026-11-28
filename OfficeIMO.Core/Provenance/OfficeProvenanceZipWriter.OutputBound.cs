namespace OfficeIMO.Provenance;

internal static partial class OfficeProvenanceZipWriter {
    private sealed class ZipOutputBound(Stream destination, long maximumBytes) : Stream {
        public override bool CanRead => false;
        public override bool CanSeek => true;
        public override bool CanWrite => true;
        public override long Length => destination.Length;
        public override long Position { get => destination.Position; set => destination.Position = value; }
        public override void Flush() => destination.Flush();
        public override long Seek(long offset, SeekOrigin origin) => destination.Seek(offset, origin);

        public override void Write(byte[] buffer, int offset, int count) {
            CheckBound(count);
            destination.Write(buffer, offset, count);
        }

#if NET8_0_OR_GREATER
        public override void Write(ReadOnlySpan<byte> buffer) {
            CheckBound(buffer.Length);
            destination.Write(buffer);
        }
#endif

        private void CheckBound(int count) {
            if (count > maximumBytes - destination.Position)
                throw OfficeProvenanceLimitException.CreateOutput(
                    $"The rewritten asset exceeds the configured output limit of {maximumBytes} bytes.");
        }

        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
    }
}
