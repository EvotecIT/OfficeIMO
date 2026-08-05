using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Keeps package creation forward-only while satisfying .NET Framework's
    /// <see cref="System.IO.Compression.ZipArchive"/> implementation, which reads
    /// <see cref="Stream.Position"/> even when the destination cannot seek.
    /// </summary>
    internal sealed class ExcelPositionReportingWriteStream : Stream {
        private readonly Stream _destination;
        private long _position;

        internal ExcelPositionReportingWriteStream(Stream destination) {
            _destination = destination ?? throw new ArgumentNullException(nameof(destination));
            if (!destination.CanWrite) {
                throw new ArgumentException("The destination stream must be writable.", nameof(destination));
            }
        }

        public override bool CanRead => false;
        public override bool CanSeek => false;
        public override bool CanWrite => _destination.CanWrite;
        public override long Length => _position;

        public override long Position {
            get => _position;
            set => throw new NotSupportedException();
        }

        public override void Flush() => _destination.Flush();

        public override Task FlushAsync(CancellationToken cancellationToken) =>
            _destination.FlushAsync(cancellationToken);

        public override int Read(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) {
            _destination.Write(buffer, offset, count);
            _position = checked(_position + count);
        }

        public override async Task WriteAsync(
            byte[] buffer,
            int offset,
            int count,
            CancellationToken cancellationToken) {
            await _destination.WriteAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false);
            _position = checked(_position + count);
        }
    }
}
