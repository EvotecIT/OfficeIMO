#nullable enable

using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Checks the ZIP-declared decompressed length as a package part is consumed.
    /// Closing a partial read does not drain or validate the remaining input.
    /// </summary>
    internal sealed class OpenXmlPartLengthStream : Stream {
        private readonly Stream _input;
        private readonly long _declaredLength;
        private readonly string _partName;
        private long _consumed;

        internal OpenXmlPartLengthStream(Stream input, long declaredLength, string partName) {
            _input = input;
            _declaredLength = declaredLength;
            _partName = partName;
        }

        public override bool CanRead => _input.CanRead;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count) =>
            CheckRead(_input.Read(buffer, offset, count), count != 0);

        public override int ReadByte() {
            int value = _input.ReadByte();
            CheckRead(value < 0 ? 0 : 1, true);
            return value;
        }

        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) {
            int read = await _input.ReadAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false);
            return CheckRead(read, count != 0);
        }

#if NET8_0_OR_GREATER
        public override int Read(Span<byte> buffer) => CheckRead(_input.Read(buffer), !buffer.IsEmpty);

        public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken cancellationToken = default) {
            int read = await _input.ReadAsync(buffer, cancellationToken).ConfigureAwait(false);
            return CheckRead(read, !buffer.IsEmpty);
        }
#endif

        private int CheckRead(int read, bool requestedBytes) {
            _consumed += read;
            if (_consumed > _declaredLength) {
                throw ExcelPackagePartLengthFailure.Create(
                    $"Package part '{_partName}' exceeds its declared decompressed length of {_declaredLength} bytes.");
            }
            if (read == 0 && requestedBytes && _consumed != _declaredLength) {
                // Integrity failures must propagate through the XML/SDK fallback
                // routes, which can recover from ordinary IO failures.
                throw ExcelPackagePartLengthFailure.Create(
                    $"Package part '{_partName}' ended after {_consumed} of {_declaredLength} declared bytes.");
            }
            return read;
        }

        public override void Flush() => _input.Flush();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing) {
            try {
                if (disposing) _input.Dispose();
            } finally {
                base.Dispose(disposing);
            }
        }
    }
}
