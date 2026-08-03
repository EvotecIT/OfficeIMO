using System.Buffers;
using System.Threading;

namespace OfficeIMO.Excel.Xlsb.Package {
    /// <summary>
    /// Seekable read-only package-part stream backed by a buffer returned to the
    /// shared pool with the worksheet reader lifetime.
    /// </summary>
    internal sealed class XlsbPooledPartStream : MemoryStream {
        private byte[]? _buffer;
        private readonly int _length;
        private IDisposable? _secondaryOwner;

        internal XlsbPooledPartStream(byte[] buffer, int length, IDisposable? secondaryOwner = null)
            : base(buffer, 0, length, writable: false, publiclyVisible: false) {
            _buffer = buffer ?? throw new ArgumentNullException(nameof(buffer));
            _length = length;
            _secondaryOwner = secondaryOwner;
        }

        internal byte[] Buffer => _buffer
            ?? throw new ObjectDisposedException(nameof(XlsbPooledPartStream));

        internal int DataLength => _length;

        protected override void Dispose(bool disposing) {
            base.Dispose(disposing);
            byte[]? buffer = Interlocked.Exchange(ref _buffer, null);
            if (buffer != null) {
                Array.Clear(buffer, 0, _length);
                ArrayPool<byte>.Shared.Return(buffer);
            }
            Interlocked.Exchange(ref _secondaryOwner, null)?.Dispose();
        }
    }
}
