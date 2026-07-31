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

        internal XlsbPooledPartStream(byte[] buffer, int length)
            : base(buffer, 0, length, writable: false, publiclyVisible: false) {
            _buffer = buffer ?? throw new ArgumentNullException(nameof(buffer));
            _length = length;
        }

        protected override void Dispose(bool disposing) {
            base.Dispose(disposing);
            byte[]? buffer = Interlocked.Exchange(ref _buffer, null);
            if (buffer != null) {
                Array.Clear(buffer, 0, _length);
                ArrayPool<byte>.Shared.Return(buffer);
            }
        }
    }
}
