using System.Buffers;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Owns a decompressed package part until its reader or shared-string index is disposed.
    /// The logical length excludes unused pool capacity and the bytes are cleared
    /// before reuse, including when parsing fails.
    /// </summary>
    internal sealed class OpenXmlPooledPartStream : MemoryStream {
        private byte[]? _buffer;
        private readonly bool _usesPartBufferPool;

        internal OpenXmlPooledPartStream(byte[] buffer, int length, bool usesPartBufferPool = false)
            : base(buffer, 0, length, writable: false, publiclyVisible: false) {
            _buffer = buffer;
            _usesPartBufferPool = usesPartBufferPool;
        }

        internal byte[] BorrowBuffer(out int length) {
            byte[]? buffer = _buffer;
            if (buffer == null) throw new ObjectDisposedException(nameof(OpenXmlPooledPartStream));
            length = checked((int)Length);
            return buffer;
        }

        public override byte[] ToArray() {
            if (_buffer == null) throw new ObjectDisposedException(nameof(OpenXmlPooledPartStream));
            return base.ToArray();
        }

        protected override void Dispose(bool disposing) {
            byte[]? buffer = _buffer;
            _buffer = null;
            try {
                base.Dispose(disposing);
            } finally {
                if (buffer != null) {
                    if (_usesPartBufferPool) OpenXmlPartBufferPool.Return(buffer);
                    else ArrayPool<byte>.Shared.Return(buffer, clearArray: true);
                }
            }
        }
    }
}
