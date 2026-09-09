using System.Buffers;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Owns a small decompressed package part until its XML reader is disposed.
    /// The logical length excludes unused pool capacity and the bytes are cleared
    /// before reuse, including when parsing fails.
    /// </summary>
    internal sealed class OpenXmlPooledPartStream : MemoryStream {
        private byte[]? _buffer;

        internal OpenXmlPooledPartStream(byte[] buffer, int length)
            : base(buffer, 0, length, writable: false, publiclyVisible: false) {
            _buffer = buffer;
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
                if (buffer != null) ArrayPool<byte>.Shared.Return(buffer, clearArray: true);
            }
        }
    }
}
