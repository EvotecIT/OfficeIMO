namespace OfficeIMO.Core.Internal {
    /// <summary>
    /// Presents loaded package bytes as the expandable stream an editable package needs, copying them only when the
    /// package first writes. Opening and reading a document therefore holds one copy of its bytes, and the source
    /// array is never modified, so callers that keep it as an original-package baseline still see the loaded bytes.
    /// </summary>
    internal sealed class OfficeCopyOnWriteStream : Stream {
        private readonly byte[] _source;
        private MemoryStream _current;
        private bool _copied;
        private bool _abandoned;
        private bool _disposed;
        private readonly bool _survivesDispose;

        /// <param name="source">The loaded bytes. They are never written.</param>
        /// <param name="survivesDispose">
        /// Keeps the content readable and writable after <see cref="Stream.Dispose()"/>, for a package stream its owner
        /// still reads after the package closes it; <see cref="Release"/> then disposes it.
        /// </param>
        internal OfficeCopyOnWriteStream(byte[] source, bool survivesDispose = false) {
            _source = source ?? throw new ArgumentNullException(nameof(source));
            _current = new MemoryStream(source, 0, source.Length, writable: false, publiclyVisible: true);
            _survivesDispose = survivesDispose;
        }

        /// <summary>Disposes the stream, including one created to survive <see cref="Stream.Dispose()"/>.</summary>
        internal void Release() {
            if (_disposed) return;
            _disposed = true;
            _current.Dispose();
        }

        public override bool CanRead => !_disposed;
        public override bool CanSeek => !_disposed;
        public override bool CanWrite => !_disposed;
        public override long Length => _current.Length;
        public override long Position { get => _current.Position; set => _current.Position = value; }

        /// <summary>Discards later writes. The closing package rewrites itself here, but nothing reads this stream again.</summary>
        internal void Abandon() => _abandoned = true;

        /// <summary>Returns the current package bytes as a new array.</summary>
        internal byte[] ToArray() => _current.ToArray();

        /// <summary>
        /// Exposes the current bytes: the loaded array until the first write, then the stream's own buffer. The
        /// loaded array must not be written through the returned segment.
        /// </summary>
        internal bool TryGetBuffer(out ArraySegment<byte> buffer) => _current.TryGetBuffer(out buffer);

        /// <summary>The expandable stream holding the bytes once they have been copied; null before the first write.</summary>
        internal MemoryStream? CopiedStream => _copied ? _current : null;

        /// <summary>
        /// Replaces the loaded bytes with an empty expandable stream for a writer that rewrites the whole content
        /// while still reading the loaded array, which stays unchanged. Valid only before the first write.
        /// </summary>
        internal MemoryStream BeginRewrite(int capacity) {
            if (_copied) throw new InvalidOperationException("The stream already holds its own copy.");
            _current = new MemoryStream(capacity);
            _copied = true;
            return _current;
        }

        /// <summary>Returns to the loaded bytes after a rewrite started by <see cref="BeginRewrite"/> failed.</summary>
        internal void CancelRewrite() {
            _current = new MemoryStream(_source, 0, _source.Length, writable: false, publiclyVisible: true);
            _copied = false;
        }

        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => _current.Read(buffer, offset, count);
        public override int ReadByte() => _current.ReadByte();
        public override long Seek(long offset, SeekOrigin origin) => _current.Seek(offset, origin);

        public override void SetLength(long value) {
            if (_abandoned) return;
            EnsureCopied(value);
            _current.SetLength(value);
        }

        public override void Write(byte[] buffer, int offset, int count) {
            if (_abandoned) { _current.Position += count; return; }
            EnsureCopied();
            _current.Write(buffer, offset, count);
        }

        public override void WriteByte(byte value) {
            if (_abandoned) { _current.Position++; return; }
            EnsureCopied();
            _current.WriteByte(value);
        }

#if NET5_0_OR_GREATER
        public override int Read(Span<byte> buffer) => _current.Read(buffer);

        public override void Write(ReadOnlySpan<byte> buffer) {
            if (_abandoned) { _current.Position += buffer.Length; return; }
            EnsureCopied();
            _current.Write(buffer);
        }
#endif

        protected override void Dispose(bool disposing) {
            if (disposing && !_survivesDispose) {
                // Later reads fail as they would on a disposed MemoryStream; the loaded array itself is untouched.
                Release();
            }
            base.Dispose(disposing);
        }

        private void EnsureCopied(long? newLength = null) {
            if (_copied) return;
            long position = _current.Position;
            // A truncating writer keeps only the bytes below the new length.
            int keep = newLength.HasValue ? (int)Math.Min(Math.Max(newLength.Value, 0), _source.Length) : _source.Length;
            var copy = new MemoryStream(keep);
            copy.Write(_source, 0, keep);
            copy.Position = position;
            _current = copy;
            _copied = true;
        }
    }
}
