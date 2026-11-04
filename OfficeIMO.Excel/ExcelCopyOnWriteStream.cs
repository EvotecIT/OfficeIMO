namespace OfficeIMO.Excel {
    /// <summary>
    /// Presents loaded package bytes as the expandable stream an editable package needs, copying them only when the
    /// package first writes. Opening and reading a workbook therefore holds one copy of its bytes.
    /// </summary>
    internal sealed class ExcelCopyOnWriteStream : Stream {
        private readonly byte[] _source;
        private MemoryStream _current;
        private bool _copied;
        private bool _abandoned;

        internal ExcelCopyOnWriteStream(byte[] source) {
            _source = source;
            _current = new MemoryStream(source, writable: false);
        }

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => true;
        public override long Length => _current.Length;
        public override long Position { get => _current.Position; set => _current.Position = value; }

        /// <summary>Discards later writes. The closing package rewrites itself here, but nothing reads this stream again.</summary>
        internal void Abandon() => _abandoned = true;

        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => _current.Read(buffer, offset, count);
        public override int ReadByte() => _current.ReadByte();
        public override long Seek(long offset, SeekOrigin origin) => _current.Seek(offset, origin);

        public override void SetLength(long value) {
            if (_abandoned) return;
            EnsureCopied();
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

        private void EnsureCopied() {
            if (_copied) return;
            long position = _current.Position;
            var copy = new MemoryStream(_source.Length);
            copy.Write(_source, 0, _source.Length);
            copy.Position = position;
            _current = copy;
            _copied = true;
        }
    }
}
