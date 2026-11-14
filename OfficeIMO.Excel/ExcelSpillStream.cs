namespace OfficeIMO.Excel {
    /// <summary>
    /// A write-then-read spool that stays in memory up to a limit and moves to a temporary file, deleted when the
    /// stream closes, once it grows past it.
    /// </summary>
    internal sealed class ExcelSpillStream : Stream {
        private readonly int _memoryLimit;
        private readonly string _filePrefix;
        private MemoryStream? _memory = new MemoryStream();
        private FileStream? _file;

        internal ExcelSpillStream(int memoryLimit, string filePrefix) {
            _memoryLimit = memoryLimit;
            _filePrefix = filePrefix;
        }

        private Stream Inner => (Stream?)_file ?? _memory ?? throw new ObjectDisposedException(nameof(ExcelSpillStream));

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => true;
        public override long Length => Inner.Length;

        public override long Position {
            get => Inner.Position;
            set => Inner.Position = value;
        }

        public override void Write(byte[] buffer, int offset, int count) {
            if (_memory != null && _memory.Length + count > _memoryLimit) MoveToFile();
            Inner.Write(buffer, offset, count);
        }

        public override int Read(byte[] buffer, int offset, int count) => Inner.Read(buffer, offset, count);

        public override long Seek(long offset, SeekOrigin origin) => Inner.Seek(offset, origin);

        public override void SetLength(long value) => Inner.SetLength(value);

        public override void Flush() => Inner.Flush();

        private void MoveToFile() {
            string path = Path.Combine(Path.GetTempPath(), _filePrefix + Guid.NewGuid().ToString("N") + ".xml");
            var file = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None,
                64 * 1024, FileOptions.SequentialScan | FileOptions.DeleteOnClose);
            try {
                long position = _memory!.Position;
                _memory.Position = 0;
                _memory.CopyTo(file);
                file.Position = position;
            } catch {
                file.Dispose();
                throw;
            }

            _memory.Dispose();
            _memory = null;
            _file = file;
        }

        protected override void Dispose(bool disposing) {
            if (disposing) {
                _memory?.Dispose();
                _memory = null;
                _file?.Dispose();
                _file = null;
            }

            base.Dispose(disposing);
        }
    }
}
