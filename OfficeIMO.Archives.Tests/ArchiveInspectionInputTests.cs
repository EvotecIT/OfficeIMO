namespace OfficeIMO.Archives.Tests {

    public sealed class ArchiveInspectionInputTests {
        [Fact]
        public void NonseekableRefusalReportsTheCompleteInputMaximum() {
            using var source = new ForwardOnlyInput(new byte[65537]);
            var error = Assert.Throws<OfficeInspectionResourceLimitException>(() => ArchiveInspector.Open(source,
                new() { MaxInputBytes = 65536 }));
            Assert.Equal("MaxInputBytes", error.LimitName);
            Assert.Equal(65536, error.Maximum);
            Assert.True(source.CanRead);
        }

        [Fact]
        public void LargeZipInspectionSnapshotsBeforeTheCallerClosesOrChangesItsSource() {
            using var source = new PaddedZipInput(64L * 1024 * 1024 + 4096);
            source.Position = 7;
            using var inspection = ArchiveInspector.Open(source);
            Assert.Equal(7, source.Position);
            Assert.Equal(source.Length, inspection.InputLength);
            source.Dispose();
            using var payload = inspection.OpenRead(0);
            Assert.Equal("complete", ArchiveZipInspectionTests.Text(payload));
        }

        [Fact]
        public void OwnedZipSnapshotCanExceedArraySizeAndReturnsDetachedMembers() {
            var source = new PaddedZipInput((long)int.MaxValue + 4096);
            var inspection = ArchiveInspector.OpenZipSnapshot(source,
                new() { MaxInputBytes = source.Length }, CancellationToken.None);
            Assert.Single(ArchiveReader.TraverseZip(inspection).Entries);
            using var payload = inspection.OpenRead(0);
            inspection.Dispose();
            Assert.False(source.CanRead);
            Assert.Equal("complete", ArchiveZipInspectionTests.Text(payload));
        }

        [Fact]
        public void FailedZipInventoryClosesItsOwnedSnapshot() {
            var source = new MemoryStream(new byte[22]);
            Assert.Throws<InvalidDataException>(() => ArchiveInspector.OpenZipSnapshot(source, new(), CancellationToken.None));
            Assert.False(source.CanRead);
        }

        private sealed class ForwardOnlyInput(byte[] bytes) : MemoryStream(bytes) {
            public override bool CanSeek => false;
        }

        // A sparse, immutable ZIP source. Only its prefix and relocated central directory exist in memory.
        // The >2 GiB case indexes this already-owned source and does not copy its padding to disk.
        private sealed class PaddedZipInput : Stream {
            private readonly byte[] _prefix, _tail;
            private readonly long _tailOffset, _length;
            private long _position;
            private bool _disposed;
            internal PaddedZipInput(long length) {
                byte[] zip = ArchiveZipInspectionTests.Zip(("member.txt", "complete"));
                int central = ArchiveZipInspectionTests.Central(zip);
                _prefix = zip.Take(central).ToArray(); _tail = zip.Skip(central).ToArray();
                _length = length; _tailOffset = length - _tail.Length;
                BitConverter.GetBytes(checked((uint)_tailOffset)).CopyTo(_tail, _tail.Length - 6);
            }
            public override bool CanRead => !_disposed;
            public override bool CanSeek => !_disposed;
            public override bool CanWrite => false;
            public override long Length => _length;
            public override long Position { get => _position; set => _position = value; }
            public override int Read(byte[] buffer, int offset, int count) {
                if (_disposed) throw new ObjectDisposedException(nameof(PaddedZipInput));
                if (_position >= _length) return 0;
                int read = (int)Math.Min(count, _length - _position);
                Array.Clear(buffer, offset, read);
                for (int i = 0; i < read; i++) {
                    long at = _position + i;
                    if (at < _prefix.Length) buffer[offset + i] = _prefix[(int)at];
                    else if (at >= _tailOffset) buffer[offset + i] = _tail[(int)(at - _tailOffset)];
                }
                _position += read; return read;
            }
            public override long Seek(long offset, SeekOrigin origin) => Position = origin == SeekOrigin.Begin ? offset
                : origin == SeekOrigin.Current ? _position + offset : _length + offset;
            public override void Flush() { }
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
            protected override void Dispose(bool disposing) { _disposed = true; base.Dispose(disposing); }
        }
    }
}
