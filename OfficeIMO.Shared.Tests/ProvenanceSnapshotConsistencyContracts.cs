using System.Threading;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class ProvenanceSnapshotConsistencyContracts {
    [Fact]
    public void SnapshotCopyRejectsEqualLengthSourceMutationBetweenReadPasses() {
        byte[] first = Enumerable.Repeat((byte)0x11, 4096).ToArray();
        byte[] second = Enumerable.Repeat((byte)0x22, first.Length).ToArray();
        using var source = new MutatingOnRewindStream(first, second);
        using var destination = new MemoryStream();

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceFileSnapshot.CopyStableSource(
                source,
                destination,
                maximumBytes: first.Length,
                limitMessage: "limit",
                CancellationToken.None,
                out _));

        Assert.Contains("changed content", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    private sealed class MutatingOnRewindStream : Stream {
        private readonly byte[] _first;
        private readonly byte[] _second;
        private long _position;
        private bool _hasRead;
        private bool _useSecond;

        internal MutatingOnRewindStream(byte[] first, byte[] second) {
            _first = first;
            _second = second;
        }

        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => _first.LongLength;
        public override long Position {
            get => _position;
            set {
                if (value < 0 || value > Length) throw new ArgumentOutOfRangeException(nameof(value));
                if (value == 0 && _hasRead) _useSecond = true;
                _position = value;
            }
        }

        public override int Read(byte[] buffer, int offset, int count) {
            byte[] active = _useSecond ? _second : _first;
            int available = checked((int)Math.Min(count, Length - _position));
            if (available == 0) return 0;
            Buffer.BlockCopy(active, checked((int)_position), buffer, offset, available);
            _position += available;
            _hasRead = true;
            return available;
        }

        public override long Seek(long offset, SeekOrigin origin) {
            long target = origin switch {
                SeekOrigin.Begin => offset,
                SeekOrigin.Current => _position + offset,
                SeekOrigin.End => Length + offset,
                _ => throw new ArgumentOutOfRangeException(nameof(origin))
            };
            Position = target;
            return _position;
        }

        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
