namespace OfficeIMO.OneNote.Tests;

public sealed class RevisionStoreAssetSafetyTests {
    private static readonly Guid FileDataHeader = new Guid("BDE316E7-2665-4511-A4C4-8D4D0B7A9EAC");
    private static readonly Guid FileDataFooter = new Guid("71FBA722-0F79-4A0B-BB13-899256426B24");

    [Fact]
    public void ReadsOnlyHeaderPayloadAndFooterFromPaddedFileDataFrame() {
        var referenceId = Guid.NewGuid();
        byte[] payload = { 1, 2, 3 };
        const ulong frameOffset = 128;
        const ulong frameLength = 1024 * 1024;
        var reference = new OneNoteFileNodeChunkReference(frameOffset, frameLength, false, 8);
        var nodeData = new byte[24];
        referenceId.ToByteArray().CopyTo(nodeData, reference.EncodedLength);
        var node = new OneNoteFileNode(
            (ushort)OneNoteFileNodeId.FileDataStoreObjectReference,
            nodeData.Length + 4,
            0,
            0,
            OneNoteFileNodeBaseType.DataReference,
            32,
            reference,
            nodeData);
        var root = new OneNoteFileNodeList(1, Array.Empty<OneNoteFileNodeListFragment>(), new[] { node });
        using var stream = new SparseFileDataFrameStream(frameOffset, frameLength, payload);
        var options = new OneNoteReaderOptions {
            MaxInputBytes = (long)(frameOffset + frameLength),
            MaxAssetBytes = 16,
            MaxTotalAssetBytes = 16
        };

        OneNoteRevisionStoreObjectReadResult result = OneNoteRevisionStoreObjectReader.Read(
            stream,
            root,
            (ulong)stream.Length,
            options);

        OneNoteFileDataStoreObject fileData = Assert.Single(result.FileDataObjects);
        Assert.Equal(referenceId, fileData.Id);
        Assert.Equal(payload, fileData.Payload.ToArray(16));
        Assert.True(stream.MaxReadRequest <= 36, "The reader requested " + stream.MaxReadRequest + " bytes at once.");
        Assert.Equal(55, stream.TotalBytesRead);
    }

    private sealed class SparseFileDataFrameStream : Stream {
        private readonly ulong _frameOffset;
        private readonly ulong _frameLength;
        private readonly byte[] _header;
        private readonly byte[] _payload;
        private readonly byte[] _footer;
        private long _position;

        internal SparseFileDataFrameStream(ulong frameOffset, ulong frameLength, byte[] payload) {
            _frameOffset = frameOffset;
            _frameLength = frameLength;
            _payload = payload;
            _header = new byte[36];
            FileDataHeader.ToByteArray().CopyTo(_header, 0);
            BitConverter.GetBytes((ulong)payload.Length).CopyTo(_header, 16);
            _footer = FileDataFooter.ToByteArray();
        }

        internal int MaxReadRequest { get; private set; }

        internal long TotalBytesRead { get; private set; }

        public override bool CanRead => true;

        public override bool CanSeek => true;

        public override bool CanWrite => false;

        public override long Length => checked((long)(_frameOffset + _frameLength));

        public override long Position {
            get => _position;
            set => _position = value >= 0 && value <= Length ? value : throw new ArgumentOutOfRangeException(nameof(value));
        }

        public override int Read(byte[] buffer, int offset, int count) {
            if (buffer == null) throw new ArgumentNullException(nameof(buffer));
            if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException();
            MaxReadRequest = Math.Max(MaxReadRequest, count);
            if (count > 64) throw new InvalidOperationException("The asset reader tried to materialize the padded frame.");
            if (_position >= Length) return 0;

            int read = (int)Math.Min(count, Length - _position);
            Array.Clear(buffer, offset, read);
            CopyIntersection(_header, checked((long)_frameOffset), buffer, offset, _position, read);
            CopyIntersection(_payload, checked((long)_frameOffset + 36), buffer, offset, _position, read);
            CopyIntersection(_footer, checked((long)(_frameOffset + _frameLength - 16)), buffer, offset, _position, read);
            _position += read;
            TotalBytesRead += read;
            return read;
        }

        private static void CopyIntersection(
            byte[] source,
            long sourceOffset,
            byte[] destination,
            int destinationOffset,
            long readOffset,
            int readLength) {
            long start = Math.Max(sourceOffset, readOffset);
            long end = Math.Min(sourceOffset + source.Length, readOffset + readLength);
            if (start >= end) return;
            Buffer.BlockCopy(
                source,
                checked((int)(start - sourceOffset)),
                destination,
                checked(destinationOffset + (int)(start - readOffset)),
                checked((int)(end - start)));
        }

        public override long Seek(long offset, SeekOrigin origin) {
            long next = origin switch {
                SeekOrigin.Begin => offset,
                SeekOrigin.Current => checked(_position + offset),
                SeekOrigin.End => checked(Length + offset),
                _ => throw new ArgumentOutOfRangeException(nameof(origin))
            };
            Position = next;
            return _position;
        }

        public override void Flush() { }

        public override void SetLength(long value) => throw new NotSupportedException();

        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
