namespace OfficeIMO.Archives.Tests;

public sealed class ArchiveHeaderTests {
    [Theory]
    [InlineData(1, false)]
    [InlineData(3, false)]
    [InlineData(1, true)]
    [InlineData(3, true)]
    public void FragmentedHeaderReadsDetectBothContainersAndRestorePosition(int chunk, bool compound) {
        byte[] bytes = compound ? new byte[] { 0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1 }
            : Encoding.ASCII.GetBytes("MSCFrest");
        using var source = new FragmentedStream(bytes, chunk);
        source.Position = 2;
        Assert.Equal(compound ? ArchiveFormat.Msi : ArchiveFormat.Cab, ArchiveExtractor.Detect(source, CancellationToken.None));
        Assert.Equal(2, source.Position);
        Assert.True(source.CanRead);
    }

    [Fact]
    public void TruncatedHeaderStopsAtEofAndRestoresPosition() {
        using var source = new FragmentedStream(new byte[] { 0xd0, 0xcf }, 1);
        source.Position = 1;
        Assert.Throws<NotSupportedException>(() => ArchiveExtractor.Detect(source, CancellationToken.None));
        Assert.Equal(1, source.Position);
        Assert.InRange(source.ReadCalls, 1, 3);
    }

    [Fact]
    public void CancellationBetweenHeaderFragmentsRestoresPosition() {
        using var cancellation = new CancellationTokenSource();
        using var source = new FragmentedStream(Encoding.ASCII.GetBytes("MSCFrest"), 1, cancellation.Cancel);
        source.Position = 2;
        Assert.Throws<OperationCanceledException>(() => ArchiveExtractor.Detect(source, cancellation.Token));
        Assert.Equal(2, source.Position);
    }

    private sealed class FragmentedStream : MemoryStream {
        private readonly int _chunk;
        private readonly Action? _afterRead;
        internal FragmentedStream(byte[] bytes, int chunk, Action? afterRead = null) : base(bytes, writable: false) {
            _chunk = chunk; _afterRead = afterRead;
        }
        internal int ReadCalls { get; private set; }
        public override int Read(byte[] buffer, int offset, int count) {
            ReadCalls++;
            int read = base.Read(buffer, offset, Math.Min(count, _chunk));
            _afterRead?.Invoke();
            return read;
        }
    }
}
