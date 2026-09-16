using OfficeIMO.ContentSafety;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ContentSafetyInputGuardTests {
    [Fact]
    public void DefaultLiteralStillBindsToTheExistingZipInspectionParameter() {
        string path = Path.GetTempFileName();
        try {
            File.WriteAllBytes(path, new byte[] { 1, 2, 3 });

            byte[] bytes = OfficeContentSafetyInputGuard.ReadAllBytes(
                path,
                new OfficeContentSafetyOptions(),
                default);

            Assert.Equal(new byte[] { 1, 2, 3 }, bytes);
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void BoundedReadObservesCancellationBetweenInputChunks() {
        using var cancellation = new CancellationTokenSource();
        using var stream = new CancelAfterFirstReadStream(new byte[160_000], cancellation);

        Assert.Throws<OperationCanceledException>(() =>
            OfficeContentSafetyInputGuard.ReadBounded(stream, 200_000, cancellation.Token));

        Assert.Equal(1, stream.ReadCount);
    }

    private sealed class CancelAfterFirstReadStream : MemoryStream {
        private readonly CancellationTokenSource _cancellation;

        internal CancelAfterFirstReadStream(byte[] buffer, CancellationTokenSource cancellation)
            : base(buffer, writable: false) {
            _cancellation = cancellation;
        }

        internal int ReadCount { get; private set; }

        public override int Read(byte[] buffer, int offset, int count) {
            int read = base.Read(buffer, offset, count);
            ReadCount++;
            _cancellation.Cancel();
            return read;
        }
    }
}