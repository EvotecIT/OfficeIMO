using OfficeIMO.Reader;
using System.Threading;
using Xunit;

namespace OfficeIMO.Reader.Tests;

public sealed class ReaderInputLimitTests {
    [Fact]
    public void LargeSnapshotBudgetsUseDeleteOnCloseTemporaryStorage() {
        byte[] payload = System.Text.Encoding.UTF8.GetBytes(
            "bounded temporary snapshot");
        using var source = new MemoryStream(payload, writable: false);
        Stream prepared = ReaderInputLimits.EnsureSeekableReadStream(
            source,
            maxInputBytes: 64L * 1024 * 1024 + 1,
            CancellationToken.None,
            out bool ownsStream);
        using (prepared) {
            Assert.True(ownsStream);
            Assert.IsAssignableFrom<FileStream>(prepared);
            Assert.True(ReaderInputLimits.IsSnapshotStream(prepared));
            using var copy = new MemoryStream();
            prepared.CopyTo(copy);
            Assert.Equal(payload, copy.ToArray());

            Stream reused = ReaderInputLimits.EnsureSeekableReadStream(
                prepared,
                maxInputBytes: 64L * 1024 * 1024 + 1,
                CancellationToken.None,
                out bool ownsReused);
            Assert.Same(prepared, reused);
            Assert.False(ownsReused);
            Assert.Equal(0, reused.Position);
        }
    }
}
