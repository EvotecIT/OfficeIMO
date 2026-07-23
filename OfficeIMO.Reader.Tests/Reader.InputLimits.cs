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
        string snapshotPath = Assert.IsAssignableFrom<FileStream>(prepared).Name;
        string snapshotDirectory = Path.GetDirectoryName(snapshotPath)!;
        using (prepared) {
            Assert.True(ownsStream);
            Assert.True(ReaderInputLimits.IsSnapshotStream(prepared));
            Assert.NotEqual(Path.GetTempPath().TrimEnd(Path.DirectorySeparatorChar), snapshotDirectory.TrimEnd(Path.DirectorySeparatorChar));
            Assert.Throws<IOException>(() => new FileStream(snapshotPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite));
#if NET6_0_OR_GREATER
            if (!OperatingSystem.IsWindows()) {
                Assert.Equal(
                    UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute,
                    File.GetUnixFileMode(snapshotDirectory));
            }
#endif
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

        Assert.False(Directory.Exists(snapshotDirectory));
    }
}
