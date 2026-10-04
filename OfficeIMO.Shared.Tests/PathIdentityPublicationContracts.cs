using OfficeIMO.Internal;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public class PathIdentityPublicationContracts {
    [Fact]
    public void PublicationEntryIdentitySurvivesReplacementAndHonorsParentAliases() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-entry-path-" + Guid.NewGuid().ToString("N"));
        string destination = Path.Combine(root, "pages");
        string recovery = Path.Combine(root, "recovery");
        Directory.CreateDirectory(root);
        try {
            string expected = OfficePathIdentity.NormalizeDirectoryEntry(destination);
            Directory.CreateDirectory(destination);
            Assert.Equal(expected, OfficePathIdentity.NormalizeDirectoryEntry(destination));
            Directory.Move(destination, recovery);
            Assert.Equal(expected, OfficePathIdentity.NormalizeDirectoryEntry(destination));
            Directory.CreateDirectory(destination);
            Assert.Equal(expected, OfficePathIdentity.NormalizeDirectoryEntry(destination));
            Assert.NotEqual(expected, OfficePathIdentity.NormalizeDirectoryEntry(recovery));
            Assert.Equal(expected, OfficePathIdentity.NormalizeDirectoryEntry(Path.Combine(root, ".", "pages")));
            if (OfficePathIdentity.IsCaseInsensitiveFileSystem(root)) {
                Assert.Equal(expected, OfficePathIdentity.NormalizeDirectoryEntry(Path.Combine(root, "PAGES")));
            } else {
                Assert.NotEqual(expected, OfficePathIdentity.NormalizeDirectoryEntry(Path.Combine(root, "PAGES")));
            }
#if NET8_0_OR_GREATER
            string alias = Path.Combine(root, "parent-alias");
            try { Directory.CreateSymbolicLink(alias, root); }
            catch (Exception exception) when (exception is UnauthorizedAccessException || exception is IOException) { return; }
            Assert.Equal(expected, OfficePathIdentity.NormalizeDirectoryEntry(Path.Combine(alias, "pages")));
            Directory.Delete(alias);
#endif
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task PhysicalResolutionAllowsAConcurrentDirectoryReplacement() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-moving-path-" + Guid.NewGuid().ToString("N"));
        string destination = Path.Combine(root, "pages");
        string recovery = Path.Combine(root, "recovery");
        Directory.CreateDirectory(destination);
        using var ready = new ManualResetEventSlim();
        using var stop = new CancellationTokenSource();
        Task mover = Task.Run(() => {
            ready.Set();
            while (!stop.IsCancellationRequested) {
                Directory.Move(destination, recovery);
                Directory.Move(recovery, destination);
            }
        });
        try {
            ready.Wait();
            for (int index = 0; index < 2_000; index++) {
                string physical = OfficePathIdentity.ResolvePhysicalPath(destination);
                Assert.True(Path.IsPathRooted(physical));
                Assert.Equal(root, Path.GetDirectoryName(physical));
            }
        } finally {
            stop.Cancel();
            await mover;
            Directory.Delete(root, recursive: true);
        }
    }
}
