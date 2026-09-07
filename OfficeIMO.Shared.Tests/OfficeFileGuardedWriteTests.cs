using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class OfficeFileGuardedWriteTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task GuardedWritePreservesAChangeMadeWhileTheOutputIsStaged(bool replace, bool identicalContent) {
        string root = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "officeimo-guarded-write-" + Guid.NewGuid().ToString("N"))).FullName;
        string path = Path.Combine(root, "document.bin");
        try {
            File.WriteAllText(path, "opened");
            string expectedIdentity = OfficePathIdentity.GetPhysicalIdentityKey(path);
            string? replacementIdentity = null;
            await Assert.ThrowsAsync<IOException>(() => OfficeFileCommit.WriteIfUnchangedAsync(path, async (stream, token) => {
                byte[] bytes = { 1, 2, 3 };
                await stream.WriteAsync(bytes, 0, bytes.Length, token);
                if (replace) {
                    string replacement = Path.Combine(root, "replacement.bin");
                    File.WriteAllText(replacement, identicalContent ? "opened" : "external");
                    replacementIdentity = OfficePathIdentity.GetPhysicalIdentityKey(replacement);
                    File.Delete(path);
                    File.Move(replacement, path);
                } else File.WriteAllText(path, "external");
            }, candidate => {
                using var input = new FileStream(candidate, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete);
                if (OfficePathIdentity.GetPhysicalIdentityKey(candidate, input.SafeFileHandle) != expectedIdentity) return false;
                using var reader = new StreamReader(input);
                return reader.ReadToEnd() == "opened";
            }));
            Assert.Equal(identicalContent ? "opened" : "external", File.ReadAllText(path));
            if (replace) {
                Assert.NotEqual(expectedIdentity, replacementIdentity);
                Assert.Equal(replacementIdentity, OfficePathIdentity.GetPhysicalIdentityKey(path));
            }
            Assert.Equal(new[] { path }, Directory.GetFileSystemEntries(root));
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task GuardedWriteRejectsAnOldSnapshotBeforeProducingOutputAndCleansCancellation() {
        string root = Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "officeimo-guarded-write-" + Guid.NewGuid().ToString("N"))).FullName;
        string path = Path.Combine(root, "document.bin");
        try {
            File.WriteAllText(path, "current");
            bool produced = false;
            await Assert.ThrowsAsync<IOException>(() => OfficeFileCommit.WriteIfUnchangedAsync(path,
                (_, _) => { produced = true; return Task.CompletedTask; }, _ => false));
            Assert.False(produced);
            using var cancellation = new CancellationTokenSource();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => OfficeFileCommit.WriteIfUnchangedAsync(path,
                (stream, _) => { stream.WriteByte(1); cancellation.Cancel(); return Task.CompletedTask; },
                candidate => File.ReadAllText(candidate) == "current", cancellation.Token));
            Assert.Equal("current", File.ReadAllText(path));
            Assert.Equal(new[] { path }, Directory.GetFileSystemEntries(root));
            await OfficeFileCommit.WriteIfUnchangedAsync(path,
                (stream, _) => { stream.WriteByte(5); return Task.CompletedTask; }, candidate => File.ReadAllText(candidate) == "current");
            Assert.Equal(new byte[] { 5 }, File.ReadAllBytes(path));
        } finally { Directory.Delete(root, true); }
    }
}
