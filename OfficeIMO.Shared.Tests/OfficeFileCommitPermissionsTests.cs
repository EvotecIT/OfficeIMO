using OfficeIMO.Core.Internal;
using System;
using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Shared.Tests {
    public sealed class OfficeFileCommitPermissionsTests {
#if NET6_0_OR_GREATER
        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public async Task NativeOwnerOnlyHandleSupportsAsyncIoAndRequestedLifetime(bool deleteOnClose) {
            if (OperatingSystem.IsWindows()) return;
            string root = Path.Combine(Path.GetTempPath(), "officeimo-native-async-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            string path = Path.Combine(root, "private.bin");
            try {
                FileOptions options = FileOptions.Asynchronous | (deleteOnClose ? FileOptions.DeleteOnClose : FileOptions.None);
                using (FileStream stream = OfficeTemporaryFile.CreateUnixOwnerOnly(path, 4096, options)) {
                    await stream.WriteAsync(new byte[] { 1, 2, 3 });
                    await stream.FlushAsync();
                    stream.Position = 0;
                    var actual = new byte[3];
                    await stream.ReadExactlyAsync(actual);
                    Assert.Equal(new byte[] { 1, 2, 3 }, actual);
                    Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite, File.GetUnixFileMode(path));
                }
                Assert.Equal(!deleteOnClose, File.Exists(path));
            } finally {
                Directory.Delete(root, recursive: true);
            }
        }
#endif

        [Theory]
        [InlineData(false)]
        [InlineData(true)]
        public void OwnerOnlyPublicationCreatesOrReplacesCompleteBytes(bool existing) {
            string root = Path.Combine(Path.GetTempPath(), "officeimo-private-commit-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            string target = Path.Combine(root, "private.bin");
            string staging = Path.Combine(root, "staging.bin");
            try {
                if (existing) File.WriteAllBytes(target, new byte[] { 1 });
                File.WriteAllBytes(staging, new byte[] { 2, 3 });
#if NET6_0_OR_GREATER
                if (!OperatingSystem.IsWindows()) {
                    File.SetUnixFileMode(staging, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.GroupRead | UnixFileMode.OtherRead);
                    if (existing) File.SetUnixFileMode(target, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.GroupRead | UnixFileMode.OtherRead);
                }
#endif
                OfficeFileCommit.CommitTemporaryFileAtomically(staging, target,
                    OfficeFileCommit.ConflictPolicy.Replace, OfficeFileCommit.UnixFileAccessPolicy.OwnerOnly);
                Assert.Equal(new byte[] { 2, 3 }, File.ReadAllBytes(target));
                Assert.False(File.Exists(staging));
#if NET6_0_OR_GREATER
                if (!OperatingSystem.IsWindows()) Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite, File.GetUnixFileMode(target));
#endif
            } finally {
                Directory.Delete(root, recursive: true);
            }
        }

        [Fact]
        public void OwnerOnlyConflictPreservesDestinationAndRestrictsStagingBeforePublication() {
            string root = Path.Combine(Path.GetTempPath(), "officeimo-private-collision-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            string target = Path.Combine(root, "private.bin");
            string staging = Path.Combine(root, "staging.bin");
            try {
                File.WriteAllBytes(target, new byte[] { 1 });
                File.WriteAllBytes(staging, new byte[] { 2 });
#if NET6_0_OR_GREATER
                UnixFileMode originalMode = 0;
                if (!OperatingSystem.IsWindows()) {
                    File.SetUnixFileMode(target, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.GroupRead);
                    File.SetUnixFileMode(staging, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.GroupRead);
                    originalMode = File.GetUnixFileMode(target);
                }
#endif
                Assert.Throws<IOException>(() => OfficeFileCommit.CommitTemporaryFileAtomically(staging, target,
                    OfficeFileCommit.ConflictPolicy.FailIfExists, OfficeFileCommit.UnixFileAccessPolicy.OwnerOnly));
                Assert.Equal(new byte[] { 1 }, File.ReadAllBytes(target));
                Assert.Equal(new byte[] { 2 }, File.ReadAllBytes(staging));
#if NET6_0_OR_GREATER
                if (!OperatingSystem.IsWindows()) {
                    Assert.Equal(originalMode, File.GetUnixFileMode(target));
                    Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite, File.GetUnixFileMode(staging));
                }
#endif
            } finally {
                Directory.Delete(root, recursive: true);
            }
        }
    }
}
