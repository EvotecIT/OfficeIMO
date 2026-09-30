using OfficeIMO.Core.Internal;
using System;
using System.IO;
using System.Threading.Tasks;
using Xunit;
using Microsoft.Win32.SafeHandles;

namespace OfficeIMO.Shared.Tests {
    public sealed class OfficeFileCommitPermissionsTests {
#if NET6_0_OR_GREATER
        [Theory]
        [InlineData(256U)] // 0400
        [InlineData(384U)] // 0600
        public void NativeUnixCreationUsesTheRequestedModeBeforeAnyPermissionRepair(uint mode) {
            if (OperatingSystem.IsWindows()) return;
            if (!OperatingSystem.IsMacOS() && !OperatingSystem.IsLinux()) return;
            string root = Path.Combine(Path.GetTempPath(), "officeimo-create-mode-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            string path = Path.Combine(root, "private.bin");
            try {
                int flags = 2 | (OperatingSystem.IsMacOS() ? 0x0200 | 0x0800 : 0x0040 | 0x0080);
                int descriptor = OfficeUnixFile.OpenWithMode(path, flags, mode);
                Assert.True(descriptor >= 0);
                using var handle = new SafeFileHandle(new IntPtr(descriptor), ownsHandle: true);
                using var stream = new FileStream(handle, FileAccess.ReadWrite);
                Assert.Equal((UnixFileMode)mode, File.GetUnixFileMode(path));
                stream.WriteByte(42);
                stream.Flush();
                Assert.Equal(new byte[] { 42 }, File.ReadAllBytes(path));
            } finally {
                Directory.Delete(root, recursive: true);
            }
        }

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
