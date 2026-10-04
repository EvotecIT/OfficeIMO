#if NET6_0_OR_GREATER
using Microsoft.Win32.SafeHandles;
using OfficeIMO.Core.Internal;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class OfficeUnixFileTests {
    [Theory]
    [InlineData(false, 384)]
    [InlineData(true, 384)]
    [InlineData(false, 256)]
    [InlineData(true, 256)]
    public void NativeCreationUsesRequestedOwnerPermissions(bool relative, int mode) {
        if (OperatingSystem.IsWindows()) return;
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.UnixMode", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "created.bin");
        int flags = OperatingSystem.IsMacOS() ? 1 | 0x200 | 0x800 : 1 | 0x40 | 0x80;
        try {
            using var parent = new SafeFileHandle((nint)OfficeUnixFile.OpenWithMode(directory, 0, 0), ownsHandle: true);
            int descriptor = relative
                ? OfficeUnixFile.OpenAtWithMode(parent.DangerousGetHandle().ToInt32(), "created.bin", flags, (uint)mode)
                : OfficeUnixFile.OpenWithMode(path, flags, (uint)mode);
            Assert.True(descriptor >= 0, "Exclusive native file creation failed.");
            using (var handle = new SafeFileHandle((nint)descriptor, ownsHandle: true))
            using (var stream = new FileStream(handle, FileAccess.Write)) stream.WriteByte(81);
            Assert.Equal((UnixFileMode)mode, File.GetUnixFileMode(path));
            Assert.Equal(new byte[] { 81 }, File.ReadAllBytes(path));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }
}
#endif
