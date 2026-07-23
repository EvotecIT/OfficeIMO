using System.Runtime.InteropServices;

namespace OfficeIMO.Email;

internal static class EmailTemporaryStorage {
    internal static void CreatePrivateDirectory(string path) {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) {
            Directory.CreateDirectory(path);
            return;
        }

        if (CreateDirectory(path, 0x1C0) == 0) return;
        int error = Marshal.GetLastWin32Error();
        throw new IOException($"Unable to create a private email temporary directory (error {error}).");
    }

    [DllImport("libc", EntryPoint = "mkdir", SetLastError = true, CharSet = CharSet.Ansi)]
    private static extern int CreateDirectory(string path, uint mode);
}
