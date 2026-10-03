using System.Runtime.InteropServices;

namespace OfficeIMO.Core.Internal {
    /// <summary>Unix file opening with the platform's variadic mode argument convention.</summary>
    internal static class OfficeUnixFile {
        internal static int Open(string path, int flags, uint mode) => IsMacArm64
            ? MacArm64Open(path, flags, 0, 0, 0, 0, 0, 0, mode)
            : UnixOpen(path, flags, mode);

        internal static int OpenAt(int directory, string path, int flags, uint mode) => IsMacArm64
            ? MacArm64OpenAt(directory, path, flags, 0, 0, 0, 0, 0, mode)
            : UnixOpenAt(directory, path, flags, mode);

        private static bool IsMacArm64 => RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
            && RuntimeInformation.ProcessArchitecture == Architecture.Arm64;

        // Darwin ARM64 passes variadic arguments on the stack. Occupy the remaining
        // registers so mode is the first stack argument. Fixing mode only after
        // creation would leave a window with unintended access permissions.
        // https://developer.apple.com/documentation/apple-silicon/addressing-architectural-differences-in-your-macos-code
        [DllImport("libc", EntryPoint = "open", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern int MacArm64Open(string path, int flags,
            nint padding2, nint padding3, nint padding4, nint padding5, nint padding6, nint padding7, uint mode);

        [DllImport("libc", EntryPoint = "openat", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern int MacArm64OpenAt(int directory, string path, int flags,
            nint padding3, nint padding4, nint padding5, nint padding6, nint padding7, uint mode);

        [DllImport("libc", EntryPoint = "open", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern int UnixOpen(string path, int flags, uint mode);

        [DllImport("libc", EntryPoint = "openat", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern int UnixOpenAt(int directory, string path, int flags, uint mode);
    }
}
