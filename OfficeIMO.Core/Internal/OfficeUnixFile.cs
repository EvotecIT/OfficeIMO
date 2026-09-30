using System;
using System.Runtime.InteropServices;

namespace OfficeIMO.Core.Internal {
    /// <summary>Unix file creation calls with platform-correct variadic mode argument placement.</summary>
    internal static class OfficeUnixFile {
        private static bool UsesAppleArm64Abi => RuntimeInformation.IsOSPlatform(OSPlatform.OSX)
            && RuntimeInformation.ProcessArchitecture == Architecture.Arm64;

        /// <summary>Opens a Unix path with native flags and an explicit creation mode.</summary>
        /// <remarks>The caller owns a successful descriptor and reads the last native error on failure.</remarks>
        internal static int OpenWithMode(string path, int flags, uint mode) => UsesAppleArm64Abi
            ? AppleArm64Open(path, flags, UIntPtr.Zero, UIntPtr.Zero, UIntPtr.Zero,
                UIntPtr.Zero, UIntPtr.Zero, UIntPtr.Zero, (UIntPtr)mode)
            : UnixOpen(path, flags, mode);

        /// <summary>Opens a Unix path relative to an already-open directory with an explicit creation mode.</summary>
        /// <remarks>The caller owns a successful descriptor and reads the last native error on failure.</remarks>
        internal static int OpenAtWithMode(int directory, string path, int flags, uint mode) => UsesAppleArm64Abi
            ? AppleArm64OpenAt(directory, path, flags, UIntPtr.Zero, UIntPtr.Zero,
                UIntPtr.Zero, UIntPtr.Zero, UIntPtr.Zero, (UIntPtr)mode)
            : UnixOpenAt(directory, path, flags, mode);

        // Apple ARM64 passes variadic arguments on the stack. Fixed P/Invoke signatures otherwise
        // put mode in x2/x3, where libc does not read it. Fill the eight integer argument registers
        // so the ninth argument occupies the first stack slot required by open/openat.
        [DllImport("libc", EntryPoint = "open", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern int AppleArm64Open(string path, int flags, UIntPtr register2,
            UIntPtr register3, UIntPtr register4, UIntPtr register5, UIntPtr register6,
            UIntPtr register7, UIntPtr mode);

        [DllImport("libc", EntryPoint = "openat", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern int AppleArm64OpenAt(int directory, string path, int flags,
            UIntPtr register3, UIntPtr register4, UIntPtr register5, UIntPtr register6,
            UIntPtr register7, UIntPtr mode);

        [DllImport("libc", EntryPoint = "open", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern int UnixOpen(string path, int flags, uint mode);

        [DllImport("libc", EntryPoint = "openat", SetLastError = true, CharSet = CharSet.Ansi)]
        private static extern int UnixOpenAt(int directory, string path, int flags, uint mode);
    }
}
