using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Win32.SafeHandles;

namespace OfficeIMO.Core.Internal {
    internal static partial class OfficeTemporaryFile {
        // OWNER RIGHTS is resolved against the new file's owner. The protected DACL excludes
        // broad inheritable directory permissions before any output bytes are written.
        private const string OwnerOnlyDacl = "D:P(A;;GA;;;OW)(A;;GA;;;SY)(A;;GA;;;BA)";

        internal static FileStream CreateWindowsOwnerOnly(string path, int bufferSize, FileOptions options) {
            if (!ConvertStringSecurityDescriptorToSecurityDescriptor(OwnerOnlyDacl, 1, out IntPtr descriptor, out _))
                throw new IOException("Unable to prepare owner-only Windows staging permissions (OS error " + Marshal.GetLastWin32Error() + ").");
            try {
                var attributes = new SecurityAttributes {
                    Length = Marshal.SizeOf<SecurityAttributes>(),
                    SecurityDescriptor = descriptor,
                    InheritHandle = false
                };
                const uint genericReadWrite = 0xC0000000U;
                const uint deleteAccess = 0x00010000U;
                const uint normalAttribute = 0x00000080U;
                uint access = genericReadWrite | ((options & FileOptions.DeleteOnClose) != 0 ? deleteAccess : 0);
                uint flags = normalAttribute | (unchecked((uint)options) & ~unchecked((uint)FileOptions.Asynchronous));
                SafeFileHandle handle = CreateFile(path, access, 0, ref attributes, 1, flags, IntPtr.Zero);
                if (handle.IsInvalid) {
                    int error = Marshal.GetLastWin32Error();
                    handle.Dispose();
                    throw new IOException("Unable to create owner-only Windows staging file (OS error " + error + ").");
                }
                try {
                    // A synchronous handle supports FileStream's asynchronous fallback without exposing the file.
                    return new FileStream(handle, FileAccess.ReadWrite, bufferSize, isAsync: false);
                } catch {
                    handle.Dispose();
                    throw;
                }
            } finally {
                LocalFree(descriptor);
            }
        }

        [StructLayout(LayoutKind.Sequential)]
        private struct SecurityAttributes {
            public int Length;
            public IntPtr SecurityDescriptor;
            [MarshalAs(UnmanagedType.Bool)] public bool InheritHandle;
        }

        [DllImport("advapi32.dll", EntryPoint = "ConvertStringSecurityDescriptorToSecurityDescriptorW", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool ConvertStringSecurityDescriptorToSecurityDescriptor(
            string securityDescriptor, uint revision, out IntPtr descriptor, out uint descriptorSize);

        [DllImport("kernel32.dll", EntryPoint = "CreateFileW", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern SafeFileHandle CreateFile(string path, uint access, uint share, ref SecurityAttributes attributes,
            uint creationDisposition, uint flags, IntPtr templateFile);

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern IntPtr LocalFree(IntPtr memory);
    }
}
